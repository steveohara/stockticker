package com.pivotal.stockticker.service.apis;

import com.jayway.jsonpath.JsonPath;
import com.pivotal.stockticker.model.Price;
import com.pivotal.stockticker.model.SettingsManager;
import com.pivotal.stockticker.service.PricesManager;
import lombok.extern.slf4j.Slf4j;

import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.time.Duration;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * Client for MarketStack API that provides live prices for stocks
 *
 * @see <a href="https://www.MarketStack.co/documentation//">...</a>
 *
 */
@Slf4j
public class MarketStackAdapter implements PricesApiAdapter {

    private static final String BASE_URL = "http://api.marketstack.com/v1/intraday/latest?access_key=%s&symbols=%s";
    private final HttpClient client = HttpClient.newBuilder().connectTimeout(Duration.ofSeconds(10)).build();

    private final SettingsManager settingsManager;
    private final PricesManager pricesManager;

    /**
     * Constructor
     *
     * @param settingsManager Application settings manager
     * @param pricesManager Prices manager
     */
    public MarketStackAdapter(SettingsManager settingsManager, PricesManager pricesManager) {
        this.settingsManager = settingsManager;
        this.pricesManager = pricesManager;
    }

    @Override
    public Collection<String> fetchAndUpdatePrices(Collection<String> symbols) {
        List<String> returnVal = new ArrayList<>(symbols);

        // If there are no symbols, return immediately
        if (symbols.isEmpty()) {
            log.debug("No symbols provided to fetch prices from {} API", getAdapterName());
            return returnVal;
        }

        // Check if API key is set
        String apiKey = settingsManager.getMarketStackToken();
        if (apiKey == null || apiKey.isEmpty()) {
            log.debug("{} API key is not set. Cannot fetch prices", getAdapterName());
            return returnVal;
        }
        else {
            // Retrieve each symbol's price and intraday data
            List<String> updatedSymbols = new ArrayList<>();

            // Build request to fetch price for symbol
            String symbolsList = String.join(",", symbols);
            log.debug("Fetching price for {} from {} API...", symbolsList, getAdapterName());
            String adjustedSymbol = symbolsList.trim().replace('^', '.');
            HttpRequest request = HttpRequest.newBuilder()
                    .uri(URI.create(String.format(BASE_URL, apiKey, adjustedSymbol)))
                    .GET()
                    .header("Accept", "application/json")
                    .build();

            try {
                HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());
                if (response.statusCode() != 200) {
                    log.error("Failed to fetch price for {} : HTTP [{}] {}", symbols, response.statusCode(), response.body());
                }
                else {

                    // Got some rates
                    log.debug("Successfully fetched prices for {} from {} API", symbolsList, getAdapterName());
                    List<Map<String, Object>> pricesData = JsonPath.read(response.body(), "$.data");

                    // Loop through the rates and update the ExchangeRatesManager
                    for (Map<String, Object> priceData : pricesData) {

                        // Get the symbol
                        String symbol = (String)priceData.get("symbol");

                        // Convert data to Price and update manager
                        Price price = pricesManager.getPrice(symbol);
                        try {
                            if (price == null) {
                                price = pricesManager.addPrice(symbol);
                            }

                            // Update rate details
                            price.setDayClose(getValue(priceData.get("close"), price.getDayClose()));
                            price.setDayStart(getValue(priceData.get("open"), price.getDayStart()));
                            price.setDayHigh(getValue(priceData.get("high"), price.getDayHigh()));
                            price.setDayLow(getValue(priceData.get("low"), price.getDayLow()));
                            price.setCurrentPrice(getValue(priceData.get("last"), price.getCurrentPrice()));
                            price.setLastUpdate(LocalDateTime.now());
                            price.setSource(getAdapterName());
                            updatedSymbols.add(symbol);
                        }
                        catch (Exception e) {
                            log.error("Failed to decode price from JSON {} - {}", symbol, e.getMessage());
                        }
                    }
                }
            }
            catch (Exception e) {
                log.error("Error fetching exchange rates from {} API: {}", getAdapterName(), e.getMessage());
            }
            log.info("Fetched {} prices from {} API", updatedSymbols.isEmpty() ? "0" : String.join(",", updatedSymbols), getAdapterName());
            returnVal.removeAll(updatedSymbols);
        }
        return returnVal;
    }
}
