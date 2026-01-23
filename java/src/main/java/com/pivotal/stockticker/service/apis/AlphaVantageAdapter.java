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
 * Client for AlphaVantage API that provides live prices for stocks
 *
 * @see <a href="https://www.alphavantage.co/documentation//">...</a>
 *
 */
@Slf4j
public class AlphaVantageAdapter implements PricesApiAdapter {

    private static final String BASE_URL = "https://www.alphavantage.co/query?function=GLOBAL_QUOTE&datatype=json&apikey=%s&symbol=%s";
    private final HttpClient client = HttpClient.newBuilder().connectTimeout(Duration.ofSeconds(10)).build();

    private final SettingsManager settingsManager;
    private final PricesManager pricesManager;

    /**
     * Constructor
     *
     * @param settingsManager Application settings manager
     * @param pricesManager Prices manager
     */
    public AlphaVantageAdapter(SettingsManager settingsManager, PricesManager pricesManager) {
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
        String apiKey = settingsManager.getAlphaVantageToken();
        if (apiKey == null || apiKey.isEmpty()) {
            log.debug("{} API key is not set. Cannot fetch prices", getAdapterName());
            return returnVal;
        }
        else {
            // Retrieve each symbol's price and intraday data
            List<String> updatedSymbols = new ArrayList<>();
            for (String symbol : symbols) {

                // Check for symbols that are not supported by AlphaVantage
                if (symbol.matches("(?i).+[.][a-z]+")) {
                    log.debug("Skipping symbol for price fetch from {} API: {}", getAdapterName(), symbol);
                    continue;
                }

                // Build request to fetch price for symbol
                log.debug("Fetching price for {} from {} API...", symbol, getAdapterName());
                String adjustedSymbol = symbol.trim().replace('^', '.');
                HttpRequest request = HttpRequest.newBuilder()
                        .uri(URI.create(String.format(BASE_URL, apiKey, adjustedSymbol)))
                        .GET()
                        .header("Accept", "application/json")
                        .build();

                try {
                    HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());
                    if (response.statusCode() == 429) {
                        log.debug("Too many requests - cannot fetch value for {}: HTTP [{}] {}", symbol, response.statusCode(), response.body());
                    }
                    else if (response.statusCode() != 200) {
                        log.error("Failed to fetch price for {} : HTTP [{}] {}", symbol, response.statusCode(), response.body());
                    }
                    else {

                        // Got some rates
                        log.debug("Successfully fetched price for {} from {} API", symbol, getAdapterName());
                        Map<String, Object> priceData = JsonPath.read(response.body(), "$");
                        if (priceData.containsKey("Information")) {
                            log.error("Error fetching price for {}: {}", symbol, priceData.get("Information"));
                            continue;
                        }

                        // Loop through the rates and update the ExchangeRatesManager
                        priceData = JsonPath.read(response.body(), "$['Global Quote']");

                        // Convert data to Price and update manager
                        Price price = pricesManager.getPrice(symbol);
                        try {
                            if (price == null) {
                                price = pricesManager.addPrice(symbol);
                            }

                            // Update rate details
                            price.setDayStart(getValue(priceData.get("02. open"), price.getDayStart()));
                            price.setDayHigh(getValue(priceData.get("03. high"), price.getDayHigh()));
                            price.setDayLow(getValue(priceData.get("04. low"), price.getDayLow()));
                            price.setCurrentPrice(getValue(priceData.get("05. price"), price.getCurrentPrice()));
                            price.setDayClose(getValue(priceData.get("08. previous close"), price.getDayClose()));
                            price.setLastUpdate(LocalDateTime.now());
                            price.setSource(getAdapterName());
                            updatedSymbols.add(symbol);
                        }
                        catch (Exception e) {
                            log.error("Failed to decode price from JSON {} - {}", symbol, e.getMessage());
                        }
                    }
                }
                catch (Exception e) {
                    log.error("Error fetching exchange rates from {} API: {}", e.getMessage(), getAdapterName());
                }
            }
            if (updatedSymbols.isEmpty()) {
                log.debug("Fetched 0 prices from {} API", getAdapterName());
            }
            else {
                log.info("Fetched {} prices from {} API", String.join(",", updatedSymbols), getAdapterName());
            }
            returnVal.removeAll(updatedSymbols);
        }
        return returnVal;
    }
}
