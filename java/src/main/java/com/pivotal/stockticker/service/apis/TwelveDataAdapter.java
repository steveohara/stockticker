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
 * Client for TwelveData API that provides live prices for stocks
 *
 * @see <a href="https://www.TwelveData.co/documentation//">...</a>
 *
 */
@Slf4j
public class TwelveDataAdapter implements PricesApiAdapter {

    private static final String BASE_URL = "https://api.twelvedata.com/quote?apikey=%s&symbol=%s";
    private static final String BASE_URL_PRICE = "https://api.twelvedata.com/price?apikey=%s&symbol=%s";
    private final HttpClient client = HttpClient.newBuilder().connectTimeout(Duration.ofSeconds(10)).build();

    private final SettingsManager settingsManager;
    private final PricesManager pricesManager;

    /**
     * Constructor
     *
     * @param pricesManager Prices manager
     */
    public TwelveDataAdapter(PricesManager pricesManager) {
        this.settingsManager = SettingsManager.getInstance();
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
        String apiKey = settingsManager.getTwelveDataToken();
        if (apiKey == null || apiKey.isEmpty()) {
            log.debug("{} API key is not set. Cannot fetch prices", getAdapterName());
            return returnVal;
        }
        else {
            // Retrieve each symbol's price and intraday data
            List<String> updatedSymbols = new ArrayList<>();
            for (String symbol : symbols) {

                // Build request to fetch price for symbol
                log.debug("Fetching price for {} from {} API...", symbol, getAdapterName());
                String adjustedSymbol = symbol.trim().replaceAll("(?i)[.]L", "");
                HttpRequest request = HttpRequest.newBuilder()
                        .uri(URI.create(String.format(BASE_URL, apiKey, adjustedSymbol)))
                        .GET().header("Accept", "application/json").build();
                try {
                    HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());
                    if (response.statusCode() == 429) {
                        log.debug("Too many requests - cannot fetch value for {}: HTTP [{}] {}", symbol, response.statusCode(), response.body());
                    }
                    else if (response.statusCode() != 200) {
                        log.error("Failed to fetch day values for {} : HTTP [{}] {}", symbol, response.statusCode(), response.body());
                    }
                    else {

                        // Got a price
                        log.debug("Successfully fetched price for {} from {} API", symbol, getAdapterName());
                        Map<String, Object> priceData = JsonPath.read(response.body(), "$");
                        if (priceData.containsKey("status")) {
                            log.debug("Error fetching price for {}: {}", symbol, priceData.get("message"));
                            continue;
                        }

                        // Convert data to Price and update manager
                        Price price = pricesManager.getPrice(symbol);
                        try {
                            if (price == null) {
                                price = pricesManager.addPrice(symbol);
                            }

                            // Update price details
                            price.setDayClose(getValue(priceData.get("close"), price.getDayClose()));
                            price.setDayStart(getValue(priceData.get("open"), price.getDayStart()));
                            price.setDayHigh(getValue(priceData.get("high"), price.getDayHigh()));
                            price.setDayLow(getValue(priceData.get("low"), price.getDayLow()));
                            price.setLastUpdate(LocalDateTime.now());
                            price.setSource(getAdapterName());

                            // Have to get the price separately
                            request = HttpRequest.newBuilder()
                                    .uri(URI.create(String.format(BASE_URL_PRICE, apiKey, adjustedSymbol)))
                                    .GET().header("Accept", "application/json").build();
                            response = client.send(request, HttpResponse.BodyHandlers.ofString());
                            if (response.statusCode() != 200) {
                                log.error("Failed to fetch price for {} : HTTP [{}] {}", symbol, response.statusCode(), response.body());
                            }
                            else {
                                price.setCurrentPrice(getValue(JsonPath.read(response.body(), "$.price"), price.getDayLow()));
                            }
                            updatedSymbols.add(symbol);
                        }
                        catch (Exception e) {
                            log.error("Failed to decode price from JSON {} - {}", symbol, e.getMessage());
                        }
                    }
                }
                catch (Exception e) {
                    log.error("Error fetching prices from ▲{} - {}", getAdapterName(), e.getMessage());
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
