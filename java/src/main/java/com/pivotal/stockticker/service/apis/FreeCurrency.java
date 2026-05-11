package com.pivotal.stockticker.service.apis;

import com.jayway.jsonpath.JsonPath;
import com.pivotal.stockticker.model.ExchangeRate;
import com.pivotal.stockticker.model.SettingsManager;
import com.pivotal.stockticker.service.ExchangeRatesManager;
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
 * Client for FreeCurrency API that provides exchange rates for currencies
 *
 * @see <a href="https://freecurrencyapi.net/">...</a>
 *
 */
@Slf4j
public class FreeCurrency {

    private static final String BASE_URL = "https://api.freecurrencyapi.com/v1/latest";
    private static final Duration REQUEST_TIMEOUT = Duration.ofSeconds(30);
    private final HttpClient client = HttpClient.newBuilder().connectTimeout(Duration.ofSeconds(10)).build();

    private final SettingsManager settingsManager;
    private final ExchangeRatesManager exchangeRatesManager;

    /**
     * Constructor
     *
     * @param exchangeRatesManager Exchange rates manager
     */
    public FreeCurrency(ExchangeRatesManager exchangeRatesManager) {
        this.settingsManager = SettingsManager.getInstance();
        this.exchangeRatesManager = exchangeRatesManager;
    }

    /**
     * Fetch exchange rates from FreeCurrency API and update the ExchangeRatesManager
     *
     * @param currencyCodes List of currency codes to fetch rates for
     * @return List of symbols that were not successfully updated
     */
    public Collection<String> fetchAndUpdateExchangeRates(Collection<String> currencyCodes) {
        List<String> returnVal = new ArrayList<>(currencyCodes);

        // If there are no symbols, return immediately
        if (currencyCodes.isEmpty()) {
            log.debug("No symbols provided to fetch prices from AlphaVantage API");
            return returnVal;
        }

        // Check if API key is set
        String apiKey = settingsManager.getFreeCurrencyToken();
        if (apiKey == null || apiKey.isEmpty()) {
            log.warn("FreeCurrency API key is not set. Cannot fetch exchange rates.");
            return returnVal;
        }
        else {
            log.debug("Fetching exchange rates from FreeCurrency API...");
            HttpRequest request = HttpRequest.newBuilder()
                    .uri(URI.create(BASE_URL + String.format("?apikey=%s&base_currency=%s&currencies=%s",
                            apiKey, settingsManager.getCurrencyCode(), String.join(",", currencyCodes))))
                    .timeout(REQUEST_TIMEOUT)
                    .GET()
                    .header("Accept", "application/json")
                    .build();

            try {
                HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());
                if (response.statusCode() != 200) {
                    log.error("Failed to fetch exchange rates: HTTP [{}] {}", response.statusCode(), response.body());
                }
                else {

                    // Got some rates
                    log.debug("Successfully fetched exchange rates from FreeCurrency API");
                    Map<String, Object> rates = JsonPath.read(response.body(), "$.data");

                    // Loop through the rates and update the ExchangeRatesManager
                    List<String> updatedSymbols = new ArrayList<>();
                    rates.forEach((currency, data) -> {

                        // Convert data to ExchangeRate and update manager
                        ExchangeRate rate = exchangeRatesManager.getRate(currency);
                        try {
                            if (rate == null) {
                                rate = ExchangeRate.getExchangeRate(currency);
                            }

                            // Update rate details
                            rate.setExchangeRate(((Number) data).doubleValue());
                            rate.setLastUpdate(LocalDateTime.now());
                            rate.setSource(this.getClass().getSimpleName());
                            updatedSymbols.add(currency);
                        }
                        catch (Exception e) {
                            log.error("Failed to decode exchange rate from JSON {}", currency, e);
                        }
                    });
                    log.info("Successfully fetched {} exchange rates from FreeCurrency API", String.join(",", updatedSymbols));
                    returnVal.removeAll(updatedSymbols);
                }
            }
            catch (Exception e) {
                log.error("Error fetching exchange rates from FreeCurrency API: {}", e.getMessage());
            }
        }
        return returnVal;
    }
}
