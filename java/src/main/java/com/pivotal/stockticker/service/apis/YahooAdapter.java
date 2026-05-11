package com.pivotal.stockticker.service.apis;

import com.jayway.jsonpath.Configuration;
import com.jayway.jsonpath.DocumentContext;
import com.jayway.jsonpath.JsonPath;
import com.jayway.jsonpath.Option;
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

/**
 * Client for Yahoo API that provides live prices for stocks
 *
 * @see <a href="https://query2.finance.yahoo.com">...</a>
 *
 */
@Slf4j
public class YahooAdapter implements PricesApiAdapter {

    private static final String AGENT_NAME = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36";
    private static final String BASE_URL = "https://query2.finance.yahoo.com/ws/fundamentals-timeseries/v6/finance/quoteSummary/%s?modules=price";
    private final HttpClient client = HttpClient.newBuilder().connectTimeout(Duration.ofSeconds(10)).build();
    private final Configuration conf = Configuration.builder().options(Option.DEFAULT_PATH_LEAF_TO_NULL).build();
    private final SettingsManager settingsManager;
    private final PricesManager pricesManager;

    /**
     * Constructor
     *
     * @param pricesManager Prices manager
     */
    public YahooAdapter(PricesManager pricesManager) {
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

        // Retrieve each symbol's price and intraday data
        List<String> updatedSymbols = new ArrayList<>();
        for (String symbol : symbols) {

            // Build request to fetch price for symbol
            log.debug("Fetching price for {} from {} API...", symbol, getAdapterName());
            String adjustedSymbol = symbol.trim().replace('^', '.');
            HttpRequest request = HttpRequest.newBuilder()
                    .uri(URI.create(String.format(BASE_URL, adjustedSymbol)))
                    .GET()
                    .timeout(REQUEST_TIMEOUT)
                    .header("Accept", "application/json")
                    .header("User-Agent", AGENT_NAME)
                    .build();
            try {
                HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());
                if (response.statusCode() != 200) {
                    log.error("Failed to fetch value for {}: HTTP [{}] {}", symbol, response.statusCode(), response.body());
                }
                else {

                    // Got a price
                    log.debug("Successfully fetched price for {} from {} API", symbol, getAdapterName());
                    if (response.body().equals("[]")) {
                        log.debug("Error fetching price for {}: Nothing returned", symbol);
                        continue;
                    }
                    DocumentContext context = JsonPath.using(conf).parse(response.body());

                    // Convert data to Price and update manager
                    Price price = pricesManager.getPrice(symbol);
                    try {
                        if (price == null) {
                            price = pricesManager.addPrice(symbol);
                        }

                        // Update price details
                        price.setDayClose(getValue(context.read("$.quoteSummary.result[0].price.regularMarketPreviousClose.fmt"), price.getDayClose()));
                        price.setDayHigh(getValue(context.read("$.quoteSummary.result[0].price.regularMarketDayHigh.fmt"), price.getDayHigh()));
                        price.setDayLow(getValue(context.read("$.quoteSummary.result[0].price.regularMarketDayLow.fmt"), price.getDayLow()));
                        price.setCurrentPrice(getValue(context.read("$.quoteSummary.result[0].price.regularMarketPrice.fmt"), price.getCurrentPrice()));
                        price.setLastUpdate(LocalDateTime.now());
                        price.setSource(getAdapterName());
                        updatedSymbols.add(symbol);
                    }
                    catch (Exception e) {
                        log.error("Failed to decode price from JSON {}", symbol, e);
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
        return returnVal;
    }
}
