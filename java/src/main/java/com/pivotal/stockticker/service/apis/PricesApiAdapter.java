package com.pivotal.stockticker.service.apis;

import java.time.Duration;
import java.util.Collection;

/**
 * Adapter interface for price APIs
 */
public interface PricesApiAdapter {

    /**
     * Maximum time to wait for a response from any price API.
     * This is a request-level timeout (distinct from the TCP connect timeout) and
     * guards against servers that accept connections but never send a response.
     */
    Duration REQUEST_TIMEOUT = Duration.ofSeconds(30);

    /**
     * Fetch prices for the given symbols and update the PricesManager
     *
     * @param symbols Collection of stock symbols to fetch prices for
     * @return Collection of symbols that were not successfully updated
     */
    Collection<String> fetchAndUpdatePrices(Collection<String> symbols);

    /**
     * Get the name of the adapter (e.g., "AlphaVantage", "MarketStack")
     *
     * @return Adapter name
     */
    default String getAdapterName() {
        return this.getClass().getSimpleName().replace("Adapter", "");
    }

    /**
     * Utility method to extract a double value from an Object
     *
     * @param value        The object to extract the double from
     * @param defaultValue The default value to return if extraction fails
     * @return Extracted double value or defaultValue if extraction fails
     */
    default double getValue(Object value, double defaultValue) {
        if (value == null) {
            return defaultValue;
        }
        switch(value) {
            case Number n:
                return n.doubleValue();
            case String s:
                try {
                    return Double.parseDouble(s.replaceAll("[^0-9.]", ""));
                }
                catch (NumberFormatException e) {
                    return defaultValue;
                }
            default:
                return defaultValue;
        }
    }
}
