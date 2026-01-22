package com.pivotal.stockticker.service.apis;

import java.util.Collection;

/**
 * Adapter interface for price APIs
 */
public interface PricesApiAdapter {

    /**
     * Fetch prices for the given symbols and update the PricesManager
     *
     * @param symbols Collection of stock symbols to fetch prices for
     * @return Collection of symbols that were not successfully updated
     */
    Collection<String> fetchAndUpdatePrices(Collection<String> symbols);

}
