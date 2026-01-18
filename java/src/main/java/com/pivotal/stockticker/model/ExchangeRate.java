package com.pivotal.stockticker.model;

import com.pivotal.stockticker.service.PersistanceManager;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.time.LocalDateTime;
import java.util.prefs.Preferences;

/**
 * Represents the current exchange rate for a currency as compared
 * to the application base currency.
 */
@Slf4j
@Getter
@Setter
public class ExchangeRate extends PersistanceManager {
    private String sourceCurrencyCode;
    private double exchangeRate;
    private String errorDescription;
    private LocalDateTime lastUpdate;
    private String source;

    /**
     * Creates a proxy instance of this class loaded from persistent storage.
     *
     * @param sourceCurrencyCode The source currency code.
     * @return A proxy instance of this class.
     * @throws Exception if proxy creation fails.
     */
    public static ExchangeRate getExchangeRate(String sourceCurrencyCode) throws Exception {
        ExchangeRate rate = createProxyInstance(ExchangeRate.class, Preferences.userRoot().node(ROOT_NODE +  ExchangeRate.class.getSimpleName() + '/' + sourceCurrencyCode), true);
        rate.setSourceCurrencyCode(sourceCurrencyCode);
        return rate;
    }
}
