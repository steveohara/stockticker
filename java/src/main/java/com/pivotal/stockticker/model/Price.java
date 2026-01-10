package com.pivotal.stockticker.model;

import lombok.AccessLevel;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.time.LocalDateTime;

/**
 * Represents the current price of a stock symbol
 */
@Slf4j
@Getter
@Setter
public class Price extends PersistanceManager {
    @Setter(AccessLevel.NONE)
    private String code = "YAHOO";
    private double currentPrice;
    private double dayStart;
    private double dayChange;
    private double dayHigh;
    private double dayLow;
    private String errorDescription;
    private LocalDateTime lastPriceUpdate;
    private String priceSource;

    /**
     * Sets the stock code, ensuring it is trimmed and uppercase.
     *
     * @param code Stock code.
     */
    public void setCode(String code) {
        this.code = code != null ? code.trim().toUpperCase() : null;
    }

}
