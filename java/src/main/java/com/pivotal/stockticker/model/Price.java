/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.model;

import com.pivotal.stockticker.service.PersistanceManager;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.time.LocalDateTime;
import java.util.prefs.Preferences;

/**
 * Represents the current price of a stock symbol
 */
@Slf4j
@Getter
@Setter
public class Price extends PersistanceManager {
    private String code;
    private double currentPrice;
    private double dayStart;
    private double dayHigh;
    private double dayLow;
    private String errorDescription;
    private LocalDateTime lastUpdate;
    private String source;

    /**
     * Creates a proxy instance of this class loaded from persistent storage.
     *
     * @param code Symbol code.
     * @return A proxy instance of this class.
     * @throws Exception if proxy creation fails.
     */
    public static Price getPrice(String code) throws Exception {
        Price price = createProxyInstance(Price.class, Preferences.userRoot().node(ROOT_NODE + Price.class.getSimpleName() + '/' + code), true);
        price.setCode(code);
        return price;
    }

    /**
     * Get the price change since the start of the day
     *
     * @return Price change since the start of the day
     */
    public double getDayChange() {
        return currentPrice - dayStart;
    }

}
