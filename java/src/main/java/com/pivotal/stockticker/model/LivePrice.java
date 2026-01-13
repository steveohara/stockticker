/*
 *
 * Copyright (c) 2026, 4NG and/or its affiliates. All rights reserved.
 * 4NG PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.model;

import com.pivotal.stockticker.Utils;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.awt.*;

/**
 *
 */
@Getter
@Setter
@Slf4j
public class LivePrice {

    private final PricesManager prices;
    private final ExchangeRatesManager exchangeRates;
    private final SymbolsManager symbols;
    private final Settings settings;
    private final String symbol;
    private final boolean aggregated;
    private SymbolTransaction symbolTransaction;
    private Rectangle bounds = null;

    /**
     * Constructor for live price for a specific transaction
     *
     * @param prices            Prices manager
     * @param exchangeRates     Exchange rates manager
     * @param settings          Application settings
     * @param symbolTransaction Specific symbol transaction
     */
    public LivePrice(SymbolsManager symbols, PricesManager prices, ExchangeRatesManager exchangeRates, Settings settings, SymbolTransaction symbolTransaction) {
        this.symbols = symbols;
        this.prices = prices;
        this.exchangeRates = exchangeRates;
        this.settings = settings;
        this.symbolTransaction = symbolTransaction;
        this.symbol = symbolTransaction.getCode();
        aggregated = false;
    }

    /**
     * Constructor for aggregated live price across all transactions for a symbol
     *
     * @param prices        Prices manager
     * @param exchangeRates Exchange rates manager
     * @param settings      Application settings
     * @param symbol        Stock symbol
     */
    public LivePrice(SymbolsManager symbols, PricesManager prices, ExchangeRatesManager exchangeRates, Settings settings, String symbol) {
        this.symbols = symbols;
        this.prices = prices;
        this.exchangeRates = exchangeRates;
        this.settings = settings;
        this.symbol = symbol;
        aggregated = true;
    }

    /**
     * Sets the bounds for this live price
     *
     * @param x X coordinate
     * @param y Y coordinate
     */
    public boolean isInBounds(int x, int y) {
        return bounds != null && bounds.contains(x, y);
    }

    /**
     * Calculates the aggregated price paid across all non-disabled transactions for the current symbol
     *
     * @return Aggregated price paid (cost base)
     */
    public double getAggregatedPricePaid() {
        double totalPaid = 0;
        double totalShares = 0;
        for (SymbolTransaction transaction : symbols.getSymbolTransactions(symbol, false)) {
            totalPaid += transaction.getPricePaid() * transaction.getSharesBought();
            totalShares += transaction.getSharesBought();
        }
        return totalShares == 0 ? 0 : totalPaid / totalShares;
    }

    /**
     * Calculates the aggregated shares bought
     *
     * @return Aggregated price paid (cost base)
     */
    public double getAggregatedSharesBought() {
        double totalShares = 0;
        for (SymbolTransaction transaction : symbols.getSymbolTransactions(symbol, false)) {
            totalShares += transaction.getSharesBought();
        }
        return totalShares;
    }

    /**
     * Calculates the percentage change from the original price to the current price.
     *
     * @return Percentage change.
     */
    public double getPercentChange() {
        double pricePaid = aggregated ? getAggregatedPricePaid() : symbolTransaction.getPricePaid();
        double currentPrice = prices.getPrice(symbol).getCurrentPrice();
        if (pricePaid == 0) {
            return 0;
        }
        return ((currentPrice - pricePaid) * 100) / pricePaid;
    }

    /**
     * Returns the formatted percentage change as a string with two decimal places.
     *
     * @return Formatted percentage change.
     */
    public String getFormattedPercentChange() {
        return String.format("%.2f%%", getPercentChange());
    }

    /**
     * Returns the formatted current price as a currency string.
     *
     * @return Formatted current price.
     */
    public String getFormattedValue() {
        double currentPrice = prices.getPrice(symbol).getCurrentPrice();
        String currencyCode = symbols.getFirst(symbol).getCurrencyCode();
        return Utils.formatCurrencyValue(currentPrice, currencyCode);
    }

    /**
     * Returns the formatted total value (current price * shares) as a currency string.
     *
     * @return Formatted total value.
     */
    public String getFormattedTotalValue() {
        double currentPrice = prices.getPrice(symbol).getCurrentPrice();
        double sharesBought = aggregated ? getAggregatedSharesBought() : symbolTransaction.getSharesBought();
        String currencyCode = symbols.getFirst(symbol).getCurrencyCode();
        return Utils.formatCurrencyValue(currentPrice * sharesBought, currencyCode);
    }

    /**
     * Calculates the profit or loss based on current price and original price.
     *
     * @return Profit or loss amount.
     */
    public double getProfitLoss() {
        double pricePaid = aggregated ? getAggregatedPricePaid() : symbolTransaction.getPricePaid();
        double currentPrice = prices.getPrice(symbol).getCurrentPrice();
        double sharesBought = aggregated ? getAggregatedSharesBought() : symbolTransaction.getSharesBought();
        return (currentPrice - pricePaid) * sharesBought;
    }

    /**
     * Returns the formatted profit or loss as a currency string.
     *
     * @return Formatted profit or loss.
     */
    public String getFormattedProfitLoss() {
        String currencyCode = symbols.getFirst(symbol).getCurrencyCode();
        return Utils.formatCurrencyValue(getProfitLoss(), currencyCode);
    }

}
