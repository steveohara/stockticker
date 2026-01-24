/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.model;

import com.pivotal.stockticker.Utils;
import com.pivotal.stockticker.service.ExchangeRatesManager;
import com.pivotal.stockticker.service.PricesManager;
import com.pivotal.stockticker.service.SymbolsManager;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.awt.*;
import java.util.ArrayList;

/**
 * Represents the live price information for a stock symbol, either for a specific transaction or aggregated across all transactions
 */
@Getter
@Setter
@Slf4j
public class LivePrice {

    private final PricesManager prices;
    private final ExchangeRatesManager exchangeRates;
    private final SymbolsManager symbols;
    private final SettingsManager settings;
    private final String symbol;
    private final boolean aggregated;
    private SymbolTransaction symbolTransaction;
    private Rectangle bounds = null;

    /**
     * Constructor for live price for a specific transaction
     *
     * @param symbols         Symbols manager
     * @param prices            Prices manager
     * @param exchangeRates     Exchange rates manager
     * @param symbolTransaction Specific symbol transaction
     * @param aggregated        Whether to aggregate across all transactions for the symbol
     */
    public LivePrice(SymbolsManager symbols, PricesManager prices, ExchangeRatesManager exchangeRates, SymbolTransaction symbolTransaction, boolean aggregated) {
        this.symbols = symbols;
        this.prices = prices;
        this.exchangeRates = exchangeRates;
        this.settings = SettingsManager.getInstance();
        this.symbolTransaction = symbolTransaction;
        this.symbol = symbolTransaction.getCode();
        this.aggregated = aggregated;
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
    private double getAggregatedPricePaid() {
        double totalPaid = 0;
        double totalShares = 0;
        for (SymbolTransaction transaction : symbols.getSymbolTransactions(false, false, symbol)) {
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
    private double getAggregatedSharesBought() {
        double totalShares = 0;
        for (SymbolTransaction transaction : symbols.getSymbolTransactions(false, false, symbol)) {
            totalShares += transaction.getSharesBought();
        }
        return totalShares;
    }

    /**
     * Calculates the difference between the current price and the original price.
     *
     * @return Difference amount.
     */
    public double getChange() {
        double pricePaid = aggregated ? getAggregatedPricePaid() : symbolTransaction.getPricePaid();
        double currentPrice = prices.getPrice(symbol).getCurrentPrice();
        return (currentPrice - pricePaid);
    }

    /**
     * Calculates the difference between the current price and the day starting price.
     *
     * @return Difference amount.
     */
    public double getDayChange() {
        Price price = prices.getPrice(symbol);
        return (price.getCurrentPrice() - (price.getDayClose() == 0 ? price.getDayStart() : price.getDayClose()));
    }

    /**
     * Calculates the difference between the current price and the original price.
     *
     * @return Difference amount.
     */
    public String getFormattedChange() {
        return Utils.formatCurrencyValue(getChange(), symbolTransaction.getCurrencySymbol());
    }

    /**
     * Calculates the difference between the current price and the day starting price.
     *
     * @return Difference amount.
     */
    public String getFormattedDayChange() {
        return Utils.formatCurrencyValue(getDayChange(), symbolTransaction.getCurrencySymbol());
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
     * Calculates the percentage change from the day starting price to the current price.
     *
     * @return Percentage change.
     */
    public double getPercentDayChange() {
        Price price = prices.getPrice(symbol);
        double start = price.getDayClose() == 0 ? price.getDayStart() : price.getDayClose();
        if (start == 0) {
            return 0;
        }
        return ((price.getCurrentPrice() - start) * 100) / start;
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
     * Returns the formatted percentage day change as a string with two decimal places.
     *
     * @return Formatted percentage change.
     */
    public String getFormattedPercentDayChange() {
        return String.format("%.2f%%", getPercentDayChange());
    }

    /**
     * Returns the formatted current price as a currency string.
     *
     * @return Formatted current price.
     */
    public String getFormattedPrice() {
        double currentPrice = prices.getPrice(symbol).getCurrentPrice();
        String currencySymbol = symbolTransaction.getCurrencySymbol();
        return Utils.formatCurrencyValue(currentPrice, currencySymbol);
    }

    /**
     * Returns the formatted total value (current price * shares) as a currency string.
     *
     * @return Formatted total value.
     */
    public String getFormattedValue() {
        double currentPrice = prices.getPrice(symbol).getCurrentPrice();
        double sharesBought = aggregated ? getAggregatedSharesBought() : symbolTransaction.getSharesBought();
        return Utils.formatCurrencyValue(exchangeRates.convertAmount(symbolTransaction, currentPrice * sharesBought), settings.getCurrencySymbol());
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
     * Calculates the profit or loss based on current price and day starting price.
     *
     * @return Profit or loss amount.
     */
    public double getDayProfitLoss() {
        Price price = prices.getPrice(symbol);
        double sharesBought = aggregated ? getAggregatedSharesBought() : symbolTransaction.getSharesBought();
        double start = price.getDayClose() == 0 ? price.getDayStart() : price.getDayClose();
        return (price.getCurrentPrice() - start) * sharesBought;
    }

    /**
     * Returns the formatted profit or loss as a currency string.
     *
     * @return Formatted profit or loss.
     */
    public String getFormattedProfitLoss() {
        return Utils.formatCurrencyValue(getProfitLoss(), symbolTransaction.getCurrencySymbol());
    }

    /**
     * Returns the formatted day profit or loss as a currency string.
     *
     * @return Formatted profit or loss.
     */
    public String getFormattedDayProfitLoss() {
        return Utils.formatCurrencyValue(getDayProfitLoss(), symbolTransaction.getCurrencySymbol());
    }

    /**
     * Determines if the current price is higher than the price paid.
     * @return True if the current price is higher, false otherwise.
     */
    public boolean isUp() {
        double pricePaid = aggregated ? getAggregatedPricePaid() : symbolTransaction.getPricePaid();
        double currentPrice = prices.getPrice(symbol).getCurrentPrice();
        return currentPrice > pricePaid;
    }

    /**
     * Determines if the current price is higher than the day starting price.
     * @return True if the current price is higher, false otherwise.
     */
    public boolean isUpToday() {
        Price price = prices.getPrice(symbol);
        return price.getCurrentPrice() > (price.getDayClose() == 0 ? price.getDayStart() : price.getDayClose());
    }

    /**
     * Determines if the current price is lower than the price paid.
     * @return True if the current price is lower, false otherwise.
     */
    public boolean isDown() {
        double pricePaid = aggregated ? getAggregatedPricePaid() : symbolTransaction.getPricePaid();
        double currentPrice = prices.getPrice(symbol).getCurrentPrice();
        return currentPrice < pricePaid;
    }

    /**
     * Determines if the current price is lower than the day starting price.
     * @return True if the current price is lower, false otherwise.
     */
    public boolean isDownToday() {
        Price price = prices.getPrice(symbol);
        return price.getCurrentPrice() < (price.getDayClose() == 0 ? price.getDayStart() : price.getDayClose());
    }

    /**
     * Generates a list of LivePrice objects based on the settings.
     *
     * @param symbols       Symbols manager
     * @param prices        Prices manager
     * @param exchangeRates Exchange rates manager
     * @param settings      Application settings
     * @return List of LivePrice objects
     */
    public static ArrayList<LivePrice> getLivePrices(SymbolsManager symbols, PricesManager prices, ExchangeRatesManager exchangeRates, SettingsManager settings) {
        ArrayList<LivePrice> livePrices = new ArrayList<>();
        boolean isAveraged = settings.isShowUniqueSymbols();
        for (SymbolTransaction symbolTransaction : symbols.getSymbolTransactions(false, isAveraged, null)) {
            livePrices.add(new LivePrice(symbols, prices, exchangeRates, symbolTransaction, isAveraged));
        }
        return livePrices;
    }
}
