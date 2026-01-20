/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.service;

import com.pivotal.stockticker.model.ExchangeRate;
import com.pivotal.stockticker.model.SettingsManager;
import com.pivotal.stockticker.model.SymbolTransaction;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import java.util.*;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.ScheduledFuture;
import java.util.concurrent.TimeUnit;
import java.util.prefs.Preferences;

/**
 * Manages exchange rates including loading from storage and periodic updates
 * All rates are stored in user preferences so that they persist between application runs
 * Rates are identified by their currency code (e.g. "USD", "EUR") and are the multiplier
 * to convert from that currency to the base currency defined in application settings @see SettingsManager.getCurrencyCode()
 * The rates are updated periodically based on the application settings
 *
 * @see ExchangeRate
 */
@Slf4j
public class ExchangeRatesManager {

    private static final String EXCHANGE_RATES_ROOT = PersistanceManager.ROOT_NODE + ExchangeRate.class.getSimpleName();
    private Preferences prefs = Preferences.userRoot().node(EXCHANGE_RATES_ROOT);

    private final Map<String, ExchangeRate> currentRates = new TreeMap<>(String::compareToIgnoreCase);
    private final UpdateTask scheduler;
    private final SettingsManager settings;

    /**
     * Constructor - loads all exchange rates from persistent storage and starts the periodic update task
     *
     * @param settings Application settings
     */
    public ExchangeRatesManager(SettingsManager settings) {
        this.settings = settings;

        // Load all the saved prices values from persistent storage
        loadFromStorage();

        // Schedule the task to run every X seconds with an initial delay of 0 seconds
        scheduler = new UpdateTask(settings, currentRates);
        scheduler.start(settings.getExchangeRateFrequency());
    }

    /**
     * Load all symbols from persistent storage into memory
     */
    public void loadFromStorage() {
        // Load all the prices from the persistent storage
        prefs = Preferences.userRoot().node(EXCHANGE_RATES_ROOT);
        try {
            for (String code : prefs.childrenNames()) {
                currentRates.put(code, ExchangeRate.getExchangeRate(code));
            }
        }
        catch (Exception e) {
            log.error("Error accessing storage: {}", e.getMessage());
        }
        log.info("Loaded {} exchange rates", currentRates.size());
    }

    /**
     * Replaces all the current exchange rates with new ones
     *
     * @param codes Collection of exchange rate codes
     */
    public void replaceExchangeRates(Collection<String> codes) {
        if (codes == null || codes.isEmpty()) {
            return;
        }
        // Add or update prices from the new list
        for (String symbol : codes) {
            addRate(symbol);
        }

        // Find all the redundant prices that are no longer needed
        List<String> codesToDelete = new ArrayList<>();
        for (String code : currentRates.keySet()) {
            if (!codes.contains(code)) {
                codesToDelete.add(code);
            }
        }

        // Remove all prices that are no longer needed
        for (String code : codesToDelete) {
            currentRates.remove(code);
            Preferences pref = prefs.node(code);
            if (pref != null) {
                try {
                    pref.removeNode();
                }
                catch (Exception e) {
                    log.error("Error removing unused price {} from storage: {}", code, e.getMessage());
                }
            }
            log.debug("Removed unused price {}", code);
        }
    }

    /**
     * Create a new Exchange Rate
     *
     * @param code Symbol code
     * @return Newly created ExchangeRate with defaults
     */
    public ExchangeRate addRate(String code) {
        if (currentRates.containsKey(code)) {
            return currentRates.get(code);
        }
        try {
            ExchangeRate rate = ExchangeRate.getExchangeRate(code);
            currentRates.put(code, rate);
            log.debug("Added {} rate", code);
            return rate;
        }
        catch (Exception e) {
            log.error("Error creating new rate: {}", e.getMessage());
            return null;
        }
    }

    /**
     * Retrieve the exchange rate for a given source code
     *
     * @param code The stock code
     * @return The ExchangeRate object for the given code, or null if not found
     */
    public ExchangeRate getRate(String code) {
        return currentRates.get(code);
    }

    /**
     * Get all current exchange rates
     *
     * @return A map of stock codes to their corresponding Price objects
     */
    public Map<String, ExchangeRate> getAllExchangeRates() {
        return currentRates;
    }

    /**
     * Convert an amount from one currency to the base currency defined in settings
     *
     * @param symbolTransaction The symbol transaction containing the from currency
     * @param amount            The amount in the from currency
     * @return The equivalent amount in the base currency
     */
    public double convertAmount(SymbolTransaction symbolTransaction, double amount) {
        return convertAmount(symbolTransaction.getCurrencyCode(), symbolTransaction.getCurrencySymbol(), amount);
    }

    /**
     * Convert an amount from one currency to the base currency defined in settings
     *
     * @param fromCode           The currency code to convert from
     * @param fromCurrencySymbol The currency symbol to convert from
     * @param amount             The amount in the from currency
     * @return The equivalent amount in the base currency
     */
    public double convertAmount(String fromCode, String fromCurrencySymbol, double amount) {
        ExchangeRate fromRate = getRate(fromCode);
        if (fromRate == null) {
            log.warn("Cannot convert amount - missing exchange rate for {}", fromCode);
            return 0.0;
        }
        double total = amount * fromRate.getExchangeRate();

        // We need to take the currency symbol into account if it indicates a different denomination
        if (fromCurrencySymbol != null && fromCurrencySymbol.matches("[a-zA-Z¢]+")) {
            total = total / 100.0;
        }

        // Now we need to do a similar adjustment for the base currency symbol
        String baseCurrencySymbol = settings.getCurrencySymbol();
        if (baseCurrencySymbol != null && baseCurrencySymbol.matches("[a-zA-Z¢]+")) {
            total = total * 100.0;
        }
        return total;
    }

    /**
     * Update settings from the application
     */
    public void resetScheduler() {
        if (settings.getExchangeRateFrequency() != scheduler.getPeriodSeconds()) {
            scheduler.start(settings.getExchangeRateFrequency());
        }
    }

    /**
     * Scheduler to update exchange rates periodically
     */
    @Slf4j
    private static class UpdateTask {

        private final ScheduledExecutorService scheduler = Executors.newSingleThreadScheduledExecutor();
        private ScheduledFuture<?> scheduledFuture;
        private final SettingsManager settings;
        private Map<String, ExchangeRate> currentExchangeRates = new HashMap<>();
        @Getter
        private int periodSeconds;

        /**
         * Constructor
         *
         * @param settings             Application settings
         * @param currentExchangeRates Map of current prices
         */
        public UpdateTask(SettingsManager settings, Map<String, ExchangeRate> currentExchangeRates) {
            this.settings = settings;
            this.currentExchangeRates = currentExchangeRates;
        }

        /**
         * The periodic task to update exchange rates
         */
        private final Runnable task = () -> {

            // Get the list of stock symbols to update rates for
            log.debug("Updating exchange rates for {} rates", currentExchangeRates.size());
            Map<String, ExchangeRate> rates = new HashMap<>(currentExchangeRates);

            // Update prices for each symbol from each source
            // TODO - Implement actual exchange rate fetching logic here
        };

        /**
         * Start or restart the periodic task with a new frequency
         *
         * @param periodSeconds Frequency in seconds
         */
        public void start(int periodSeconds) {
            this.periodSeconds = periodSeconds;
            if (scheduledFuture != null && !scheduledFuture.isCancelled()) {
                log.debug("Stopping current price currency update task to re-schedule");
                scheduledFuture.cancel(false);
            }
            log.debug("Starting price currency updates - scheduling every {} seconds", periodSeconds);
            scheduledFuture = scheduler.scheduleAtFixedRate(task, 0, periodSeconds, TimeUnit.SECONDS);
        }

        /**
         * Stop the periodic task
         */
        public void stop() {
            if (scheduledFuture != null) {
                scheduledFuture.cancel(false);
            }
        }

        /**
         * Shutdown the scheduler
         */
        public void shutdownScheduler() {
            scheduler.shutdown();
        }
    }
}
