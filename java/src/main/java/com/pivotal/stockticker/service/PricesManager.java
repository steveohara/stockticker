/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.service;

import com.pivotal.stockticker.model.Price;
import com.pivotal.stockticker.model.SettingsManager;
import com.pivotal.stockticker.service.apis.*;
import com.pivotal.stockticker.utils.CallbackInterface;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import java.util.*;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.ScheduledFuture;
import java.util.concurrent.TimeUnit;
import java.util.prefs.Preferences;

/**
 * Manages stock prices including loading from storage and periodic updates
 * All prices are stored in user preferences so that they persist between application runs
 * Prices are identified by their stock code (e.g. "AAPL", "GOOGL")
 * The prices are updated periodically based on the application settings
 */
@Slf4j
public class PricesManager {

    private static final String PRICES_ROOT = PersistanceManager.ROOT_NODE + Price.class.getSimpleName();
    private Preferences prefs = Preferences.userRoot().node(PRICES_ROOT);

    private final Map<String, Price> currentPrices = new TreeMap<>(String::compareToIgnoreCase);

    private final PriceCurrencyUpdateTask scheduler;
    private final CallbackInterface callback;

    /**
     * Constructor - loads all prices from persistent storage and starts the periodic update task
     *
     * @param callback Callback interface for notifying UI of updates
     */
    public PricesManager(CallbackInterface callback) {
        this.callback = callback;

        // Load all the saved prices values from persistent storage
        loadFromStorage();

        // Schedule the task to run every X seconds with an initial delay of 0 seconds
        scheduler = new PriceCurrencyUpdateTask(this);
    }

    /**
     * Load all symbols from persistent storage into memory
     */
    public void loadFromStorage() {

        // Load all the prices from the persistent storage
        prefs = Preferences.userRoot().node(PRICES_ROOT);
        try {
            for (String code : prefs.childrenNames()) {
                currentPrices.put(code, Price.getPrice(code));
            }
        }
        catch (Exception e) {
            log.error("Error accessing storage: {}", e.getMessage());
        }
        log.info("Loaded {} prices", currentPrices.size());
    }

    /**
     * Replaces all the current prices with new ones
     *
     * @param codes Collection of stock codes
     */
    public void replacePrices(Collection<String> codes) {
        if (codes == null || codes.isEmpty()) {
            return;
        }
        // Add or update prices from the new list
        for (String symbol : codes) {
            addPrice(symbol);
        }

        // Find all the redundant prices that are no longer needed
        List<String> codesToDelete = new ArrayList<>();
        for (String code : currentPrices.keySet()) {
            if (!codes.contains(code)) {
                codesToDelete.add(code);
            }
        }

        // Remove all prices that are no longer needed
        for (String code : codesToDelete) {
            currentPrices.remove(code);
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
     * Create a new Price
     *
     * @param code Symbol code
     * @return Newly created Price with defaults
     */
    public Price addPrice(String code) {
        if (currentPrices.containsKey(code)) {
            return currentPrices.get(code);
        }
        try {
            Price price = Price.getPrice(code);
            currentPrices.put(code, price);
            log.debug("Added {} price", code);
            return price;
        }
        catch (Exception e) {
            log.error("Error creating new price: {}", e.getMessage());
            return null;
        }
    }

    /**
     * Retrieve the price for a given stock code
     *
     * @param code The stock code
     * @return The Price object for the given code, or null if not found
     */
    public Price getPrice(String code) {
        return currentPrices.get(code);
    }

    /**
     * Get all current prices
     *
     * @return A map of stock codes to their corresponding Price objects
     */
    public Map<String, Price> getAllPrices() {
        return currentPrices;
    }

    /**
     * Update settings from the application
     */
    public void startScheduler() {
        SettingsManager settings = SettingsManager.getInstance();
        if (!scheduler.isRunning() || settings.getFrequency() != scheduler.getPeriodSeconds()) {
            scheduler.start(settings.getFrequency());
        }
    }

    /**
     * Stop the scheduler
     */
    public void stopScheduler() {
        if (scheduler.isRunning()) {
            scheduler.stop();
        }
    }

    /**
     * Refresh prices from all sources
     * This is a synchronous call that will block until all rates are updated
     * It's useful to call this method when the user requests a manual refresh
     * after adding/deleting symbols or changing settings
     */
    public void refreshPrices() {
        refreshPrices(false);
    }

    /**
     * Refresh prices from all sources and notify the callback if requested
     *
     * @param notifyCallback True to notify the callback after updating prices
     */
    public void refreshPrices(boolean notifyCallback) {

        // Take a snapshot of the symbols under the lock so we don't hold it during I/O
        Collection<String> symbols;
        synchronized (this) {
            List<String> symbolsList = new ArrayList<>(currentPrices.keySet());
            symbolsList.sort(Comparator.comparingDouble(code -> {
                Price price = currentPrices.get(code);
                return price != null ? price.getCurrentPrice() : Double.MIN_VALUE;
            }));
            symbols = symbolsList;
        }

        // Perform all HTTP calls outside the lock to avoid blocking other callers
        PricesApiAdapter adapter = new AlphaVantageAdapter(this);
        symbols = adapter.fetchAndUpdatePrices(symbols);

        adapter = new MarketStackAdapter(this);
        symbols = adapter.fetchAndUpdatePrices(symbols);

        adapter = new TwelveDataAdapter(this);
        symbols = adapter.fetchAndUpdatePrices(symbols);

        adapter = new FinnHubAdapter(this);
        symbols = adapter.fetchAndUpdatePrices(symbols);

        adapter = new TiingoAdapter(this);
        symbols = adapter.fetchAndUpdatePrices(symbols);

        adapter = new YahooAdapter(this);
        symbols = adapter.fetchAndUpdatePrices(symbols);

        // Log any symbols that were not updated
        if (!symbols.isEmpty()) {
            log.warn("Prices not updated for symbols: {}", String.join(", ", symbols));
        }

        // Notify the callback that prices have been updated
        if (notifyCallback && callback != null) {
            callback.changed(this);
        }
    }

    /**
     * Scheduler to update prices periodically
     */
    @Slf4j
    private static class PriceCurrencyUpdateTask {

        private final ScheduledExecutorService scheduler = Executors.newSingleThreadScheduledExecutor();
        private ScheduledFuture<?> scheduledFuture;
        private PricesManager prices = null;
        @Getter
        private int periodSeconds;

        /**
         * Constructor
         *
         * @param prices Map of current prices
         */
        public PriceCurrencyUpdateTask(PricesManager prices) {
            this.prices = prices;
        }

        /**
         * Return a true if the periodic task is currently running
         */
        public boolean isRunning() {
            return scheduledFuture != null && !scheduledFuture.isCancelled();
        }

        /**
         * The periodic task to update prices
         */
        private final Runnable task = () -> {

            // Get the list of stock symbols to update prices for
            log.debug("Updating prices for {} symbols", prices.currentPrices.size());

            // Update prices for each symbol from each source
            prices.refreshPrices(true);
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
