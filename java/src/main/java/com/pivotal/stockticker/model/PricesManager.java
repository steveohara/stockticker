package com.pivotal.stockticker.model;

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
 */
@Slf4j
public class PricesManager {

    private static final String SYMBOLS_ROOT = PersistanceManager.ROOT_NODE + Price.class.getSimpleName();
    private final Preferences prefs = Preferences.userRoot().node(SYMBOLS_ROOT);

    private final Map<String, Price> currentPrices = new TreeMap<>(String::compareToIgnoreCase);
    private final PriceCurrencyUpdateTask scheduler;
    private final Settings settings;

    /**
     * Constructor - loads all prices from persistent storage and starts the periodic update task
     *
     * @param settings Application settings
     */
    public PricesManager(Settings settings) {
        this.settings = settings;

        // Load all the saved prices values from persistent storage
        loadPricesFromStorage();

        // Schedule the task to run every X seconds with an initial delay of 0 seconds
        scheduler = new PriceCurrencyUpdateTask(settings, currentPrices);
        scheduler.start(settings.getFrequency());
    }

    /**
     * Load all symbols from persistent storage into memory
     */
    private void loadPricesFromStorage() {
        // Load all the prices from the persistent storage
        try {
            for (String code : prefs.childrenNames()) {
                currentPrices.put(code, Price.getPrice(code));
            }
        }
        catch (Exception e) {
            log.error("Error accessing storage: {}", e.getMessage());
        }
        log.debug("Loaded {} prices", currentPrices.size());
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
    public void updateSettings() {
        if (settings.getFrequency() != scheduler.getPeriodSeconds()) {
            scheduler.start(settings.getFrequency());
        }
    }

    /**
     * Scheduler to update prices periodically
     */
    @Slf4j
    private static class PriceCurrencyUpdateTask {

        private final ScheduledExecutorService scheduler = Executors.newSingleThreadScheduledExecutor();
        private ScheduledFuture<?> scheduledFuture;
        private final Settings settings;
        private Map<String, Price> currentPrices = new HashMap<>();
        @Getter
        private int periodSeconds;

        /**
         * Constructor
         *
         * @param settings      Application settings
         * @param currentPrices Map of current prices
         */
        public PriceCurrencyUpdateTask(Settings settings, Map<String, Price> currentPrices) {
            this.settings = settings;
            this.currentPrices = currentPrices;
        }

        /**
         * The periodic task to update prices
         */
        private final Runnable task = () -> {

            // Get the list of stock symbols to update prices for
            log.debug("Updating prices for {} symbols", currentPrices.size());
            Map<String, Price> prices = new HashMap<>(currentPrices);

            // Update prices for each symbol from each source
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
