/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.service;

import com.pivotal.stockticker.Utils;
import com.pivotal.stockticker.model.LivePrice;
import com.pivotal.stockticker.model.Price;
import com.pivotal.stockticker.model.SettingsManager;
import com.pivotal.stockticker.model.SymbolTransaction;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import java.time.LocalDateTime;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CopyOnWriteArrayList;

/**
 * Watches live prices against the low/high alarm thresholds configured on each symbol transaction
 * and raises {@link AlarmEvent}s when a threshold is crossed. Equivalent to the alarm checking
 * that used to live in the main polling loop of the VB6 version (Z_GetSymbolData).
 */
@Slf4j
public class AlarmManager {

    /**
     * The type of alarm that was triggered.
     */
    public enum AlarmType {
        LOW, HIGH
    }

    /**
     * Listener notified whenever the set of active alarms changes.
     */
    public interface AlarmListener {
        void alarmsChanged();
    }

    /**
     * Represents a single triggered alarm that has not yet been dismissed.
     */
    @Getter
    public static class AlarmEvent {
        private final SymbolTransaction symbolTransaction;
        private final AlarmType type;
        private final double threshold;
        private final boolean percent;
        private final double priceAtTrigger;
        private final double percentAtTrigger;
        private final LocalDateTime triggeredTime = LocalDateTime.now();
        private boolean muted;

        AlarmEvent(SymbolTransaction symbolTransaction, AlarmType type, double threshold, boolean percent, double priceAtTrigger, double percentAtTrigger) {
            this.symbolTransaction = symbolTransaction;
            this.type = type;
            this.threshold = threshold;
            this.percent = percent;
            this.priceAtTrigger = priceAtTrigger;
            this.percentAtTrigger = percentAtTrigger;
        }

        public void setMuted(boolean muted) {
            this.muted = muted;
        }
    }

    private final SymbolsManager symbols;
    private final PricesManager prices;
    private final ExchangeRatesManager rates;
    private final List<AlarmEvent> activeAlarms = new CopyOnWriteArrayList<>();
    private final Map<String, Double> lastCheckedPrice = new HashMap<>();
    private final Map<String, Double> lastCheckedPercent = new HashMap<>();
    private AlarmListener listener;

    /**
     * Constructor
     *
     * @param symbols Symbols manager to source alarm configuration from
     * @param prices  Prices manager to source live prices from
     * @param rates   Exchange rates manager, required to build {@link LivePrice} instances
     */
    public AlarmManager(SymbolsManager symbols, PricesManager prices, ExchangeRatesManager rates) {
        this.symbols = symbols;
        this.prices = prices;
        this.rates = rates;

        // Any symbol left showing an alarm from a previous run (e.g. after a crash) can't be
        // reliably attributed back to a LOW or HIGH alarm, so clear it and let the next check
        // re-evaluate the current state from scratch.
        for (SymbolTransaction symbol : symbols.getSymbolTransactions(true, false, null)) {
            if (symbol.isAlarmShowing()) {
                symbol.setAlarmShowing(false);
                symbol.saveToStorage();
            }
        }
    }

    /**
     * Sets the listener to be notified when the active alarm list changes.
     *
     * @param listener Listener to notify
     */
    public void setListener(AlarmListener listener) {
        this.listener = listener;
    }

    /**
     * Returns the list of currently active (undismissed) alarms.
     *
     * @return List of active alarms
     */
    public List<AlarmEvent> getActiveAlarms() {
        return activeAlarms;
    }

    /**
     * Checks all enabled symbol transactions against their configured alarm thresholds and raises
     * any new alarms that have been crossed since the last check.
     */
    public void checkAlarms() {
        for (SymbolTransaction symbol : symbols.getSymbolTransactions(false, false, null)) {
            if (symbol.isDisabled() || (!symbol.isLowAlarmEnabled() && !symbol.isHighAlarmEnabled())) {
                continue;
            }
            Price price = prices.getPrice(symbol.getCode());
            if (price == null || price.getCurrentPrice() == 0) {
                continue;
            }

            LivePrice livePrice = new LivePrice(symbols, prices, rates, symbol, false);
            double currentPrice = livePrice.getPrice();
            double currentPercent = livePrice.getPercentChange();

            Double previousPrice = lastCheckedPrice.get(symbol.getKey());
            Double previousPercent = lastCheckedPercent.get(symbol.getKey());

            if (!symbol.isAlarmShowing()) {
                if (symbol.isLowAlarmEnabled()) {
                    double value = symbol.isLowAlarmIsPercent() ? currentPercent : currentPrice;
                    Double previousValue = symbol.isLowAlarmIsPercent() ? previousPercent : previousPrice;
                    boolean triggeredNow = value <= symbol.getLowAlarmValue();
                    boolean triggeredPreviously = previousValue != null && previousValue <= symbol.getLowAlarmValue();
                    if (triggeredNow && !triggeredPreviously) {
                        fireAlarm(symbol, AlarmType.LOW, symbol.getLowAlarmValue(), symbol.isLowAlarmIsPercent(), currentPrice, currentPercent, symbol.isLowAlarmSoundEnabled());
                    }
                }
                if (symbol.isHighAlarmEnabled() && !symbol.isAlarmShowing()) {
                    double value = symbol.isHighAlarmIsPercent() ? currentPercent : currentPrice;
                    Double previousValue = symbol.isHighAlarmIsPercent() ? previousPercent : previousPrice;
                    boolean triggeredNow = value >= symbol.getHighAlarmValue();
                    boolean triggeredPreviously = previousValue != null && previousValue >= symbol.getHighAlarmValue();
                    if (triggeredNow && !triggeredPreviously) {
                        fireAlarm(symbol, AlarmType.HIGH, symbol.getHighAlarmValue(), symbol.isHighAlarmIsPercent(), currentPrice, currentPercent, symbol.isHighAlarmSoundEnabled());
                    }
                }
            }

            lastCheckedPrice.put(symbol.getKey(), currentPrice);
            lastCheckedPercent.put(symbol.getKey(), currentPercent);
        }
    }

    /**
     * Raises a new alarm event, marks the symbol as showing an alarm, persists the change and
     * plays the configured alarm sound if requested.
     */
    private void fireAlarm(SymbolTransaction symbol, AlarmType type, double threshold, boolean percent, double currentPrice, double currentPercent, boolean soundEnabled) {
        log.info("{} alarm triggered for {} - threshold {}{}, current {}", type, symbol.getDisplayName(), threshold, percent ? "%" : "", percent ? currentPercent : currentPrice);

        symbol.setAlarmShowing(true);
        symbol.saveToStorage();

        AlarmEvent event = new AlarmEvent(symbol, type, threshold, percent, currentPrice, currentPercent);
        activeAlarms.add(event);

        if (soundEnabled) {
            SettingsManager settings = SettingsManager.getInstance();
            Utils.playAlarmSound(type == AlarmType.HIGH ? settings.getHighAlarmWaveFile() : settings.getLowAlarmWaveFile());
        }

        if (listener != null) {
            listener.alarmsChanged();
        }
    }

    /**
     * Dismisses an alarm, allowing it to be triggered again once the price recovers past the
     * threshold and crosses it again.
     *
     * @param event Alarm event to dismiss
     */
    public void dismissAlarm(AlarmEvent event) {
        activeAlarms.remove(event);
        event.getSymbolTransaction().setAlarmShowing(false);
        event.getSymbolTransaction().saveToStorage();
        if (listener != null) {
            listener.alarmsChanged();
        }
    }

    /**
     * Dismisses an alarm and disables it so that it will not be checked again until re-enabled
     * from the Symbols dialog.
     *
     * @param event Alarm event to disable
     */
    public void disableAlarm(AlarmEvent event) {
        if (event.getType() == AlarmType.LOW) {
            event.getSymbolTransaction().setLowAlarmEnabled(false);
        }
        else {
            event.getSymbolTransaction().setHighAlarmEnabled(false);
        }
        dismissAlarm(event);
    }

    /**
     * Dismisses all active alarms.
     */
    public void dismissAllAlarms() {
        for (AlarmEvent event : activeAlarms) {
            event.getSymbolTransaction().setAlarmShowing(false);
            event.getSymbolTransaction().saveToStorage();
        }
        activeAlarms.clear();
        if (listener != null) {
            listener.alarmsChanged();
        }
    }
}
