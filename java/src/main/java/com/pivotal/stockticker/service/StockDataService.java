package com.pivotal.stockticker.service;

import com.pivotal.stockticker.model.SymbolTransaction;
import lombok.extern.slf4j.Slf4j;

import java.time.LocalDateTime;
import java.util.HashMap;
import java.util.Map;
import java.util.Random;

/**
 * Service for fetching stock data. Uses mock data for demonstration.
 */
@Slf4j
public class StockDataService {
    private final Random random = new Random();
    private final Map<String, Double> lastPrices = new HashMap<>();

    public void updateSymbolPrice(SymbolTransaction symbolTransaction) {
        if (symbolTransaction.isDisabled()) {
            return;
        }
        String code = symbolTransaction.getCode();
        if (!lastPrices.containsKey(code)) {
            lastPrices.put(code, symbolTransaction.getPricePaid() > 0 ? symbolTransaction.getPricePaid() : 100.0);
            symbolTransaction.setDayStart(symbolTransaction.getPricePaid());
        }
        double lastPrice = lastPrices.get(code);
        double changePercent = (random.nextDouble() * 4) - 2;
        double newPrice = lastPrice * (1 + changePercent / 100);
        symbolTransaction.setCurrentPrice(newPrice);
        symbolTransaction.setLastPriceUpdate(LocalDateTime.now());
        symbolTransaction.setDayHigh(Math.max(symbolTransaction.getDayHigh(), newPrice));
        symbolTransaction.setDayLow(Math.min(symbolTransaction.getDayLow() == 0 ? newPrice : symbolTransaction.getDayLow(), newPrice));
        if (symbolTransaction.getDayStart() > 0) {
            symbolTransaction.setDayChange(newPrice - symbolTransaction.getDayStart());
        }
        lastPrices.put(code, newPrice);
        checkAlarms(symbolTransaction);
    }

    private void checkAlarms(SymbolTransaction symbolTransaction) {
        double currentPrice = symbolTransaction.getCurrentPrice();
        double baseCost = symbolTransaction.getPricePaid();
        if (symbolTransaction.isHighAlarmEnabled() && !symbolTransaction.isAlarmShowing()) {
            double threshold = symbolTransaction.isHighAlarmIsPercent() ?
                    baseCost * (1 + symbolTransaction.getHighAlarmValue() / 100) : symbolTransaction.getHighAlarmValue();
            if (currentPrice >= threshold) {
                triggerAlarm(symbolTransaction, true);
            }
        }
        if (symbolTransaction.isLowAlarmEnabled() && !symbolTransaction.isAlarmShowing()) {
            double threshold = symbolTransaction.isLowAlarmIsPercent() ?
                    baseCost * (1 - symbolTransaction.getLowAlarmValue() / 100) : symbolTransaction.getLowAlarmValue();
            if (currentPrice <= threshold) {
                triggerAlarm(symbolTransaction, false);
            }
        }
    }

    private void triggerAlarm(SymbolTransaction symbolTransaction, boolean isHighAlarm) {
        symbolTransaction.setAlarmShowing(true);
        System.out.println("ALARM: " + symbolTransaction.getCode() + " - " +
                (isHighAlarm ? "HIGH" : "LOW") + " threshold reached!");
    }
}
