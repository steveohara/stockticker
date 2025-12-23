package com.pivotal.stockticker.service;

import com.pivotal.stockticker.model.Symbol;
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

    public void updateSymbolPrice(Symbol symbol) {
        if (symbol.isDisabled()) {
            return;
        }
        String code = symbol.getCode();
        if (!lastPrices.containsKey(code)) {
            lastPrices.put(code, symbol.getPrice() > 0 ? symbol.getPrice() : 100.0);
            symbol.setDayStart(symbol.getPrice());
        }
        double lastPrice = lastPrices.get(code);
        double changePercent = (random.nextDouble() * 4) - 2;
        double newPrice = lastPrice * (1 + changePercent / 100);
        symbol.setCurrentPrice(newPrice);
        symbol.setLastUpdate(LocalDateTime.now());
        symbol.setDayHigh(Math.max(symbol.getDayHigh(), newPrice));
        symbol.setDayLow(Math.min(symbol.getDayLow() == 0 ? newPrice : symbol.getDayLow(), newPrice));
        if (symbol.getDayStart() > 0) {
            symbol.setDayChange(newPrice - symbol.getDayStart());
        }
        lastPrices.put(code, newPrice);
        checkAlarms(symbol);
    }

    private void checkAlarms(Symbol symbol) {
        double currentPrice = symbol.getCurrentPrice();
        double baseCost = symbol.getPrice();
        if (symbol.isHighAlarmEnabled() && !symbol.isAlarmShowing()) {
            double threshold = symbol.isHighAlarmIsPercent() ?
                    baseCost * (1 + symbol.getHighAlarmValue() / 100) : symbol.getHighAlarmValue();
            if (currentPrice >= threshold) {
                triggerAlarm(symbol, true);
            }
        }
        if (symbol.isLowAlarmEnabled() && !symbol.isAlarmShowing()) {
            double threshold = symbol.isLowAlarmIsPercent() ?
                    baseCost * (1 - symbol.getLowAlarmValue() / 100) : symbol.getLowAlarmValue();
            if (currentPrice <= threshold) {
                triggerAlarm(symbol, false);
            }
        }
    }

    private void triggerAlarm(Symbol symbol, boolean isHighAlarm) {
        symbol.setAlarmShowing(true);
        System.out.println("ALARM: " + symbol.getCode() + " - " +
                (isHighAlarm ? "HIGH" : "LOW") + " threshold reached!");
    }
}
