package com.pivotal.stockticker.model;

import com.pivotal.stockticker.service.ExchangeRatesManager;
import com.pivotal.stockticker.service.PricesManager;
import com.pivotal.stockticker.service.SymbolsManager;
import lombok.AllArgsConstructor;

@AllArgsConstructor
public class SummaryStats {
    private final SymbolsManager symbols;
    private final PricesManager prices;
    private final ExchangeRatesManager rates;
    private final SettingsManager settings;

    public double calculateTotalValue() {
//        return symbols.getAllEnabledSymbols().stream()
//                .mapToDouble(symbol -> {
//                    Double price = prices.getPrice(symbol.getSymbol());
//                    if (price == null) {
//                        return 0.0;
//                    }
//                    double exchangeRate = 1.0;
//                    if (symbol.getCurrencyCode() != null && !symbol.getCurrencyCode().equals(settings.getCurrencyCode())) {
//                        Double rate = rates.getExchangeRate(symbol.getCurrencyCode(), settings.getCurrencyCode());
//                        if (rate != null) {
//                            exchangeRate = rate;
//                        }
//                    }
//                    return price * symbol.getQuantity() * exchangeRate;
//                })
//                .sum();
        return 0;
    }

    public double calculateTotalCost() {
        return 0;
    }
}
