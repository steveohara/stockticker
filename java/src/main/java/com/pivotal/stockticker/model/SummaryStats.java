/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.model;

import com.pivotal.stockticker.service.ExchangeRatesManager;
import com.pivotal.stockticker.service.PricesManager;
import com.pivotal.stockticker.service.SymbolsManager;
import lombok.AllArgsConstructor;
import lombok.extern.slf4j.Slf4j;

@Slf4j
@AllArgsConstructor
public class SummaryStats {
    private final SymbolsManager symbols;
    private final PricesManager prices;
    private final ExchangeRatesManager rates;

    /**
     * Calculate the total value of all symbol transactions in the base currency
     *
     * @return Total value of all symbol transactions
     */
    public double calculateTotalValue() {
        double total = 0;

        // Iterate through all symbol transactions
        // Note: - we use getSymbolTransactions without condensing to ensure we get all transactions
        //         in whatever currencies they were paid in
        for (SymbolTransaction transaction : symbols.getSymbolTransactions(false, false, null)) {

            // Get the total value of each transaction
            Price price = prices.getPrice(transaction.getCode());
            if (price == null) {
                log.warn("No available Price for transaction {} ", transaction.getCode());
                continue;
            }
            double totalValue = transaction.getSharesBought() * price.getCurrentPrice();

            // Convert the total value to the base currency if needed
            total += rates.convertAmount(transaction, totalValue);
        }
        return total;
    }

    /**
     * Calculate the total cost of all symbol transactions in the base currency
     *
     * @return Total cost of all symbol transactions
     */
    public double calculateTotalCost() {
        double total = 0;

        // Iterate through all symbol transactions
        // Note: - we use getSymbolTransactions without condensing to ensure we get all transactions
        //         in whatever currencies they were paid in
        for (SymbolTransaction transaction : symbols.getSymbolTransactions(false, false, null)) {

            // Get the total value of each transaction
            double totalValue = transaction.getSharesBought() * transaction.getPricePaid();

            // Convert the total value to the base currency
            total += rates.convertAmount(transaction, totalValue);
        }
        return total;
    }

    /**
     * Calculate the total value of all symbol transactions at the start of the day in the base currency
     *
     * @return Total value of all symbol transactions at the start of the day
     */
    public double calculateTotalValueAtStartOfDay() {
        double total = 0;

        // Iterate through all symbol transactions
        // Note: - we use getSymbolTransactions without condensing to ensure we get all transactions
        //         in whatever currencies they were paid in
        for (SymbolTransaction transaction : symbols.getSymbolTransactions(false, false, null)) {

            // Get the total value of each transaction
            Price price = prices.getPrice(transaction.getCode());
            if (price == null) {
                log.warn("No available Price for transaction {} ", transaction.getCode());
                continue;
            }
            double totalValue = transaction.getSharesBought() * price.getDayStart();

            // Convert the total value to the base currency if needed
            total += rates.convertAmount(transaction, totalValue);
        }
        return total;
    }

}
