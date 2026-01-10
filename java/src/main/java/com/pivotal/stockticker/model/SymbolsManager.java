/*
 *
 * Copyright (c) 2026, 4NG and/or its affiliates. All rights reserved.
 * 4NG PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.model;

import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import java.util.*;
import java.util.prefs.Preferences;

/**
 *
 */
@Slf4j
public class SymbolsManager {

    private static final String SYMBOLS_ROOT = PersistanceManager.ROOT_NODE + SymbolsManager.class.getSimpleName();
    private final Preferences prefs = Preferences.userRoot().node(SYMBOLS_ROOT);

    @Getter
    private final Map<String, SymbolTransaction> symbolTransactions = new LinkedHashMap<>();
    private final Map<String, SymbolTransaction> newSymbolTransactions = new LinkedHashMap<>();
    private final Set<String> modifiedSymbolTransactions = new LinkedHashSet<>();
    private final Set<String> deletedSymbolTransactions = new LinkedHashSet<>();

    /**
     * Constructor - loads all symbols from persistent storage
     */
    public SymbolsManager() {
        loadSymbolsFromStorage();
    }

    /**
     * Load all symbols from persistent storage into memory
     */
    private void loadSymbolsFromStorage() {

        // Load all the symbols from the persistent storage
        try {
            for (String timestamp : prefs.keys()) {
                SymbolTransaction symbolTransaction = SymbolTransaction.getSymbolTransaction(timestamp);
                log.debug("Found symbol: {} with timestamp: {}", symbolTransaction.getCurrencySymbol(), timestamp);
                symbolTransactions.put(timestamp, symbolTransaction);
            }
        }
        catch (Exception e) {
            log.error("Error accessing storage: {}", e.getMessage());
        }
    }

    /**
     * Create a new symbol transaction
     *
     * @return Newly created SymbolTransaction with defaults
     */
    public SymbolTransaction createNewSymbolTransaction() {
        try {
            SymbolTransaction symbolTransaction = SymbolTransaction.getSymbolTransaction();
            newSymbolTransactions.put(symbolTransaction.getKey(), symbolTransaction);
            symbolTransactions.put(symbolTransaction.getKey(), symbolTransaction);
            return symbolTransaction;
        }
        catch (Exception e) {
            log.error("Error creating new symbol: {}", e.getMessage());
            return null;
        }
    }

    /**
     * Mark a symbol as modified
     *
     * @param symbolKey Key of the symbol to mark as modified
     */
    public void markSymbolTransactionAsModified(String symbolKey) {
        modifiedSymbolTransactions.add(symbolKey);
    }

    /**
     * Mark a symbol as deleted
     *
     * @param symbolKey Key of the symbol to mark as deleted
     */
    public void markSymbolTransactionAsDeleted(String symbolKey) {
        deletedSymbolTransactions.add(symbolKey);
    }

    /**
     * Persist all changes (new, modified, deleted symbols) to storage
     */
    public void persistChanges() {

        // Persist new symbols
        for (SymbolTransaction symbolTransaction : newSymbolTransactions.values()) {
            symbolTransaction.saveToStorage();
        }
        newSymbolTransactions.clear();

        // Persist modified symbols
        for (String symbolKey : modifiedSymbolTransactions) {
            SymbolTransaction symbolTransaction = symbolTransactions.get(symbolKey);
            if (symbolTransaction != null) {
                symbolTransaction.saveToStorage();
            }
        }
        modifiedSymbolTransactions.clear();

        // Remove deleted symbols
        for (String symbolKey : deletedSymbolTransactions) {
            symbolTransactions.remove(symbolKey);
            prefs.remove(symbolKey);
        }
        deletedSymbolTransactions.clear();
    }

    /**
     * Get a set of all unique symbol codes (case insensitive)
     *
     * @return Set of unique symbol codes
     */
    public Set<String> getAllSymbolCodes() {
        Set<String> symbols = new TreeSet<>(String.CASE_INSENSITIVE_ORDER);
        for (SymbolTransaction transaction : symbolTransactions.values()) {
            symbols.add(transaction.getCode());
        }
        return symbols;
    }

    /**
     * Get a set of all unique currency codes (case insensitive)
     *
     * @return Set of unique currency codes
     */
    public Set<String> getAllCurrencyCodes() {
        Set<String> symbols = new TreeSet<>(String.CASE_INSENSITIVE_ORDER);
        for (SymbolTransaction transaction : symbolTransactions.values()) {
            symbols.add(transaction.getCurrencySymbol());
        }
        return symbols;
    }

}
