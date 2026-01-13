/*
 *
 * Copyright (c) 2026, 4NG and/or its affiliates. All rights reserved.
 * 4NG PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.model;

import lombok.extern.slf4j.Slf4j;

import java.util.*;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.prefs.BackingStoreException;
import java.util.prefs.Preferences;

/**
 * Manages symbol transactions including loading from storage, tracking changes, and persisting updates
 */
@Slf4j
public class SymbolsManager {

    private static final String SYMBOLS_ROOT = PersistanceManager.ROOT_NODE + SymbolTransaction.class.getSimpleName();
    private final Preferences prefs = Preferences.userRoot().node(SYMBOLS_ROOT);

    private final Set<SymbolTransaction> symbolTransactions = new LinkedHashSet<>();
    private final Set<SymbolTransaction> newSymbolTransactions = new LinkedHashSet<>();
    private final Set<SymbolTransaction> modifiedSymbolTransactions = new LinkedHashSet<>();
    private final Set<SymbolTransaction> deletedSymbolTransactions = new LinkedHashSet<>();

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
            List<SymbolTransaction> symbols = new ArrayList<>();
            ExecutorService executor = Executors.newVirtualThreadPerTaskExecutor();
            for (String timestamp : prefs.childrenNames()) {
                executor.submit(() -> {
                    try {
                        symbols.add(SymbolTransaction.getSymbolTransaction(timestamp));
                    }
                    catch (Exception e) {
                        log.error("Failed to load symbol transaction", e);
                    }
                });
            }
            // Wait for all the tasks to complete
            executor.close();

            // Sort the symbols by code
            symbols.sort(Comparator.comparing(SymbolTransaction::getSortKey, String.CASE_INSENSITIVE_ORDER));
            symbolTransactions.clear();
            symbolTransactions.addAll(symbols);
        }
        catch (Exception e) {
            log.error("Error accessing storage: {}", e.getMessage());
        }
        log.info("Loaded {} symbol transactions", symbolTransactions.size());
    }

    /**
     * Create a new symbol transaction
     *
     * @return Newly created SymbolTransaction with defaults
     */
    public SymbolTransaction createNewSymbolTransaction() {
        try {
            SymbolTransaction symbolTransaction = SymbolTransaction.getSymbolTransaction();
            newSymbolTransactions.add(symbolTransaction);
            symbolTransactions.add(symbolTransaction);
            symbolTransaction.setAdded(true);
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
     * @param symbol Symbol that has been changed
     */
    public void markSymbolTransactionAsModified(SymbolTransaction symbol) {
        modifiedSymbolTransactions.add(symbol);
        symbol.setEdited(true);
    }

    /**
     * Mark a symbol as deleted
     *
     * @param symbolTransaction The symbol to mark as deleted
     */
    public void markSymbolTransactionAsDeleted(SymbolTransaction symbolTransaction) {
        symbolTransactions.remove(symbolTransaction);
        deletedSymbolTransactions.add(symbolTransaction);
    }

    /**
     * Persist all changes (new, modified, deleted symbols) to storage
     */
    public void persistChanges() {

        // Persist new symbols
        for (SymbolTransaction symbolTransaction : newSymbolTransactions) {
            symbolTransaction.saveToStorage();
            symbolTransaction.setEdited(false);
            symbolTransaction.setAdded(false);
        }
        newSymbolTransactions.clear();

        // Persist modified symbols
        for (SymbolTransaction symbolTransaction : modifiedSymbolTransactions) {
            symbolTransaction.saveToStorage();
            symbolTransaction.setEdited(false);
            symbolTransaction.setAdded(false);
        }
        modifiedSymbolTransactions.clear();

        // Remove deleted symbols
        for (SymbolTransaction symbolTransaction : deletedSymbolTransactions) {
            symbolTransaction.setEdited(false);
            symbolTransaction.setAdded(false);
            symbolTransactions.remove(symbolTransaction);

            // Remove the values from storage
            try {
                Preferences node = prefs.node(symbolTransaction.getKey());
                if (node != null) {
                    node.removeNode();
                    prefs.flush();
                }
            }
            catch (BackingStoreException e) {
                log.error("Error removing symbol transaction: {}", e.getMessage());
            }
        }
        deletedSymbolTransactions.clear();
    }

    /**
     * Get a set of all unique symbol codes (case insensitive)
     *
     * @param includeDisabled Whether to include disabled symbols
     * @return Set of unique symbol codes
     */
    public Set<String> getAllSymbolCodes(boolean includeDisabled) {
        Set<String> symbols = new TreeSet<>(String.CASE_INSENSITIVE_ORDER);
        for (SymbolTransaction transaction : symbolTransactions) {
            if (!includeDisabled && transaction.isDisabled()) {
                continue;
            }
            symbols.add(transaction.getCode());
        }
        return symbols;
    }

    /**
     * Get a set of all unique currency codes (case insensitive)
     *
     * @param includeDisabled Whether to include disabled symbols
     * @return Set of unique currency codes
     */
    public Set<String> getAllCurrencyCodes(boolean includeDisabled) {
        Set<String> symbols = new TreeSet<>(String.CASE_INSENSITIVE_ORDER);
        for (SymbolTransaction transaction : symbolTransactions) {
            if (!includeDisabled && transaction.isDisabled()) {
                continue;
            }
            symbols.add(transaction.getCurrencyCode());
        }
        return symbols;
    }

    /**
     * Returns a sorted list of all symbol transactions
     *
     * @return List of SymbolTransaction objects
     */
    public List<SymbolTransaction> getSymbolTransactions() {
        return getSymbolTransactions(null, true);
    }

    /**
     * Returns a sorted list of all symbol transactions
     *
     * @param symbolCode Filter by symbol code (case insensitive), or null for all symbols
     * @param includeDisabled Whether to include disabled symbols
     * @return List of SymbolTransaction objects
     */
    public List<SymbolTransaction> getSymbolTransactions(String symbolCode, boolean includeDisabled) {
        List<SymbolTransaction> symbols = new ArrayList<>();
        for (SymbolTransaction transaction : symbolTransactions) {
            if ((symbolCode == null || transaction.getCode().equalsIgnoreCase(symbolCode))
                    && (includeDisabled || !transaction.isDisabled())) {
                symbols.add(transaction);
            }
        }
        symbols.sort(Comparator.comparing(SymbolTransaction::getSortKey, String.CASE_INSENSITIVE_ORDER));
        return symbols;
    }

    /**
     * Get the first symbol transaction matching the given symbol code (case insensitive)
     *
     * @param symbol Symbol code to search for
     * @return Matching SymbolTransaction or null if not found
     */
    public SymbolTransaction getFirst(String symbol) {
        for (SymbolTransaction transaction : symbolTransactions) {
            if (transaction.getCode().equalsIgnoreCase(symbol)) {
                return transaction;
            }
        }
        return null;

    }
}
