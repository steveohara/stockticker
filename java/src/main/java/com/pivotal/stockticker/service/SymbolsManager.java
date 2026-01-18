/*
 *
 * Copyright (c) 2026, 4NG and/or its affiliates. All rights reserved.
 * 4NG PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.service;

import com.pivotal.stockticker.model.SymbolTransaction;
import lombok.extern.slf4j.Slf4j;

import java.util.*;
import java.util.prefs.BackingStoreException;
import java.util.prefs.Preferences;

/**
 * Manages symbol transactions including loading from storage, tracking changes, and persisting updates
 */
@Slf4j
public class SymbolsManager {

    private static final String SYMBOLS_ROOT = PersistanceManager.ROOT_NODE + SymbolTransaction.class.getSimpleName();
    private Preferences prefs = Preferences.userRoot().node(SYMBOLS_ROOT);

    private final Set<SymbolTransaction> symbolTransactions = new LinkedHashSet<>();
    private final Set<SymbolTransaction> newSymbolTransactions = new LinkedHashSet<>();
    private final Set<SymbolTransaction> modifiedSymbolTransactions = new LinkedHashSet<>();
    private final Set<SymbolTransaction> deletedSymbolTransactions = new LinkedHashSet<>();

    /**
     * Constructor - loads all symbols from persistent storage
     */
    public SymbolsManager() {
        loadFromStorage();
    }

    /**
     * Load all symbols from persistent storage into memory
     */
    public void loadFromStorage() {

        // Load all the symbols from the persistent storage
        prefs = Preferences.userRoot().node(SYMBOLS_ROOT);
        try {
            List<SymbolTransaction> symbols = new ArrayList<>();
            for (String timestamp : prefs.childrenNames()) {
                try {
                    symbols.add(SymbolTransaction.getSymbolTransaction(timestamp));
                }
                catch (Exception e) {
                    log.error("Failed to load symbol transaction", e);
                }
            }

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
        return createNewSymbolTransaction(null);
    }

    /**
     * Create a new symbol transaction using a key
     *
     * @param key Key to use for the new symbol transaction
     * @return Newly created SymbolTransaction with defaults
     */
    public SymbolTransaction createNewSymbolTransaction(String key) {
        try {
            SymbolTransaction symbolTransaction = SymbolTransaction.getSymbolTransaction(key);
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
     * Clear all tracked changes
     */
    public void clearChanges() {
        newSymbolTransactions.clear();
        modifiedSymbolTransactions.clear();
        deletedSymbolTransactions.clear();
        loadFromStorage();
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

        // Clear everything
        deletedSymbolTransactions.clear();

        // Load them all back from storage
        loadFromStorage();
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
        return getSymbolTransactions(true, false, null);
    }

    /**
     * Returns a sorted list of all symbol transactions
     *
     * @param includeDisabled Whether to include disabled symbols
     * @param uniqueOnly     Whether to include only unique symbols (first occurrence)
     * @param symbolCode Filter by symbol code (case insensitive), or null for all symbols
     * @return List of SymbolTransaction objects
     */
    public List<SymbolTransaction> getSymbolTransactions(boolean includeDisabled, boolean uniqueOnly, String symbolCode) {
        Map<String, SymbolTransaction> symbols = new TreeMap<>(String::compareToIgnoreCase);

        // Loop round all the symbol transactions
        for (SymbolTransaction transaction : symbolTransactions) {

            // If only unique symbols are requested, skip if already added
            String code = transaction.getCode();
            if (!uniqueOnly || !symbols.containsKey(code)) {

                // Check if symbol code matches filter and if disabled symbols are included
                if ((symbolCode == null || transaction.getCode().equalsIgnoreCase(symbolCode))
                        && (includeDisabled || !transaction.isDisabled())) {
                    symbols.put(uniqueOnly ? transaction.getCode() : transaction.getKey(), transaction);
                }
            }
        }
        List<SymbolTransaction> list = new ArrayList<>(symbols.values());
        list.sort(Comparator.comparing(SymbolTransaction::getSortKey, String.CASE_INSENSITIVE_ORDER));
        return list;
    }
}
