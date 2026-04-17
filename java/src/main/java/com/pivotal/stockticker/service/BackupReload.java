/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.service;

import com.pivotal.stockticker.Utils;
import com.pivotal.stockticker.model.SettingsManager;
import com.pivotal.stockticker.model.SymbolTransaction;
import com.pivotal.stockticker.ui.MessageDialog;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.*;
import java.nio.charset.MalformedInputException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.prefs.BackingStoreException;
import java.util.prefs.InvalidPreferencesFormatException;
import java.util.prefs.Preferences;

/**
 * Handles backup and reload operations
 */
@Slf4j
public class BackupReload {

    // File chooser for backup/restore operations
    private static OverwritePromptChooser chooser = null;

    /**
     * Backs up the given Preferences subtree to a user-selected file.
     *
     * @param dialog The parent dialog for the file chooser.
     */
    public static void backupPreferences(Dialog dialog) {

        // Use JFileChooser to let user pick the file location
        if (chooser == null) {
            chooser = new OverwritePromptChooser();
            chooser.setFileFilter(new FileNameExtensionFilter("Backup Files (*.bck)", "bck"));
        }
        chooser.setChooserType(OverwritePromptChooser.CHOOSER_TYPE.SAVE);
        chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        chooser.setAcceptAllFileFilterUsed(false);
        chooser.setMultiSelectionEnabled(false);
        chooser.setDialogType(JFileChooser.SAVE_DIALOG);
        chooser.setDialogTitle("Backup Settings");
        int userSelection = chooser.showDialog(dialog, "Save");

        // If user approved, export the preferences to the selected file
        if (userSelection == JFileChooser.APPROVE_OPTION) {

            // Get the file and make sure it has an extension
            File selectedFile = chooser.getSelectedFile();
            if (!selectedFile.getName().toLowerCase().endsWith(".bck")) {
                selectedFile = new File(selectedFile.getAbsolutePath() + ".bck");
            }

            // Get the file from the chooser and export the preferences
            try (OutputStream os = Files.newOutputStream(selectedFile.toPath())) {
                Preferences prefs = Preferences.userRoot().node(PersistanceManager.ROOT_NODE_NAME);
                prefs.exportSubtree(os);
                MessageDialog.show(dialog, "Settings backed up successfully!", "Backup Successful", JOptionPane.INFORMATION_MESSAGE);
                log.info("Preferences backed up to {}", selectedFile);
            }
            catch (IOException | BackingStoreException ex) {
                MessageDialog.show(dialog, "Export failed: " + ex.getMessage(), "Backup Failed", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    /**
     * Restores Preferences subtree from a user-selected file.
     *
     * @param dialog The parent dialog for the file chooser.
     * @return true if restore was successful, false otherwise.
     */
    public static boolean restorePreferences(JDialog dialog) {

        // Use JFileChooser to let user pick the file location
        if (chooser == null) {
            chooser = new OverwritePromptChooser();
            chooser.setFileFilter(new FileNameExtensionFilter("Backup Files (*.bck)", "bck"));
        }
        chooser.setChooserType(OverwritePromptChooser.CHOOSER_TYPE.OPEN);
        chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        chooser.setAcceptAllFileFilterUsed(false);
        chooser.setMultiSelectionEnabled(false);
        chooser.setDialogType(JFileChooser.OPEN_DIALOG);
        chooser.setDialogTitle("Restore Settings");
        int userSelection = chooser.showDialog(dialog, "Open");

        // If user approved, import the preferences from the selected file
        if (userSelection == JFileChooser.APPROVE_OPTION) {

            // Get the file and make sure it has an extension
            File selectedFile = chooser.getSelectedFile();
            if (selectedFile == null || !selectedFile.exists()) {
                MessageDialog.show(dialog, "Selected file does not exist.", "Restore Failed", JOptionPane.ERROR_MESSAGE);
                return false;
            }

            // Load the preferences from the file
            try {
                loadPreferencesFromFile(selectedFile);
                MessageDialog.show(dialog, "Settings restored successfully!", "Restore Successful", JOptionPane.INFORMATION_MESSAGE);
                log.info("Preferences restored from {}", chooser.getSelectedFile());
                return true;
            }
            catch (Exception ex) {
                log.error("Restore from {} failed", selectedFile, ex);
                MessageDialog.show(dialog, "Restore failed: " + ex.getMessage(), "Restore Failed", JOptionPane.ERROR_MESSAGE);
            }
        }
        return false;
    }

    /**
     * Loads preferences from the given file, determining the format automatically.
     *
     * @param file The file to load preferences from.
     */
    private static void loadPreferencesFromFile(File file) throws Exception {

        // Read the file contents in one go
        String content;
        try {
            content = Files.readString(Path.of(file.getAbsolutePath()), StandardCharsets.UTF_8);
        }
        catch (MalformedInputException mie) {
            content = Files.readString(Path.of(file.getAbsolutePath()), StandardCharsets.ISO_8859_1);
        }

        // Get some transient values that need to be maintained
        SettingsManager settings = SettingsManager.getInstance();
        int left = settings.getWindowX();
        int top = settings.getWindowY();
        int width = settings.getWindowWidth();

        // Determine the format and load accordingly
        if (content.startsWith("Windows Registry Editor Version 5.00")) {
            loadPreferencesFromRegistryString(content);
        }
        else {
            loadPreferencesFromBackupString(content);
        }

        // Replace the positioning settings to maintain window position
        settings = SettingsManager.getInstance(true);
        settings.setWindowX(left);
        settings.setWindowY(top);
        settings.setWindowWidth(width);
    }

    /**
     * Loads preferences from a backup string.
     *
     * @param content The backup string to load preferences from.
     */
    private static void loadPreferencesFromBackupString(String content) throws IOException, BackingStoreException, InvalidPreferencesFormatException {

        // Get the file from the chooser and export the preferences
        try (InputStream is = new ByteArrayInputStream(content.getBytes())) {

            // Load the preferences from the content
            Preferences prefs = Preferences.userRoot().node(PersistanceManager.ROOT_NODE_NAME);
            prefs.removeNode();
            Preferences.importPreferences(is);
        }
    }

    /**
     * Loads preferences from a Windows Registry export string.
     *
     * @param content The registry export string to load preferences from.
     * @throws Exception if an error occurs during loading.
     */
    private static void loadPreferencesFromRegistryString(String content) throws Exception {

        // Delete the preferences node first
        Preferences prefs = Preferences.userRoot().node(PersistanceManager.ROOT_NODE_NAME);
        prefs.removeNode();

        // Now create new managers to load the data into
        SettingsManager settings = SettingsManager.getInstance(true);
        SymbolsManager symbolsManager = new SymbolsManager();
        PricesManager pricesManager = new PricesManager(null);
        ExchangeRatesManager exchangeRatesManager = new ExchangeRatesManager(null);

        // Split the content into lines
        String[] lines = content.split("\\r?\\n");
        for (int line = 0; line < lines.length; line++) {

            // If this is the settings section, load settings
            if (lines[line].equalsIgnoreCase("[HKEY_CURRENT_USER\\SOFTWARE\\Pivotal\\StockTicker\\Settings]")) {
                line = loadSettingsFromRegistryLines(settings, lines, line);
            }

            // If this is a symbol section, load symbol
            else if (lines[line].startsWith("[HKEY_CURRENT_USER\\SOFTWARE\\Pivotal\\StockTicker\\Symbols\\")) {
                line = loadSymbolFromRegistryLines(symbolsManager, pricesManager, exchangeRatesManager, lines, line);
            }
        }
        symbolsManager.persistChanges();
    }

    /**
     * Loads settings from registry lines starting at the given line index.
     *
     * @param settings The Settings instance to load into.
     * @param lines    The registry lines.
     * @param line     The starting line index.
     * @return The next line index after processing.
     */
    private static int loadSettingsFromRegistryLines(SettingsManager settings, String[] lines, int line) {

        // Loop through all the lines until we hit another section
        int row;
        for (row = line + 1; row < lines.length && !lines[row].startsWith("["); row++) {

            // Check we have something to process
            if (lines[row].trim().isEmpty()) {
                continue;
            }

            // Get the key and value
            String key = lines[row].split("=", 2)[0].replace("\"", "");
            String value = lines[row].split("=", 2)[1].replace("\"", "");

            // Create a new Settings entry
            switch (key.toLowerCase()) {
                case "proxy":
                    settings.setProxyServer(value);
                    break;
                case "frequency":
                    settings.setFrequency(Integer.parseInt(value));
                    break;
                case "alwaysontop":
                    settings.setAlwaysOnTop(Boolean.parseBoolean(value));
                    break;

                case "background colour":
                    settings.setBackgroundColor(vb6ColorToJavaColor(value));
                    break;
                case "text colour":
                    settings.setNormalTextColor(vb6ColorToJavaColor(value));
                    break;
                case "up colour":
                    settings.setUpColor(vb6ColorToJavaColor(value));
                    break;
                case "down colour":
                    settings.setDownColor(vb6ColorToJavaColor(value));
                    break;
                case "up arrow colour":
                    settings.setUpArrowColor(vb6ColorToJavaColor(value));
                    break;
                case "down arrow colour":
                    settings.setDownArrowColor(vb6ColorToJavaColor(value));
                    break;

                case "font":
                    settings.setFontName(value);
                    break;
                case "bold":
                    settings.setFontBold(Boolean.parseBoolean(value));
                    break;
                case "italic":
                    settings.setFontItalic(Boolean.parseBoolean(value));
                    break;
                case "font size":
                    float size = Float.parseFloat(value);
                    if (size < SettingsManager.FONT_SIZE_SMALL) {
                        size = SettingsManager.FONT_SIZE_SMALL;
                    }
                    else if (size < SettingsManager.FONT_SIZE_MEDIUM) {
                        size = SettingsManager.FONT_SIZE_MEDIUM;
                    }
                    else {
                        size = SettingsManager.FONT_SIZE_LARGE;
                    }
                    settings.setFontSize(size);
                    break;

                case "show total profit and loss":
                    settings.setShowPortfolioProfitAndLoss(Boolean.parseBoolean(value));
                    break;
                case "show total profit and loss as percent":
                    settings.setShowPortfolioProfitAndLossPercent(Boolean.parseBoolean(value));
                    break;
                case "show total cost":
                    settings.setShowTotalCost(Boolean.parseBoolean(value));
                    break;
                case "show total value":
                    settings.setShowTotalValue(Boolean.parseBoolean(value));
                    break;

                case "show daily change":
                    settings.setShowDailySummary(Boolean.parseBoolean(value));
                    break;
                case "summarise":
                    settings.setShowUniqueSymbols(Boolean.parseBoolean(value));
                    break;

                case "currency":
                    settings.setCurrencyCode(value);
                    break;
                case "currency symbol":
                    settings.setCurrencySymbol(value);
                    break;

                case "total investment":
                    settings.setTotalInvestment(Integer.parseInt(value));
                    break;
                case "margin":
                    settings.setMargin(Integer.parseInt(value));
                    break;

                case "day summary sort order":
                    settings.setDaySortOrder(value);
                    break;
                case "day summary sort column":
                    settings.setDaySortColumn(value);
                    break;
                case "summary sort order":
                    settings.setSummarySortOrder(value);
                    break;
                case "summary sort column":
                    settings.setSummarySortColumn(value);
                    break;

                case "high alarm wave file":
                    settings.setHighAlarmWaveFile(value);
                    break;
                case "low alarm wave file":
                    settings.setHighAlarmWaveFile(value);
                    break;

                case "alphavantage key":
                    settings.setAlphaVantageToken(value);
                    break;
                case "marketstack key":
                    settings.setMarketStackToken(value);
                    break;
                case "twelvedata key":
                    settings.setTwelveDataToken(value);
                    break;
                case "finhub key":
                    settings.setFinhubToken(value);
                    break;
                case "tiingo key":
                    settings.setTiingoToken(value);
                    break;
                case "freecurrency key":
                    settings.setFreeCurrencyToken(value);
                    break;

                case "scroll speed":
                    int speed = Integer.parseInt(value.replaceFirst("^[^,]+,", ""));
                    if (speed < SettingsManager.SCROLL_SPEED_SLOW) {
                        speed = SettingsManager.SCROLL_SPEED_SLOW;
                    }
                    else if (speed < SettingsManager.SCROLL_SPEED_MEDIUM) {
                        speed = SettingsManager.SCROLL_SPEED_MEDIUM;
                    }
                    else {
                        speed = SettingsManager.SCROLL_SPEED_FAST;
                    }
                    settings.setTickerSpeed(speed);
                    break;
            }
        }
        return row - 1;
    }

    /**
     * Loads a symbol from registry lines starting at the given line index.
     *
     * @param symbolsManager The SymbolsManager to add the symbol to.
     * @param pricesManager  The PricesManager to add prices to.
     * @param ratesManager   The ExchangeRatesManager to add exchange rates to.
     * @param lines          The registry lines.
     * @param line           The starting line index.
     * @return The next line index after processing.
     */
    private static int loadSymbolFromRegistryLines(SymbolsManager symbolsManager, PricesManager pricesManager, ExchangeRatesManager ratesManager, String[] lines, int line) {

        // Create a new symbol instance
        String timestamp = lines[line].replaceAll("(^.+\\\\)|(])", "");

        // We need to convert this timestamp to a unix version for the key
        // this timestamp is actually the number of seconds since 1st Jan 2008
        timestamp = (Long.parseLong(timestamp) + 1199145600L) * 1000 + ""; // Convert to milliseconds
        SymbolTransaction symbol = symbolsManager.createNewSymbolTransaction(timestamp);

        // Loop through all the lines until we hit another section
        int row;
        for (row = line + 1; row < lines.length && !lines[row].startsWith("["); row++) {

            // Check we have something to process
            if (lines[row].trim().isEmpty()) {
                continue;
            }

            // Get the key and value
            String key = lines[row].split("=", 2)[0].replace("\"", "");
            String value = lines[row].split("=", 2)[1].replace("\"", "");

            // Create a new SymbolTransaction entry
            switch (key.toLowerCase()) {
                case "symbol":
                    symbol.setCode(value);
                    break;
                case "alias":
                    symbol.setAlias(value);
                    break;
                case "disabled":
                    symbol.setDisabled(Boolean.parseBoolean(value));
                    break;
                case "price":
                    symbol.setPricePaid(Double.parseDouble(value));
                    break;
                case "currency":
                    symbol.setCurrencyCode(value);
                    break;
                case "currency symbol":
                    symbol.setCurrencySymbol(value);
                    break;
                case "shares":
                    symbol.setSharesBought(Double.parseDouble(value));
                    break;

                case "show price":
                    symbol.setShowPrice(Boolean.parseBoolean(value));
                    break;
                case "exclude from summary":
                    symbol.setExcludeFromSummary(Boolean.parseBoolean(value));
                    break;
                case "show change":
                    symbol.setShowChange(Boolean.parseBoolean(value));
                    break;
                case "show change percent":
                    symbol.setShowChangePercent(Boolean.parseBoolean(value));
                    break;
                case "show change indicator":
                    symbol.setShowChangeUpDown(Boolean.parseBoolean(value));
                    break;
                case "show profit and loss":
                    symbol.setShowProfitLoss(Boolean.parseBoolean(value));
                    break;

                case "show day change":
                    symbol.setShowDayChange(Boolean.parseBoolean(value));
                    break;
                case "show day change percent":
                    symbol.setShowDayChangePercent(Boolean.parseBoolean(value));
                    break;
                case "show day change indicator":
                    symbol.setShowDayChangeUpDown(Boolean.parseBoolean(value));
                    break;

                case "low alarm enabled":
                    symbol.setLowAlarmEnabled(Boolean.parseBoolean(value));
                    break;
                case "low alarm as percent":
                    symbol.setLowAlarmIsPercent(Boolean.parseBoolean(value));
                    break;
                case "low alarm sound enabled":
                    symbol.setLowAlarmSoundEnabled(Boolean.parseBoolean(value));
                    break;
                case "low alarm value":
                    symbol.setLowAlarmValue(Double.parseDouble(value));
                    break;

                case "high alarm enabled":
                    symbol.setHighAlarmEnabled(Boolean.parseBoolean(value));
                    break;
                case "high alarm as percent":
                    symbol.setHighAlarmIsPercent(Boolean.parseBoolean(value));
                    break;
                case "high alarm sound enabled":
                    symbol.setHighAlarmSoundEnabled(Boolean.parseBoolean(value));
                    break;
                case "high alarm value":
                    symbol.setHighAlarmValue(Double.parseDouble(value));
                    break;
            }
        }

        // Save the symbol
        log.info("Added symbol from registry: {}", symbol);
        pricesManager.addPrice(symbol.getCode());
        ratesManager.addRate(symbol.getCurrencySymbol());
        return row - 1;
    }

    /**
     * Converts a VB6 OLE color (BBGGRR) to a java.awt.Color.
     */
    private static Color vb6ColorToJavaColor(String vb6ColorString) {
        int vb6Color = Integer.parseInt(vb6ColorString);
        int r = vb6Color & 0xFF;
        int g = (vb6Color >> 8) & 0xFF;
        int b = (vb6Color >> 16) & 0xFF;
        return new Color(r, g, b);
    }

    /**
     * Custom JFileChooser that prompts for overwrite confirmation
     */
    @Getter
    @Setter
    public static class OverwritePromptChooser extends JFileChooser {
        public enum CHOOSER_TYPE {
            OPEN,
            SAVE
        }

        private CHOOSER_TYPE chooserType = CHOOSER_TYPE.OPEN;
        private JDialog dialog = null;
        private int returnValue = ERROR_OPTION;

        @Override
        protected JDialog createDialog(Component parent) throws HeadlessException {
            JDialog dialog = super.createDialog(parent);
            dialog.setLocationRelativeTo(parent);
            Utils.recenterDialog(dialog, parent);
            return dialog;
        }

        @Override
        public void approveSelection() {
            File file = getSelectedFile();

            // If we are saving, check if the file exists
            if (chooserType == CHOOSER_TYPE.SAVE) {
                if (file != null && file.exists()) {
                    int answer = JOptionPane.showConfirmDialog(
                            this,
                            "File \"" + file.getName() + "\" already exists.\n\nOverwrite?",
                            "Confirm Overwrite",
                            JOptionPane.YES_NO_OPTION,
                            JOptionPane.QUESTION_MESSAGE);
                    if (answer != JOptionPane.YES_OPTION) {
                        return;
                    }
                }
            }

            // If we are loading, then confirm overwrite of settings
            else {
                int answer = JOptionPane.showConfirmDialog(
                        this,
                        "You are about to overwrite your settings from \"" + file.getName() + "\".\nAre you sure?",
                        "Confirm Overwrite",
                        JOptionPane.YES_NO_OPTION,
                        JOptionPane.QUESTION_MESSAGE);
                if (answer != JOptionPane.YES_OPTION) {
                    return;
                }
            }
            super.approveSelection();
        }
    }
}
