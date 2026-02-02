/*
 *
 * Copyright (c) 2026, 4NG and/or its affiliates. All rights reserved.
 * 4NG PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.service;

import com.pivotal.stockticker.Utils;
import com.pivotal.stockticker.model.LivePrice;
import com.pivotal.stockticker.model.Price;
import com.pivotal.stockticker.model.SettingsManager;
import com.pivotal.stockticker.model.SymbolTransaction;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

/**
 * Handles exporting data in human readable formats
 */
@Slf4j
public class ExportData {

    // File chooser for backup/restore operations
    private static BackupReload.OverwritePromptChooser chooser = null;

    /**
     * Selects the file to export data to in CSV format
     *
     * @param dialog The parent dialog for the file chooser.
     */
    private static File getCsvExportFile(Window dialog) {

        // Use JFileChooser to let user pick the file location
        if (chooser == null) {
            chooser = new BackupReload.OverwritePromptChooser();
            chooser.setFileFilter(new FileNameExtensionFilter("Commas Separated Files (*.csv)", "csv"));
        }
        chooser.setChooserType(BackupReload.OverwritePromptChooser.CHOOSER_TYPE.SAVE);
        chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        chooser.setAcceptAllFileFilterUsed(false);
        chooser.setMultiSelectionEnabled(false);
        chooser.setDialogType(JFileChooser.SAVE_DIALOG);
        chooser.setDialogTitle("Save Data");
        int userSelection = chooser.showDialog(dialog, "Save");

        // If user approved, export the preferences to the selected file
        if (userSelection == JFileChooser.APPROVE_OPTION) {

            // Get the file and make sure it has an extension
            File selectedFile = chooser.getSelectedFile();
            if (!selectedFile.getName().toLowerCase().endsWith(".csv")) {
                selectedFile = new File(selectedFile.getAbsolutePath() + ".csv");
            }

            return selectedFile;
        }
        return null;
    }

    /**
     * Exports symbol data to a CSV file
     *
     * @param parent          The parent dialog for the file chooser
     * @param includeDisabled Whether to include disabled symbols in the export
     * @param uniqueOnly      Whether to include only unique symbols in the export
     */
    public static void exportSymbolsToCSV(Window parent, boolean includeDisabled, boolean uniqueOnly) {

        // Get a file to use for export
        File file = getCsvExportFile(parent);
        if (file != null) {

            // Get all the symbols from storage
            SymbolsManager symbolsManager = new SymbolsManager();

            // Select the symbols to export
            List<SymbolTransaction> symbols = symbolsManager.getSymbolTransactions(includeDisabled, uniqueOnly, null);
            PricesManager prices = new PricesManager(null);

            // Map live prices by symbol code for easy lookup
            try (FileWriter writer = new FileWriter(file);
                 CSVPrinter csvPrinter = new CSVPrinter(writer, CSVFormat.DEFAULT)) {

                // Write header
                csvPrinter.printRecord("Date:", LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")));
                csvPrinter.printRecord("Code", "Display Name", "Trade Date", "Disabled", "Currency Name", "Currency Symbol", "Shares", "Cost", "Current Price");

                // If this is unique only, we need to filter the list to only one transaction per symbol
                if (uniqueOnly) {
                    SettingsManager settings = SettingsManager.getInstance();
                    ExchangeRatesManager rates = new ExchangeRatesManager(null);
                    ArrayList<LivePrice> livePricesList = LivePrice.getLivePrices(symbolsManager, prices, rates, settings);

                    // Write data rows using live, aggregated prices
                    for (LivePrice symbol : livePricesList) {
                        csvPrinter.printRecord(
                                symbol.getSymbol(),
                                symbol.getSymbolTransaction().getDisplayName(),
                                symbol.getDisplayTimestamp(),
                                Boolean.toString(symbol.getSymbolTransaction().isDisabled()),
                                symbol.getSymbolTransaction().getCurrencyCode(),
                                symbol.getSymbolTransaction().getCurrencySymbol(),
                                Double.toString(symbol.getSharesBought()),
                                Double.toString(symbol.getPricePaid()),
                                Double.toString(symbol.getPrice())
                        );
                    }
                }
                else {
                    // Write data rows direct from storage
                    for (SymbolTransaction symbol : symbols) {
                        Price price = prices.getPrice(symbol.getCode());
                        csvPrinter.printRecord(
                                symbol.getCode(),
                                symbol.getDisplayName(),
                                symbol.getDisplayTimestamp(),
                                Boolean.toString(symbol.isDisabled()),
                                symbol.getCurrencyCode(),
                                symbol.getCurrencySymbol(),
                                Double.toString(symbol.getSharesBought()),
                                Double.toString(symbol.getPricePaid()),
                                Double.toString(price == null ? 0.0 : price.getCurrentPrice())
                        );
                    }
                }
                Utils.showTopmostMessage(parent, String.format("Successfully exported symbols to [%s]", file.getAbsolutePath()), "Export Successful", JOptionPane.INFORMATION_MESSAGE);
            }
            catch (IOException e) {
                log.error("Error exporting symbols to [{}] - {}", file.getAbsolutePath(), e.getMessage());
                Utils.showTopmostMessage(parent, String.format("Error exporting symbols to [%s]: %s", file.getAbsolutePath(), e.getMessage()), "Export Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }
}
