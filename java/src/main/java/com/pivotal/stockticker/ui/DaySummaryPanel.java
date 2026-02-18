/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.Utils;
import com.pivotal.stockticker.model.LivePrice;
import com.pivotal.stockticker.model.SettingsManager;
import com.pivotal.stockticker.model.SummaryStats;
import com.pivotal.stockticker.service.ExchangeRatesManager;
import com.pivotal.stockticker.service.PricesManager;
import com.pivotal.stockticker.service.SymbolsManager;
import com.pivotal.stockticker.ui.components.ColouredTextPanel;
import com.pivotal.stockticker.ui.components.SettingsLabel;
import com.pivotal.stockticker.utils.CallbackInterface;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.util.ArrayList;

/**
 * Form showing the portfolio day summary
 */
@Slf4j
public class DaySummaryPanel extends JDialog implements CallbackInterface {

    private static final int LEFT_MARGIN = 10;
    private static final int VALUE_SEP = 2;

    private ColouredTextPanel pnlSummary, pnlTotals;
    private SettingsLabel lblStock, lblPrice, lblValue, lblGainLoss, lblSource, lblHeader, lblDivider;
    private final TickerBar tickerBar;
    private SettingsLabel lblSort = null;
    private boolean sortAsc = true;

    /**
     * Creates new form
     *
     * @param tickerBar Parent callback interface
     */
    public DaySummaryPanel(TickerBar tickerBar) {
        this.tickerBar = tickerBar;
        initComponents();
        initListeners();

        // Set a timer to keep track of the mouse position for whether to close the
        // window automatically
        Timer timer = new Timer(750, e -> {
            try {

                // If we are not visible, do nothing
                if (!isVisible()) {
                    return;
                }

                // If the popup menu is open, do not hide the preview as the user is likely trying to click on it
                if (tickerBar.getPopupMenu().isVisible()) {
                    setVisible(false);
                    return;
                }

                // If the mouse is not over the dialog and not over the ticker bar at
                // the summary, hide the panel
                Point mousePos = MouseInfo.getPointerInfo().getLocation();

                Point screenLocation = getLocationOnScreen();
                Rectangle windowBounds = new Rectangle(screenLocation.x, screenLocation.y, getWidth(), getHeight());

                screenLocation = tickerBar.pnlDaySummary.getLocationOnScreen();
                Rectangle summaryBounds = new Rectangle(screenLocation.x, screenLocation.y, tickerBar.pnlDaySummary.getWidth(), tickerBar.pnlSummary.getHeight());

                if (!summaryBounds.contains(mousePos) && !windowBounds.contains(mousePos)) {
                    setVisible(false);
                }
            }
            catch (Exception ex) {
                log.error("Error in day summary timer: ", ex);
            }
        });
        timer.start();
    }

    /**
     * Sets up component listeners
     */
    private void initListeners() {
        MouseAdapter mouseAdapter = new ColumnHeaderMouseAdapter();
        lblStock.addMouseListener(mouseAdapter);
        lblPrice.addMouseListener(mouseAdapter);
        lblValue.addMouseListener(mouseAdapter);
        lblGainLoss.addMouseListener(mouseAdapter);
        lblSource.addMouseListener(mouseAdapter);
    }

    /**
     * Shows the summary at the given location
     *
     * @param location Location to show the preview at
     */
    public void showSummary(Point location) {

        // Only do something if we're not visible
        if (isVisible()) {
            return;
        }

        // Position the form near the ticker bar and the summary
        Rectangle screen = Utils.getScreensBounds(location);
        int x = location == null ? 50 : (int) location.getX();
        int y = tickerBar.getY() + tickerBar.getHeight();

        // Need to make sure the form is fully on screen
        if (x + getWidth() > (screen.getX() + screen.getWidth())) {
            x = (int)screen.getX() + (int)screen.getWidth() - getWidth();
        }
        if (x < 0) {
            x = 0;
        }
        if (y + getHeight() > screen.getHeight()) {
            y = tickerBar.getY() - getHeight();
        }
        if (y < 0) {
            y = 0;
        }
        setLocation(new Point(x, y));
        setVisible(true);

        // Show the graph and summary
        displaySummary();
    }

    /**
     * Displays the summary information for the current live price
     */
    private void displaySummary() {

        // Get all the live prices
        SettingsManager settings = SettingsManager.getInstance();
        SymbolsManager symbols = new SymbolsManager();
        PricesManager prices = new PricesManager(this);
        ExchangeRatesManager rates = new ExchangeRatesManager(this);

        // Get the live prices in the correct order
        ArrayList<LivePrice> livePricesList = getLivePrices(symbols, prices, rates);

        // Draw the summary data
        drawSummaryData(settings, livePricesList);

        // Position the summary panel
        lblHeader.withWidth(pnlSummary);
        lblDivider.withDimensions(lblHeader).at(lblHeader.getX(), pnlSummary.getBottom() + VALUE_SEP * 2);

        // Get the summary stats
        drawTotalsData(symbols, prices, rates, settings);

        // Resize the dialog to fit the data
        setSize(pnlSummary.getRight() + LEFT_MARGIN * 2, pnlTotals.getBottom() + LEFT_MARGIN);
        setPreferredSize(getSize());
        setMaximumSize(getSize());
        setMinimumSize(getSize());
    }

    /**
     * Draws the totals data panel
     */
    private void drawTotalsData(SymbolsManager symbols, PricesManager prices, ExchangeRatesManager rates, SettingsManager settings) {
        // Create a summary stats object to calculate the summary data
        SummaryStats summaryStats = new SummaryStats(symbols, prices, rates);
        double totalCostAtPreviousClose = summaryStats.calculateTotalValueAtPreviousClose();
        double totalValue = summaryStats.calculateTotalValue();

        // Now draw the totals panel
        pnlTotals.cls();
        pnlTotals.setOpaque(true);
        pnlTotals.setBackColor(settings.getBackgroundColor());
        pnlTotals.setFontColor(settings.getLabelColor());
        pnlTotals.setFont(settings.getFont());

        pnlTotals.atTop(lblDivider.getBottom() + VALUE_SEP * 3);
        pnlTotals.cls();
        pnlTotals.setCurrentX(LEFT_MARGIN);

        // Draw the summary data
        pnlTotals.print("Today: ");
        pnlTotals.setFontColor(totalValue < totalCostAtPreviousClose ? settings.getDownColor() : totalValue > totalCostAtPreviousClose ? settings.getUpColor() : settings.getNormalTextColor());
        pnlTotals.print(String.format("%s", Utils.formatCurrencyValue(totalCostAtPreviousClose == 0.0 ? 0 : (totalValue - totalCostAtPreviousClose), settings.getCurrencySymbol())));
        pnlTotals.setCurrentX(pnlTotals.getCurrentX());

        // Percentage change
        pnlTotals.print(String.format("  (%.2f%%)", totalCostAtPreviousClose == 0.0 ? 0 : (totalValue - totalCostAtPreviousClose) / totalCostAtPreviousClose * 100));
    }

    /**
     * Draws the summary data panel
     */
    private void drawSummaryData(SettingsManager settings, ArrayList<LivePrice> livePricesList) {

        // Update the label colours
        lblStock.setForeColor(settings.getPreviewColor()).setBackColor(settings.getBackgroundColor());
        lblPrice.setForeColor(settings.getPreviewColor()).setBackColor(settings.getBackgroundColor());
        lblValue.setForeColor(settings.getPreviewColor()).setBackColor(settings.getBackgroundColor());
        lblGainLoss.setForeColor(settings.getPreviewColor()).setBackColor(settings.getBackgroundColor());
        lblSource.setForeColor(settings.getPreviewColor()).setBackColor(settings.getBackgroundColor());
        lblHeader.setForeColor(settings.getPreviewColor()).setBackColor(settings.getPreviewColor());
        lblDivider.setForeColor(settings.getPreviewColor()).setBackColor(settings.getPreviewColor());

        // Clear the summary panel
        pnlSummary.cls();
        pnlSummary.setOpaque(true);
        pnlSummary.setBackColor(settings.getBackgroundColor());
        pnlSummary.setFontColor(settings.getLabelColor());
        pnlSummary.setFont(settings.getFont());
        pnlSummary.setContiguousBackground(true);
        Color altColor = Utils.lighten(settings.getBackgroundColor(), 0.2f);

        // Now draw all the live prices
        boolean altRow = false;
        pnlSummary.setCurrentY(VALUE_SEP);
        for (LivePrice livePrice : livePricesList) {
            pnlSummary.setBackColor(altRow ? altColor : settings.getBackgroundColor());
            pnlSummary.setFontColor(settings.getLabelColor());
            pnlSummary.setCurrentX(lblStock.getX());
            pnlSummary.print(livePrice.getSymbol());

            pnlSummary.setFontColor(livePrice.isUpToday() ? settings.getUpColor() : (livePrice.isDownToday() ? settings.getDownColor() : settings.getNormalTextColor()));
            pnlSummary.setCurrentX(lblPrice.getX());
            pnlSummary.print(livePrice.getFormattedPrice());

            pnlSummary.setFontColor(livePrice.isUpToday() ? settings.getUpColor() : (livePrice.isDownToday() ? settings.getDownColor() : settings.getNormalTextColor()));
            pnlSummary.setCurrentX(lblValue.getX());
            pnlSummary.print(livePrice.getFormattedDayProfitLossLocal());

            pnlSummary.setCurrentX(lblGainLoss.getX());
            pnlSummary.print(livePrice.getFormattedPercentDayChange());
            pnlSummary.setCurrentX(lblGainLoss.getX() + lblGainLoss.getWidth() / 2);
            pnlSummary.print("(" + livePrice.getFormattedDayChange() + ")");

            pnlSummary.setFontColor(settings.getLabelColor());
            pnlSummary.setCurrentX(lblSource.getX());
            pnlSummary.print(livePrice.getSource());
            pnlSummary.setCurrentY(pnlSummary.getCurrentY() + pnlSummary.getFontMetrics(pnlSummary.getFont()).getHeight() + VALUE_SEP);
            pnlSummary.setBackColor(settings.getBackgroundColor());
            altRow = !altRow;
        }
    }

    /**
     * Gets the live prices sorted based on the current sort column and order
     */
    private ArrayList<LivePrice> getLivePrices(SymbolsManager symbols, PricesManager prices, ExchangeRatesManager rates) {
        ArrayList<LivePrice> livePricesList = LivePrice.getLivePrices(symbols, prices, rates, true, true);

        // Sort them based on the current sort column and order
        livePricesList.sort((lp1, lp2) -> {
            int result = 0;
            if (lblSort.equals(lblStock)) {
                result = lp1.getSymbol().compareToIgnoreCase(lp2.getSymbol());
            }
            else if (lblSort.equals(lblPrice)) {
                result = Double.compare(lp1.getPriceLocal(), lp2.getPriceLocal());
            }
            else if (lblSort.equals(lblValue)) {
                result = Double.compare(lp1.getDayProfitLossLocal(), lp2.getDayProfitLossLocal());
            }
            else if (lblSort.equals(lblGainLoss)) {
                result = Double.compare(lp1.getPercentDayChange(), lp2.getPercentDayChange());
            }
            else if (lblSort.equals(lblSource)) {
                result = lp1.getSource().compareToIgnoreCase(lp2.getSource());
            }
            return sortAsc ? result : -result;
        });
        return livePricesList;
    }

    @Override
    public void changed(Object source) {
        if (isVisible()) {
            // We need to make sure that all UI changes are done on the Swing thread
            SwingUtilities.invokeLater(this::displaySummary);
        }
    }

    /**
     * Initializes all the UI components
     */
    private void initComponents() {

        setAlwaysOnTop(true);
        setResizable(false);
        setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
        ((JPanel)getContentPane()).setBorder(BorderFactory.createLineBorder(Color.DARK_GRAY));
        SettingsManager settings = SettingsManager.getInstance();
        setForeground(settings.getLabelColor());
        getContentPane().setBackground(settings.getBackgroundColor());

        setSize(100, 100);
        setPreferredSize(getSize());
        setMaximumSize(getSize());
        setMinimumSize(getSize());

        setResizable(false);
        getContentPane().setLayout(null);

        // Create all the column headers
        lblStock = SettingsLabel.create("Stock", "Sort by stock name").withDimensions(100, 25).atLeft(LEFT_MARGIN).atTop(LEFT_MARGIN / 2).setAlignment(SwingConstants.LEFT).setForeColor(settings.getPreviewColor()).setBackColor(settings.getBackgroundColor()).to(getContentPane());
        lblPrice = SettingsLabel.create("Price", "Sort by the current price").sameAs(lblStock).tail(lblStock).withWidth(65).to(getContentPane());
        lblValue = SettingsLabel.create("Value", "Sort by the value of the gain/loss today").sameAs(lblStock).withWidth(80).tail(lblPrice).to(getContentPane());
        lblGainLoss = SettingsLabel.create("Change", "Sort by the percentage change between the start of day and current price").sameAs(lblStock).withWidth(110).tail(lblValue).to(getContentPane());
        lblSource = SettingsLabel.create("Source", "Sort by the source of the data").sameAs(lblPrice).tail(lblGainLoss).to(getContentPane());
        lblHeader = SettingsLabel.create().below(lblStock, 0).withDimensions(lblSource.getRight() - lblStock.getX(), 1).setForeColor(settings.getLabelColor()).setBackColor(settings.getLabelColor()).to(getContentPane());

        // Now position the summary panel
        pnlSummary = ColouredTextPanel.create().withDimensions(lblHeader.getWidth(), 10).at(0, lblHeader.getBottom() + VALUE_SEP).to(getContentPane());
        pnlSummary.setBackground(settings.getBackgroundColor());
        pnlSummary.setDisplayStyle(ColouredTextPanel.DISPLAY_STYLE.FIT);

        // Divider line
        lblDivider = SettingsLabel.create().sameAs(lblHeader).below(pnlSummary).to(getContentPane());

        // Now position the totals panel
        pnlTotals = ColouredTextPanel.create().withDimensions(pnlSummary).below(lblDivider).to(getContentPane());
        pnlTotals.setBackground(settings.getBackgroundColor());
        pnlTotals.setDisplayStyle(ColouredTextPanel.DISPLAY_STYLE.FIT);

        // Set the current sort column and order
        if (settings.getDaySortColumn() != null && !settings.getDaySortColumn().isEmpty()) {
            lblSort = switch (settings.getDaySortColumn().toLowerCase()) {
                case "stock" -> lblStock;
                case "price" -> lblPrice;
                case "value" -> lblValue;
                case "gain/loss" -> lblGainLoss;
                case "source" -> lblSource;
                default -> lblStock;
            };
        }
        sortAsc = settings.getDaySortOrder().equalsIgnoreCase("ASC");
        lblSort.setText(lblSort.getText() + (sortAsc ? " ▲" : " ▼"));

        // Set the size of the dialog
        setSize(pnlSummary.getRight(), pnlSummary.getBottom());
        setPreferredSize(getSize());
        setMaximumSize(getSize());
        setMinimumSize(getSize());
        setType(Type.UTILITY);
        setUndecorated(true);
    }

    /**
     * Mouse adapter for the column headers
     */
    private class ColumnHeaderMouseAdapter extends MouseAdapter {
        @Override
        public void mouseClicked(MouseEvent e) {

            // Figure out the current sort column and order
            SettingsManager settings = SettingsManager.getInstance();
            SettingsLabel column = (SettingsLabel) e.getComponent();
            lblSort.setText(lblSort.getText().replaceAll("[^a-zA-Z0-9]", ""));

            // Reversing the sort
            if (lblSort.equals(column)) {
                sortAsc = !sortAsc;
            }

            // Changing the sort column
            else {
                lblSort = column;
                sortAsc = true;
            }

            // Save the settings
            settings.setDaySortColumn(lblSort.getText());
            settings.setDaySortOrder(sortAsc ? "ASC" : "DESC");

            // Re-display the summary
            column.setText(lblSort.getText() + (sortAsc ? " ▲" : " ▼"));
            SwingUtilities.invokeLater(DaySummaryPanel.this::displaySummary);
        }
    }

}
