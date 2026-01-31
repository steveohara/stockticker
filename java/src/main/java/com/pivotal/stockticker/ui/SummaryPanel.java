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
 * Form showing the portfolio summary
 */
@Slf4j
public class SummaryPanel extends JDialog implements CallbackInterface {

    private static final int LEFT_MARGIN = 10;
    private static final int VALUE_SEP = 2;

    private ColouredTextPanel pnlSummary, pnlTotals;
    private SettingsLabel lblStock, lblPaid, lblPrice, lblShares, lblCost, lblValue, lblPercent, lblGainLoss, lblSource, lblHeader, lblDivider;
    private final TickerBar tickerBar;
    private SettingsLabel lblSort = null;
    private boolean sortAsc = true;

    /**
     * Creates new form
     *
     * @param tickerBar Parent callback interface
     */
    public SummaryPanel(TickerBar tickerBar) {
        this.tickerBar = tickerBar;
        initComponents();
        initListeners();

        // Set a timer to keep track of the mouse position for whether to close the
        // window automatically
        Timer timer = new Timer(750, e -> {

            // If we are not visible, do nothing
            if (!isVisible()) {
                return;
            }

            // If the mouse is not over the dialog and not over the ticker bar at
            // the summary, hide the panel
            Point mousePos = MouseInfo.getPointerInfo().getLocation();

            Point screenLocation = getLocationOnScreen();
            Rectangle windowBounds = new Rectangle(screenLocation.x, screenLocation.y, getWidth(), getHeight());

            screenLocation = tickerBar.pnlSummary.getLocationOnScreen();
            Rectangle summaryBounds = new Rectangle(screenLocation.x, screenLocation.y, tickerBar.pnlSummary.getWidth(), tickerBar.pnlSummary.getHeight());

            if (!summaryBounds.contains(mousePos) && !windowBounds.contains(mousePos)) {
                setVisible(false);
            }
        });
        timer.start();
    }

    /**
     * Sets up component listeners
     */
    private void initListeners() {

        // Add listeners to the column headers for sorting
        lblStock.addMouseListener(new MouseAdapter() {
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
                settings.setSummarySortColumn(lblSort.getText());
                settings.setSummarySortOrder(sortAsc ? "ASC" : "DESC");

                // Re-display the summary
                column.setText(lblSort.getText() + (sortAsc ? " ▲" : " ▼"));
                SwingUtilities.invokeLater(() -> displaySummary());
            }
        });
        lblPaid.addMouseListener(lblStock.getMouseListeners()[0]);
        lblPaid.addMouseListener(lblStock.getMouseListeners()[0]);
        lblPrice.addMouseListener(lblStock.getMouseListeners()[0]);
        lblShares.addMouseListener(lblStock.getMouseListeners()[0]);
        lblCost.addMouseListener(lblStock.getMouseListeners()[0]);
        lblValue.addMouseListener(lblStock.getMouseListeners()[0]);
        lblPercent.addMouseListener(lblStock.getMouseListeners()[0]);
        lblGainLoss.addMouseListener(lblStock.getMouseListeners()[0]);
        lblSource.addMouseListener(lblStock.getMouseListeners()[0]);
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
        Rectangle screen = Utils.getAllScreensBounds();
        int x = location == null ? 50 : (int) location.getX();
        int y = tickerBar.getY() + tickerBar.getHeight();

        // Need to make sure the form is fully on screen
        if (x + getWidth() > screen.getWidth()) {
            x = (int) screen.getWidth() - getWidth();
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
        SummaryStats summaryStats = new SummaryStats(symbols, prices, rates);
        double totalValue = summaryStats.calculateTotalValue();
        double totalCost = summaryStats.calculateTotalCost();
        double cashCost = settings.getTotalInvestment() + settings.getMargin();
        double adjustedTotalValue = totalValue - cashCost;

        // Now draw the totals panel
        pnlTotals.cls();
        pnlTotals.setOpaque(true);
        pnlTotals.setBackColor(settings.getBackgroundColor());
        pnlTotals.setFontColor(settings.getLabelColor());
        pnlTotals.setFont(settings.getFont());

        pnlTotals.atTop(lblDivider.getBottom() + VALUE_SEP * 3);
        pnlTotals.cls();
        pnlTotals.setCurrentX(LEFT_MARGIN);

        pnlTotals.setFontColor(settings.getLabelColor());
        pnlTotals.print("Investment: ");
        pnlTotals.setFontColor(settings.getNormalTextColor());
        pnlTotals.print(Utils.formatCurrencyValue(cashCost, settings.getCurrencySymbol()));

        pnlTotals.setFontColor(settings.getLabelColor());
        pnlTotals.print("  Cost: ");
        pnlTotals.setFontColor(settings.getNormalTextColor());
        pnlTotals.print(Utils.formatCurrencyValue(totalCost, settings.getCurrencySymbol()));

        pnlTotals.setFontColor(settings.getLabelColor());
        pnlTotals.print("  Value: ");
        pnlTotals.setFontColor(settings.getNormalTextColor());
        pnlTotals.print(Utils.formatCurrencyValue(totalValue, settings.getCurrencySymbol()));

        pnlTotals.setCurrentY(pnlTotals.getCurrentY() + pnlTotals.getFontMetrics(pnlTotals.getFont()).getHeight() + VALUE_SEP);

        pnlTotals.setCurrentX(LEFT_MARGIN);
        pnlTotals.setFontColor(settings.getLabelColor());
        pnlTotals.print("Summary: ");
        pnlTotals.setFontColor(totalValue < totalCost ? settings.getDownColor() : totalValue > totalCost ? settings.getUpColor() : settings.getNormalTextColor());
        pnlTotals.print(Utils.formatCurrencyValue(totalValue - totalCost, settings.getCurrencySymbol()));

        pnlTotals.setFontColor(totalValue < 0 ? settings.getDownColor() : totalValue > 0 ? settings.getUpColor() : settings.getNormalTextColor());
        pnlTotals.print(String.format("  (%s)", Utils.formatCurrencyValue(adjustedTotalValue, settings.getCurrencySymbol())));
        pnlTotals.setFontColor(totalValue < totalCost ? settings.getDownColor() : totalValue > totalCost ? settings.getUpColor() : settings.getNormalTextColor());

        pnlTotals.print(String.format("  %.2f%%", totalValue == 0.0 ? 0 : (totalValue - totalCost) / totalCost * 100));
    }

    /**
     * Draws the summary data panel
     */
    private void drawSummaryData(SettingsManager settings, ArrayList<LivePrice> livePricesList) {

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

            pnlSummary.setFontColor(settings.getNormalTextColor());
            pnlSummary.setCurrentX(lblPaid.getX());
            pnlSummary.print(livePrice.getFormattedPricePaid());

            pnlSummary.setFontColor(livePrice.isUp() ? settings.getUpColor() : (livePrice.isDown() ? settings.getDownColor() : settings.getNormalTextColor()));
            pnlSummary.setCurrentX(lblPrice.getX());
            pnlSummary.print(livePrice.getFormattedPrice());

            pnlSummary.setFontColor(settings.getNormalTextColor());
            pnlSummary.setCurrentX(lblShares.getX());
            pnlSummary.print(livePrice.getFormattedSharesBought());

            pnlSummary.setCurrentX(lblCost.getX());
            pnlSummary.print(livePrice.getFormattedCostLocal());

            pnlSummary.setFontColor(livePrice.isUp() ? settings.getUpColor() : (livePrice.isDown() ? settings.getDownColor() : settings.getNormalTextColor()));
            pnlSummary.setCurrentX(lblValue.getX());
            pnlSummary.print(livePrice.getFormattedValueLocal());

            pnlSummary.setCurrentX(lblPercent.getX());
            pnlSummary.print(livePrice.getFormattedPercentChange());

            pnlSummary.setCurrentX(lblGainLoss.getX());
            pnlSummary.print(livePrice.getFormattedProfitLossLocal());

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
        ArrayList<LivePrice> livePricesList = LivePrice.getLivePrices(symbols, prices, rates, true);

        // Sort them based on the current sort column and order
        livePricesList.sort((lp1, lp2) -> {
            int result = 0;
            if (lblSort.equals(lblStock)) {
                result = lp1.getSymbol().compareToIgnoreCase(lp2.getSymbol());
            }
            else if (lblSort.equals(lblPaid)) {
                result = Double.compare(lp1.getSymbolTransaction().getPricePaid(), lp2.getSymbolTransaction().getPricePaid());
            }
            else if (lblSort.equals(lblPrice)) {
                result = Double.compare(prices.getPrice(lp1.getSymbol()).getCurrentPrice(), prices.getPrice(lp2.getSymbol()).getCurrentPrice());
            }
            else if (lblSort.equals(lblShares)) {
                result = Integer.compare(lp1.getSharesBought(), lp2.getSharesBought());
            }
            else if (lblSort.equals(lblCost)) {
                result = Double.compare(lp1.getCostLocal(), lp2.getCostLocal());
            }
            else if (lblSort.equals(lblValue)) {
                result = Double.compare(lp1.getValueLocal(), lp2.getValueLocal());
            }
            else if (lblSort.equals(lblPercent)) {
                result = Double.compare(lp1.getPercentChange(), lp2.getPercentChange());
            }
            else if (lblSort.equals(lblGainLoss)) {
                result = Double.compare(lp1.getProfitLoss(), lp2.getProfitLoss());
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
        lblStock = SettingsLabel.create("Stock", "Sort by stock name").withDimensions(100, 25).atLeft(LEFT_MARGIN).atTop(LEFT_MARGIN / 2).setAlignment(SwingConstants.LEFT).setForeColor(settings.getLabelColor()).setBackColor(settings.getBackgroundColor()).to(getContentPane());
        lblPaid = SettingsLabel.create("Paid", "Sort by the cost base price of the stock").withDimensions(65, lblStock.getHeight()).tail(lblStock, 0).setAlignment(SwingConstants.LEFT).setForeColor(settings.getLabelColor()).setBackColor(settings.getBackgroundColor()).to(getContentPane());
        lblPrice = SettingsLabel.create("Price", "Sort by the current price").withDimensions(lblPaid).tail(lblPaid, 0).setAlignment(SwingConstants.LEFT).setForeColor(settings.getLabelColor()).setBackColor(settings.getBackgroundColor()).to(getContentPane());
        lblShares = SettingsLabel.create("Shares", "Sort by the number of shares").withDimensions(lblPaid).tail(lblPrice, 0).setAlignment(SwingConstants.LEFT).setForeColor(settings.getLabelColor()).setBackColor(settings.getBackgroundColor()).to(getContentPane());
        lblCost = SettingsLabel.create("Cost", "Sort by the total cost in local currency").withDimensions(80, lblStock.getHeight()).tail(lblShares, 0).setAlignment(SwingConstants.LEFT).setForeColor(settings.getLabelColor()).setBackColor(settings.getBackgroundColor()).to(getContentPane());
        lblValue = SettingsLabel.create("Value", "Sort by the total current value").withDimensions(lblCost).tail(lblCost, 0).setAlignment(SwingConstants.LEFT).setForeColor(settings.getLabelColor()).setBackColor(settings.getBackgroundColor()).to(getContentPane());
        lblPercent = SettingsLabel.create("Percent", "Sort by the percentage difference between value and cost").withDimensions(lblPaid).tail(lblValue, 0).setAlignment(SwingConstants.LEFT).setForeColor(settings.getLabelColor()).setBackColor(settings.getBackgroundColor()).to(getContentPane());
        lblGainLoss = SettingsLabel.create("Gain/Loss", "Sort by the difference between current total value and original cost").withDimensions(lblCost).tail(lblPercent, 0).setAlignment(SwingConstants.LEFT).setForeColor(settings.getLabelColor()).setBackColor(settings.getBackgroundColor()).to(getContentPane());
        lblSource = SettingsLabel.create("Source", "Sort by the source of the data").withDimensions(lblCost).tail(lblGainLoss, 0).setAlignment(SwingConstants.LEFT).setForeColor(settings.getLabelColor()).setBackColor(settings.getBackgroundColor()).to(getContentPane());
        lblHeader = SettingsLabel.create("").below(lblStock, 0).withDimensions(lblSource.getRight() - lblStock.getX(), 1).setForeColor(settings.getLabelColor()).setBackColor(settings.getLabelColor()).to(getContentPane());

        // Now position the summary panel
        pnlSummary = ColouredTextPanel.create().withDimensions(lblHeader.getWidth(), 10).at(0, lblHeader.getBottom() + VALUE_SEP).to(getContentPane());
        pnlSummary.setBackground(settings.getBackgroundColor());
        pnlSummary.setDisplayStyle(ColouredTextPanel.DISPLAY_STYLE.FIT);

        // Divider line
        lblDivider = SettingsLabel.create("").below(pnlSummary, 0).withDimensions(lblSource.getRight() - lblStock.getX(), 1).setForeColor(settings.getLabelColor()).setBackColor(settings.getLabelColor()).to(getContentPane());

        // Now position the totals panel
        pnlTotals = ColouredTextPanel.create().withDimensions(pnlSummary).below(lblDivider, 0).to(getContentPane());
        pnlTotals.setBackground(settings.getBackgroundColor());
        pnlTotals.setDisplayStyle(ColouredTextPanel.DISPLAY_STYLE.FIT);

        // Set the current sort column and order
        if (settings.getSummarySortColumn() != null && !settings.getSummarySortColumn().isEmpty()) {
            lblSort = switch (settings.getSummarySortColumn().toLowerCase()) {
                case "stock" -> lblStock;
                case "paid" -> lblPaid;
                case "price" -> lblPrice;
                case "shares" -> lblShares;
                case "cost" -> lblCost;
                case "value" -> lblValue;
                case "percent" -> lblPercent;
                case "gain/loss" -> lblGainLoss;
                case "source" -> lblSource;
                default -> lblStock;
            };
        }
        sortAsc = settings.getSummarySortOrder().equalsIgnoreCase("ASC");
        lblSort.setText(lblSort.getText() + (sortAsc ? " ▲" : " ▼"));

        // Set the size of the dialog
        setSize(pnlSummary.getRight(), pnlSummary.getBottom());
        setPreferredSize(getSize());
        setMaximumSize(getSize());
        setMinimumSize(getSize());
        setType(Type.UTILITY);
        setUndecorated(true);
    }

}
