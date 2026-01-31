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
import com.pivotal.stockticker.model.SymbolTransaction;
import com.pivotal.stockticker.service.ExchangeRatesManager;
import com.pivotal.stockticker.service.PricesManager;
import com.pivotal.stockticker.service.SymbolsManager;
import com.pivotal.stockticker.ui.components.ColouredTextPanel;
import com.pivotal.stockticker.utils.CallbackInterface;
import com.pivotal.stockticker.utils.StartupManager;
import com.pivotal.stockticker.utils.VersionInfo;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseMotionAdapter;
import java.net.URI;
import java.util.ArrayList;
import java.util.concurrent.CopyOnWriteArrayList;

/**
 * The main ticker bar UI class that displays stock prices and related information.
 */
@Slf4j
public class TickerBar extends JFrame implements CallbackInterface {

    public static final int STOCK_SEPARATION = 7;
    public static final int VALUE_SEPARATION = 4;
    public static final int UP_DOWN_SEPARATION = 1;

    private final SettingsManager settings = SettingsManager.getInstance();
    private final SymbolsManager symbols = new SymbolsManager();
    private final PricesManager prices = new PricesManager(this);
    private final ExchangeRatesManager rates = new ExchangeRatesManager(this);
    private final CopyOnWriteArrayList<LivePrice> livePrices = new CopyOnWriteArrayList<>();

    private final StockPanel stockPanel = new StockPanel(this);
    private final SummaryPanel summaryPanel = new SummaryPanel(this);

    private JPanel pnlLeftDrag;
    private JPanel pnlRightDrag;
    private ColouredTextPanel pnlStocks;
    protected ColouredTextPanel pnlSummary;
    protected ColouredTextPanel pnlDaySummary;
    private JPanel pnlTicker;

    private JCheckBoxMenuItem fontSizeItemSmall;
    private JCheckBoxMenuItem fontSizeItemMedium;
    private JCheckBoxMenuItem fontSizeItemLarge;
    private JCheckBoxMenuItem onTop;
    private JCheckBoxMenuItem scrollItemSlow;
    private JCheckBoxMenuItem scrollItemNormal;
    private JCheckBoxMenuItem scrollItemFast;

    private Point dragStart = null;
    private Point leftDragStart = null;
    private Point rightDragStart = null;
    private int left = 0;
    private int right = 0;

    /**
     * Constructor to initialize the ticker bar UI.
     *
     */
    public TickerBar() {
        createUIComponents();
        setupContextMenu();
        initializeUI();
        initListeners();

        // Add all the symbols to the prices manager
        prices.replacePrices(symbols.getAllSymbolCodes(false));

        // Add all the currencies to the exchange rates manager
        rates.replaceExchangeRates(symbols.getAllCurrencyCodes(false));

        // Draw the ticket now
        drawTickerContent();

        // Start the schedulers to update prices and exchange rates
        prices.startScheduler();
        rates.startScheduler();
    }

    /**
     * Draw the content of the ticker
     */
    synchronized private void drawTickerContent() {
        log.debug("Drawing ticker symbols");
        drawLivePrices();
        drawSummary();
        drawDaySummary();

    }

    /**
     * Draw the content of the ticker
     */
    private void drawLivePrices() {

        // Get a fresh list of live prices to work with
        ArrayList<LivePrice> livePricesList = LivePrice.getLivePrices(symbols, prices, rates, settings);
        log.debug("Drawing live prices for {} symbols", livePricesList.size());
        pnlStocks.cls();
        livePrices.clear();
        livePrices.addAll(livePricesList);

        // Draw these on the ticker panel
        int x = 0;
        pnlStocks.setCurrentX(VALUE_SEPARATION);
        for (LivePrice livePrice : livePrices) {

            // Draw the stock data
            drawLivePrice(livePrice);

            // Separate from the next stock
            pnlStocks.setCurrentX(pnlStocks.getCurrentX() + STOCK_SEPARATION);

            // Set the bounds for this live price
            livePrice.setBounds(new Rectangle(x, 0, pnlStocks.getCurrentX() - x, pnlStocks.getHeight()));
            x = pnlStocks.getCurrentX();
        }
        log.debug("Drawing live prices complete");
    }

    /**
     * Draws the summary panel on the ticker.
     */
    private void drawSummary() {
        log.debug("Drawing summary panel");
        pnlSummary.cls();
        if (settings.isShowSummary()) {
            pnlSummary.setBackground(settings.getBackgroundColor());
            pnlSummary.setFontColor(settings.getNormalTextColor());
            pnlSummary.setFont(settings.getFont());
            pnlSummary.setFontBold(settings.isFontBold());
            pnlSummary.setFontItalic(settings.isFontItalic());
            pnlSummary.setCurrentX(VALUE_SEPARATION);
            pnlSummary.print("");

            // Create a summary stats object to calculate the summary data
            SummaryStats summaryStats = new SummaryStats(symbols, prices, rates);
            double totalValue = summaryStats.calculateTotalValue();
            double totalCost = summaryStats.calculateTotalCost();
            double cashCost = settings.getTotalInvestment() + settings.getMargin();
            double adjustedTotalValue = totalValue - cashCost;

            // Draw the summary data
            pnlSummary.setFontColor(settings.getLabelColor());
            pnlSummary.print("Summary: ");
            pnlSummary.setFontColor(settings.getNormalTextColor());
            if (settings.isShowPortfolioProfitAndLoss()) {
                pnlSummary.setFontColor(totalValue < totalCost ? settings.getDownColor() : totalValue > totalCost ? settings.getUpColor() : settings.getNormalTextColor());
                pnlSummary.print(Utils.formatCurrencyValue(totalValue - totalCost, settings.getCurrencySymbol()));
                pnlSummary.setCurrentX(pnlSummary.getCurrentX() + VALUE_SEPARATION / 2);

                pnlSummary.setFontColor(totalValue < cashCost ? settings.getDownColor() : totalValue > cashCost ? settings.getUpColor() : settings.getNormalTextColor());
                pnlSummary.print(String.format("(%s)", Utils.formatCurrencyValue(adjustedTotalValue, settings.getCurrencySymbol())));
                pnlSummary.setCurrentX(pnlSummary.getCurrentX() + VALUE_SEPARATION);
            }
            if (settings.isShowPortfolioProfitAndLossPercent()) {
                pnlSummary.setFontColor(totalValue < totalCost ? settings.getDownColor() : totalValue > totalCost ? settings.getUpColor() : settings.getNormalTextColor());
                pnlSummary.print(String.format("%.2f%%", totalValue == 0.0 ? 0 : (totalValue - totalCost) / totalCost * 100));
                pnlSummary.setCurrentX(pnlSummary.getCurrentX() + VALUE_SEPARATION);
            }
            if (settings.isShowTotalCost()) {
                pnlSummary.setFontColor(settings.getNormalTextColor());
                pnlSummary.print(String.format("Cost:%s", Utils.formatCurrencyValue(totalCost, settings.getCurrencySymbol())));
                pnlSummary.setCurrentX(pnlSummary.getCurrentX() + VALUE_SEPARATION);
            }
            if (settings.isShowTotalValue()) {
                pnlSummary.setFontColor(settings.getNormalTextColor());
                pnlSummary.print(String.format("Val:%s", Utils.formatCurrencyValue(totalValue, settings.getCurrencySymbol())));
                pnlSummary.setCurrentX(pnlSummary.getCurrentX() + VALUE_SEPARATION);
            }
            pnlSummary.print("");
        }
    }

    /**
     * Draws the day summary panel on the ticker.
     */
    private void drawDaySummary() {
        log.debug("Drawing day summary panel");
        pnlDaySummary.cls();
        if (settings.isShowDailySummary()) {
            pnlDaySummary.setBackground(settings.getBackgroundColor());
            pnlDaySummary.setFontColor(settings.getNormalTextColor());
            pnlDaySummary.setFont(settings.getFont());
            pnlDaySummary.setCurrentX(VALUE_SEPARATION);
            pnlDaySummary.print("");

            // Create a summary stats object to calculate the summary data
            SummaryStats summaryStats = new SummaryStats(symbols, prices, rates);
            double totalCostAtPreviousClose = summaryStats.calculateTotalValueAtPreviousClose();
            double totalValue = summaryStats.calculateTotalValue();

            // Draw the summary data
            pnlDaySummary.setFontColor(settings.getLabelColor());
            pnlDaySummary.print("Today: ");
            pnlDaySummary.setFontColor(totalValue < totalCostAtPreviousClose ? settings.getDownColor() : totalValue > totalCostAtPreviousClose ? settings.getUpColor() : settings.getNormalTextColor());
            pnlDaySummary.print(String.format("%s", Utils.formatCurrencyValue(totalCostAtPreviousClose == 0.0 ? 0 : (totalValue - totalCostAtPreviousClose), settings.getCurrencySymbol())));
            pnlDaySummary.setCurrentX(pnlDaySummary.getCurrentX() + VALUE_SEPARATION);

            // Percentage change
            pnlDaySummary.print(String.format("(%.2f%%)", totalCostAtPreviousClose == 0.0 ? 0 : (totalValue - totalCostAtPreviousClose) / totalCostAtPreviousClose * 100));
            pnlDaySummary.setCurrentX(pnlDaySummary.getCurrentX() + VALUE_SEPARATION);
            pnlDaySummary.print("");
        }
    }

    /**
     * Draws a live price on the ticker.
     *
     * @param livePrice The live price data to draw.
     */
    private void drawLivePrice(LivePrice livePrice) {

        // If this symbol is selected, set the background colour
        pnlStocks.setBackground(livePrice.getSymbolTransaction().isSelected() ? settings.getHoverColor() : settings.getBackgroundColor());
        pnlStocks.setContiguousBackground(true);

        // Draw the price and other data
        pnlStocks.setCurrentY(1);
        boolean bShownOtherData = drawSymbolPrice(livePrice);

        // Draw the day changes
        drawSymbolDayChanges(livePrice, bShownOtherData);

        // Reset the background colour
        pnlStocks.setBackground(settings.getBackgroundColor());
        pnlStocks.setContiguousBackground(false);
    }

    /**
     * Draws the day changes for a symbol on the ticker.
     *
     * @param livePrice       The live price data for the symbol.
     * @param bShownOtherData Indicates if other data has already been shown for this symbol.
     */
    private void drawSymbolDayChanges(LivePrice livePrice, boolean bShownOtherData) {

        // Check if we need to show day changes
        SymbolTransaction symbol = livePrice.getSymbolTransaction();
        if ((symbol.isShowDayChange() || symbol.isShowDayChangePercent() || symbol.isShowDayChangeUpDown())) {
            boolean showBraces = (bShownOtherData && (symbol.isShowDayChange() || symbol.isShowDayChangePercent())) || (symbol.isShowChangeUpDown() && symbol.isShowDayChangeUpDown());
            pnlStocks.setFontColor(livePrice.isUpToday() ? settings.getUpColor() : livePrice.isDownToday() ? settings.getDownColor() : settings.getNormalTextColor());
            pnlStocks.setFontBold(settings.isFontBold());
            pnlStocks.setFontItalic(settings.isFontItalic());
            if (showBraces) {
                pnlStocks.setCurrentX(pnlStocks.getCurrentX() + VALUE_SEPARATION);
                pnlStocks.print("(");
            }

            // Show the Day price difference
            if (symbol.isShowDayChange()) {
                pnlStocks.print(livePrice.getFormattedDayChange());
                pnlStocks.setCurrentX(pnlStocks.getCurrentX() + VALUE_SEPARATION);
            }

            // Show the Day change in percent
            if (symbol.isShowDayChangePercent()) {
                pnlStocks.print(livePrice.getFormattedPercentDayChange());
                pnlStocks.setCurrentX(pnlStocks.getCurrentX() + VALUE_SEPARATION);
            }

            // Show the day profit/loss
            if (symbol.isShowProfitLoss() && symbol.isShowDayChange()) {
                pnlStocks.print(livePrice.getFormattedDayProfitLoss());
                pnlStocks.setCurrentX(pnlStocks.getCurrentX() + VALUE_SEPARATION);
            }

            // Show the Day up/down arrows
            if (symbol.isShowDayChangeUpDown()) {
                pnlStocks.setFontBold(true);
                pnlStocks.setCurrentX(pnlStocks.getCurrentX() - VALUE_SEPARATION + UP_DOWN_SEPARATION);
                pnlStocks.setFontColor(livePrice.isUpToday() ? settings.getUpArrowColor() : livePrice.isDownToday() ? settings.getDownArrowColor() : settings.getLabelColor());
                pnlStocks.print(livePrice.isUpToday() ? "↑" : livePrice.isDownToday() ? "↓" : "↕");
                pnlStocks.setFontBold((settings.getFontStyle() | Font.BOLD) > 0);
            }
            if (showBraces) {
                pnlStocks.setFontColor(livePrice.isUpToday() ? settings.getUpColor() : livePrice.isDownToday() ? settings.getDownColor() : settings.getNormalTextColor());
                pnlStocks.print(")");
            }
        }
    }

    /**
     * Draws the symbol price and related information on the ticker.
     *
     * @param livePrice The live price data for the symbol.
     * @return True if other data was shown, false otherwise.
     */
    private boolean drawSymbolPrice(LivePrice livePrice) {

        // Show the symbol
        SymbolTransaction symbol = livePrice.getSymbolTransaction();
        pnlStocks.setFontColor(livePrice.isUp() ? settings.getUpColor() : livePrice.isDown() ? settings.getDownColor() : settings.getNormalTextColor());
        pnlStocks.setFontBold(settings.isFontBold());
        pnlStocks.setFontItalic(settings.isFontItalic());
        pnlStocks.print(livePrice.getSymbolTransaction().getDisplayName());
        pnlStocks.setCurrentX(pnlStocks.getCurrentX() + VALUE_SEPARATION);

        // Show the price and other data
        boolean showData = symbol.isShowPrice() || symbol.isShowChangeUpDown() || symbol.isShowChange() || symbol.isShowChangePercent() || symbol.isShowProfitLoss();
        if (showData) {

            // Show the price
            if (symbol.isShowPrice()) {
                pnlStocks.setFontColor(livePrice.isUp() ? settings.getUpColor() : livePrice.isDown() ? settings.getDownColor() : settings.getNormalTextColor());
                pnlStocks.print(livePrice.getFormattedPrice());
                pnlStocks.setCurrentX(pnlStocks.getCurrentX() + VALUE_SEPARATION);
            }

            // Show the price difference
            if (symbol.isShowChange()) {
                pnlStocks.print(livePrice.getFormattedChange());
                pnlStocks.setCurrentX(pnlStocks.getCurrentX() + VALUE_SEPARATION);
            }

            // Show the change in percent
            if (symbol.isShowChangePercent()) {
                pnlStocks.print(livePrice.getFormattedPercentChange());
                pnlStocks.setCurrentX(pnlStocks.getCurrentX() + VALUE_SEPARATION);
            }

            // Show the profit/loss
            if (symbol.isShowProfitLoss()) {
                pnlStocks.print(livePrice.getFormattedProfitLoss());
                pnlStocks.setCurrentX(pnlStocks.getCurrentX() + VALUE_SEPARATION);
            }

            // Show the up/down arrows
            if (symbol.isShowChangeUpDown()) {
                pnlStocks.setFontBold(true);
                pnlStocks.setCurrentX(pnlStocks.getCurrentX() - VALUE_SEPARATION - UP_DOWN_SEPARATION);
                pnlStocks.setFontColor(livePrice.isUp() ? settings.getUpArrowColor() : livePrice.isDown() ? settings.getDownArrowColor() : settings.getLabelColor());
                pnlStocks.print(livePrice.isUp() ? "↑" : livePrice.isDown() ? "↓" : "↕");
                pnlStocks.setFontBold((settings.getFontStyle() | Font.BOLD) > 0);
            }
        }
        return showData;
    }

    @Override
    public void changed(Object source) {
        log.debug("TickerBar change notification received from source: {}", source == null ? "null" : source.getClass().getSimpleName());

        // Notify any preview windows about the change
        stockPanel.changed(source);
        summaryPanel.changed(source);

        // We need to make sure that all UI changes are done on the Swing thread
        SwingUtilities.invokeLater(() -> {

            // If there is no component, then this is a complete initialization
            switch (source) {
                case null -> {

                    // Load all the changed settings from storage
                    settings.loadFromStorage();
                    setTicketSpeed(settings.getTickerSpeed());
                    initializeUI();

                    // Load all symbols, prices and exchange rates from storage
                    rates.loadFromStorage(true);
                    symbols.loadFromStorage();
                    prices.loadFromStorage();

                    // Reset the schedulers to pick up any changes
                    prices.startScheduler();
                    rates.startScheduler();

                    // Draw the ticker content
                    drawTickerContent();
                }

                // If the settings form was the source, update settings
                case SettingsForm settingsForm -> {
                    setTicketSpeed(settings.getTickerSpeed());
                    initializeUI();
                    drawTickerContent();
                }

                // If the symbols form was the source, update the prices and redraw
                case SymbolsForm symbolsForm -> {
                    symbols.loadFromStorage();
                    prices.replacePrices(symbols.getAllSymbolCodes(false));
                    rates.replaceExchangeRates(symbols.getAllCurrencyCodes(false));

                    prices.startScheduler();
                    rates.startScheduler();
                    drawTickerContent();
                }
                default -> drawTickerContent();
            }
        });
    }

    /**
     * Initializes the UI components and layout.
     */
    private void initializeUI() {

        // Position and size the main frame
        setAlwaysOnTop(settings.isAlwaysOnTop());
        setTextCharacteristics();
        setLocation(settings.getWindowX(), settings.getWindowY());
        pnlTicker.setPreferredSize(new Dimension(settings.getWindowWidth(), getFontMetrics(getFont()).getHeight() + 1));
        pack();

        // Apply background color settings to all panels
        pnlTicker.setBackground(settings.getBackgroundColor());
        pnlLeftDrag.setBackground(pnlTicker.getBackground());
        pnlRightDrag.setBackground(pnlTicker.getBackground());
        pnlSummary.setBackground(pnlTicker.getBackground());
        pnlDaySummary.setBackground(pnlTicker.getBackground());
        pnlStocks.setBackground(pnlTicker.getBackground());
        setBackground(pnlTicker.getBackground());

        // Apply font settings to all panels
        pnlDaySummary.setFont(getFont());
        pnlSummary.setFont(getFont());
        pnlStocks.setFont(getFont());
        pnlSummary.setVisible(settings.isShowSummary());
        pnlDaySummary.setVisible(settings.isShowDailySummary());
        pnlStocks.setScrollSpeed(settings.getTickerSpeed());
        pnlDaySummary.setBackground(pnlTicker.getBackground());
        pnlSummary.setBackground(pnlTicker.getBackground());
        pnlStocks.setBackground(pnlTicker.getBackground());
        pnlDaySummary.setForeground(pnlTicker.getForeground());
        pnlSummary.setForeground(pnlTicker.getForeground());
        pnlStocks.setForeground(pnlTicker.getForeground());

        // Set the colors for the borders
        pnlSummary.setBorder(BorderFactory.createMatteBorder(0, 0, 0, 1, settings.getLabelColor()));
        pnlDaySummary.setBorder(BorderFactory.createMatteBorder(0, 0, 0, 1, settings.getLabelColor()));

        // Finalize and display the frame
        setVisible(true);
    }

    /**
     * Sets up dragging functionality for the ticker panel.
     */
    private void initListeners() {

        // Add mouse listeners for dragging the ticker and double clicking
        pnlTicker.addMouseListener(new MouseAdapter() {
            @Override
            public void mousePressed(MouseEvent e) {
                dragStart = e.getPoint();
            }

            @Override
            public void mouseReleased(MouseEvent e) {
                if (dragStart != null) {
                    settings.setWindowX(getX());
                    settings.setWindowY(getY());
                    dragStart = null;
                }
            }

            @Override
            public void mouseClicked(MouseEvent e) {
                if (e.getClickCount() == 2) {
                    showSymbolInBrowser();
                }
            }
        });
        pnlTicker.addMouseMotionListener(new MouseMotionAdapter() {
            @Override
            public void mouseDragged(MouseEvent e) {
                if (dragStart != null) {
                    Point current = e.getLocationOnScreen();
                    int x = current.x - dragStart.x;
                    int y = current.y - dragStart.y;
                    x = Math.max(x, 0);
                    y = Math.max(y, 0);
                    Rectangle screen = Utils.getAllScreensBounds();
                    x = x + getWidth() > screen.width ? screen.width - getWidth() : x;
                    y = y + getHeight() > screen.height ? screen.height - getHeight() : y;
                    setLocation(x, y);
                }
            }
        });

        // Set a timer to keep track of the mouse position for tooltips
        Timer timer = new Timer(500, e -> {

            // Get the current mouse position and check if it's over a symbol
            Point mousePos = MouseInfo.getPointerInfo().getLocation();
            LivePrice price = getLivePriceAtPoint(mousePos);

            // If over a symbol, select it; otherwise, clear selection
            if (price != null && price.getSymbolTransaction() != null && !price.getSymbolTransaction().isSelected()) {

                // Clear any previous selected symbols and select this one
                symbols.clearSelected();
                price.getSymbolTransaction().setSelected(true);

                // Draw the prices with the selection
                drawLivePrices();

                // Show the stock preview
                stockPanel.showSymbol(price, mousePos);
            }

            // Not over a symbol or the panel, clear any selected symbols
            else {
                Point screenLocation = pnlStocks.getLocationOnScreen();
                Rectangle bounds = new Rectangle(screenLocation.x, screenLocation.y, pnlStocks.getWidth(), pnlStocks.getHeight());
                if (!bounds.contains(mousePos)) {
                    if (symbols.clearSelected()) {
                        drawLivePrices();
                    }
                }

                // Check to see if it's over the summary panel
                screenLocation = pnlSummary.getLocationOnScreen();
                bounds = new Rectangle(screenLocation.x, screenLocation.y, pnlSummary.getWidth(), pnlSummary.getHeight());
                if (bounds.contains(mousePos)) {
                    summaryPanel.showSummary(mousePos);
                }
            }
        });
        timer.start();

        // Set up the resize cursors
        pnlLeftDrag.setCursor(Cursor.getPredefinedCursor(Cursor.E_RESIZE_CURSOR));
        pnlRightDrag.setCursor(Cursor.getPredefinedCursor(Cursor.W_RESIZE_CURSOR));

        // Add mouse listeners for resizing the ticker
        pnlLeftDrag.addMouseListener(new MouseAdapter() {
            @Override
            public void mousePressed(MouseEvent e) {
                leftDragStart = e.getPoint();
                right = getLocationOnScreen().x + getWidth();
            }

            @Override
            public void mouseReleased(MouseEvent e) {
                if (leftDragStart != null) {
                    settings.setWindowX(getX());
                    settings.setWindowWidth(getWidth());
                    leftDragStart = null;
                }
            }

            @Override
            public void mouseEntered(MouseEvent e) {
                pnlLeftDrag.setBackground(Color.LIGHT_GRAY);
            }

            @Override
            public void mouseExited(MouseEvent e) {
                pnlLeftDrag.setBackground(pnlTicker.getBackground());
            }
        });

        // Track the dragging for left resize
        pnlLeftDrag.addMouseMotionListener(new MouseMotionAdapter() {
            @Override
            public void mouseDragged(MouseEvent e) {
                if (leftDragStart != null) {
                    Point current = e.getLocationOnScreen();
                    int x = current.x - leftDragStart.x;
                    x = Math.max(x, 0);
                    int width = right - x;
                    width = Math.max(width, 150);
                    x = right - width;
                    setLocation(x, getY());
                    pnlTicker.setPreferredSize(new Dimension(width, pnlTicker.getHeight()));
                    setSize(new Dimension(width, pnlTicker.getHeight()));
                }
            }
        });

        // Add mouse listeners for resizing the ticker
        pnlRightDrag.addMouseListener(new MouseAdapter() {
            @Override
            public void mousePressed(MouseEvent e) {
                rightDragStart = e.getPoint();
                left = getLocationOnScreen().x;
                pnlRightDrag.setBackground(Color.LIGHT_GRAY);
            }

            @Override
            public void mouseReleased(MouseEvent e) {
                if (rightDragStart != null) {
                    settings.setWindowWidth(getWidth());
                    rightDragStart = null;
                }
                pnlLeftDrag.setBackground(pnlTicker.getBackground());
            }

            @Override
            public void mouseEntered(MouseEvent e) {
                pnlRightDrag.setBackground(Color.LIGHT_GRAY);
            }

            @Override
            public void mouseExited(MouseEvent e) {
                pnlRightDrag.setBackground(pnlTicker.getBackground());
            }
        });

        // Track the dragging for right resize
        pnlRightDrag.addMouseMotionListener(new MouseMotionAdapter() {
            @Override
            public void mouseDragged(MouseEvent e) {
                if (rightDragStart != null) {
                    Point current = e.getLocationOnScreen();
                    int x = current.x - rightDragStart.x + pnlRightDrag.getWidth();
                    x = Math.min(x, Toolkit.getDefaultToolkit().getScreenSize().width);
                    int width = x - left;
                    width = Math.max(width, 150);
                    pnlTicker.setPreferredSize(new Dimension(width, pnlTicker.getHeight()));
                    setSize(new Dimension(width, pnlTicker.getHeight()));
                }
            }
        });
    }

    /**
     * Opens the symbol under the mouse cursor in the default web browser.
     */
    private void showSymbolInBrowser() {
        Point mousePos = MouseInfo.getPointerInfo().getLocation();
        LivePrice price = getLivePriceAtPoint(mousePos);
        if (price != null) {
            try {
                String url = String.format("%s/%s?p=%s", SettingsManager.BROWSER_STOCK_LAUNCH_URL, price.getSymbol(), price.getSymbol());
                Desktop desktop = Desktop.getDesktop();
                desktop.browse(new URI(url));
            }
            catch (Exception ex) {
                log.error("Failed to open Browser URL", ex);
            }
        }
    }

    /**
     * Retrieves the live price at a given point.
     *
     * @param point The point on the screen to check.
     * @return The LivePrice at the point, or null if none found.
     */
    protected LivePrice getLivePriceAtPoint(Point point) {

        // Check if the point is within the stocks panel
        Point screenLocation = pnlStocks.getLocationOnScreen();
        Rectangle bounds = new Rectangle(screenLocation.x, screenLocation.y, pnlStocks.getWidth(), pnlStocks.getHeight());
        if (!bounds.contains(point)) {
            return null;
        }

        // Adjust the mouse point to be relative to the stocks panel
        Point pointOverStocks = new Point(point.x - bounds.x, point.y - bounds.y);

        // Adjust point for scrolling if necessary
        int scrollPosition = pnlStocks.getScrollPosition();
        if (pointOverStocks.x + scrollPosition > pnlStocks.getTotalTextWidth() &&
                pnlStocks.getTotalTextWidth() > pnlStocks.getWidth()) {
            scrollPosition -= pnlStocks.getTotalTextWidth();
        }

        // Find the live price at the adjusted point
        pointOverStocks.x += scrollPosition;
        for (LivePrice livePrice : livePrices) {
            bounds = livePrice.getBounds();
            if (bounds != null && bounds.contains(pointOverStocks)) {
                log.debug("Found live price at point {}: {}", pointOverStocks, livePrice.getSymbolTransaction().getCode());
                return livePrice;
            }
        }
        return null;
    }

    /**
     * Sets up the context menu for the ticker panel.
     */
    private void setupContextMenu() {
        JPopupMenu contextMenu = new JPopupMenu();
        JMenuItem symbolsItem = new JMenuItem("Edit Symbols...");
        symbolsItem.addActionListener(e -> showSymbolsDialog());
        contextMenu.add(symbolsItem);
        JMenuItem settingsItem = new JMenuItem("Edit Settings...");
        settingsItem.addActionListener(e -> showSettingsDialog());
        contextMenu.add(settingsItem);
        contextMenu.addSeparator();

        JMenuItem fontSize = new JMenu("Font Size");
        fontSizeItemSmall = new JCheckBoxMenuItem("Small", settings.getFontSize() == SettingsManager.FONT_SIZE_SMALL);
        fontSizeItemSmall.addActionListener(e -> setFontSize(SettingsManager.FONT_SIZE_SMALL));
        fontSize.add(fontSizeItemSmall);
        fontSizeItemMedium = new JCheckBoxMenuItem("Normal", settings.getFontSize() == SettingsManager.FONT_SIZE_MEDIUM);
        fontSizeItemMedium.addActionListener(e -> setFontSize(SettingsManager.FONT_SIZE_MEDIUM));
        fontSize.add(fontSizeItemMedium);
        fontSizeItemLarge = new JCheckBoxMenuItem("Large", settings.getFontSize() == SettingsManager.FONT_SIZE_LARGE);
        fontSizeItemLarge.addActionListener(e -> setFontSize(SettingsManager.FONT_SIZE_LARGE));
        fontSize.add(fontSizeItemLarge);
        contextMenu.add(fontSize);
        contextMenu.addSeparator();

        JMenuItem scroll = new JMenu("Scroll");
        scrollItemSlow = new JCheckBoxMenuItem("Slow", settings.getTickerSpeed() == SettingsManager.SCROLL_SPEED_SLOW);
        scrollItemSlow.addActionListener(e -> setTicketSpeed(SettingsManager.SCROLL_SPEED_SLOW));
        scroll.add(scrollItemSlow);
        scrollItemNormal = new JCheckBoxMenuItem("Medium", settings.getTickerSpeed() == SettingsManager.SCROLL_SPEED_MEDIUM);
        scrollItemNormal.addActionListener(e -> setTicketSpeed(SettingsManager.SCROLL_SPEED_MEDIUM));
        scroll.add(scrollItemNormal);
        scrollItemFast = new JCheckBoxMenuItem("Fast", settings.getTickerSpeed() == SettingsManager.SCROLL_SPEED_FAST);
        scrollItemFast.addActionListener(e -> setTicketSpeed(SettingsManager.SCROLL_SPEED_FAST));
        scroll.add(scrollItemFast);
        contextMenu.add(scroll);
        contextMenu.addSeparator();

        JMenuItem refresh = new JMenuItem("Refresh");
        refresh.addActionListener(e -> refreshTicker());
        contextMenu.add(refresh);
        contextMenu.addSeparator();

        onTop = new JCheckBoxMenuItem("Keep the ticker on top of other windows", settings.isAlwaysOnTop());
        onTop.addActionListener(e -> setOnTopMost(onTop.isSelected()));
        contextMenu.add(onTop);
        JCheckBoxMenuItem runAtStartup = new JCheckBoxMenuItem("Run at Startup", StartupManager.isStartupEnabled());
        runAtStartup.addActionListener(e -> StartupManager.enableStartup(runAtStartup.isSelected()));
        contextMenu.add(runAtStartup);
        contextMenu.addSeparator();

        JMenuItem help = new JMenuItem("Help...");
        help.addActionListener(e -> {
            try {
                Desktop desktop = Desktop.getDesktop();
                desktop.browse(new URI("https://github.com/steveohara/stockticker"));
            }
            catch (Exception ex) {
                log.error("Failed to open help URL", ex);
            }
        });
        contextMenu.add(help);
        JMenuItem about = new JMenuItem("About...");
        about.addActionListener(e -> Utils.showTopmostMessage(VersionInfo.getVersionString(), "About", JOptionPane.INFORMATION_MESSAGE));
        contextMenu.add(about);
        contextMenu.addSeparator();

        JMenuItem export = new JMenu("Export");
        contextMenu.add(export);
        JMenuItem exportCsvAll = new JMenuItem("Export to CSV (All)");
        export.add(exportCsvAll);
        JMenuItem exportCsvLive = new JMenuItem("Export to CSV (live)");
        export.add(exportCsvLive);
        JMenuItem exportCsvSummarised = new JMenuItem("Export to CSV (Summarised)");
        export.add(exportCsvSummarised);
        contextMenu.addSeparator();

        JMenuItem exitItem = new JMenuItem("Exit");
        exitItem.addActionListener(e -> exitApplication());
        contextMenu.add(exitItem);

        pnlTicker.setComponentPopupMenu(contextMenu);
        pnlLeftDrag.setComponentPopupMenu(contextMenu);
        pnlRightDrag.setComponentPopupMenu(contextMenu);
    }

    /**
     * Displays the symbols dialog.
     */
    private void showSymbolsDialog() {
        SwingUtilities.invokeLater(() -> {
            try {
                new SymbolsForm(this);
            }
            catch (Exception ex) {
                Utils.showTopmostMessage("Failed to open Symbols Form\n" + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        });
    }

    /**
     * Displays the settings dialog.
     */
    private void showSettingsDialog() {
        SwingUtilities.invokeLater(() -> new SettingsForm(this));
    }

    /**
     * Sets whether the ticker should always be on top and updates settings.
     *
     * @param onTopMost True to keep the ticker on top, false otherwise.
     */
    private void setOnTopMost(boolean onTopMost) {
        setAlwaysOnTop(onTopMost);
        settings.setAlwaysOnTop(onTopMost);
        onTop.setSelected(onTopMost);
    }

    /**
     * Sets the font characteristics for the ticker and updates settings.
     */
    private void setTextCharacteristics() {

        // Set the font for the ticker and cascade it to the children panels
        Font masterFont = new Font(settings.getFontName(), settings.getFontStyle(), Math.round(settings.getFontSize())).deriveFont(settings.getFontSize());
        setFont(masterFont);
        pnlDaySummary.setFont(masterFont);
        pnlSummary.setFont(masterFont);
        pnlStocks.setFont(masterFont);

        // Base the size of the ticker on the font height
        pnlTicker.setPreferredSize(new Dimension(settings.getWindowWidth(), getFontMetrics(masterFont).getHeight() + 1));
        setSize(new Dimension(getWidth(), getFontMetrics(masterFont).getHeight() + 1));

        // Update the font size menu items
        float size = settings.getFontSize();
        fontSizeItemSmall.setSelected(size <= SettingsManager.FONT_SIZE_SMALL);
        fontSizeItemMedium.setSelected(size > SettingsManager.FONT_SIZE_SMALL && size < SettingsManager.FONT_SIZE_LARGE);
        fontSizeItemLarge.setSelected(size >= SettingsManager.FONT_SIZE_LARGE);
    }

    /**
     * Sets the font size for the ticker and updates settings.
     *
     * @param size The new font size.
     */
    private void setFontSize(float size) {

        // Set the font for the ticker and cascade it to the children panels
        settings.setFontSize(size);
        setTextCharacteristics();

        // Redraw all the ticker content
        drawTickerContent();
    }

    /**
     * Refreshes the ticker data and redraws the content.
     * This is a synchronous call to the data sources so could take some time
     */
    private void refreshTicker() {
        SwingUtilities.invokeLater(() -> {
            rates.refreshExchangeRates();
            prices.refreshPrices(true);
        });
    }

    /**
     * Sets the ticker scroll speed and updates settings.
     *
     * @param speed The new scroll speed.
     */
    private void setTicketSpeed(int speed) {
        settings.setTickerSpeed(speed);
        pnlStocks.setScrollSpeed(speed);
        scrollItemSlow.setSelected(speed == SettingsManager.SCROLL_SPEED_SLOW);
        scrollItemNormal.setSelected(speed == SettingsManager.SCROLL_SPEED_MEDIUM);
        scrollItemFast.setSelected(speed == SettingsManager.SCROLL_SPEED_FAST);
    }

    /**
     * Exits the application, saving settings and stopping timers.
     */
    private void exitApplication() {
        pnlStocks.stopScrolling();
        prices.stopScheduler();
        rates.stopScheduler();
        System.exit(0);
    }

    /**
     * Creates all the UI components
     */
    private void createUIComponents() {
        pnlTicker = new JPanel();
        pnlTicker.setLayout(new BoxLayout(pnlTicker, BoxLayout.X_AXIS));
        pnlTicker.setBackground(Color.black);

        // Fixed width drag panels
        pnlLeftDrag = new JPanel();
        pnlLeftDrag.setBackground(pnlTicker.getBackground());
        pnlLeftDrag.setPreferredSize(new Dimension(7, 50));
        pnlLeftDrag.setMaximumSize(pnlLeftDrag.getPreferredSize());
        pnlLeftDrag.setMinimumSize(pnlLeftDrag.getPreferredSize());
        pnlTicker.add(pnlLeftDrag);

        // Growing panels
        pnlSummary = new ColouredTextPanel();
        pnlSummary.setBackground(pnlTicker.getBackground());
        pnlSummary.setDisplayStyle(ColouredTextPanel.DISPLAY_STYLE.FIT);
        pnlSummary.setBorder(BorderFactory.createMatteBorder(0, 0, 0, 1, settings.getLabelColor()));
        pnlTicker.add(pnlSummary);

        pnlDaySummary = new ColouredTextPanel();
        pnlDaySummary.setBackground(pnlTicker.getBackground());
        pnlDaySummary.setDisplayStyle(ColouredTextPanel.DISPLAY_STYLE.FIT);
        pnlDaySummary.setBorder(BorderFactory.createMatteBorder(0, 0, 0, 1, settings.getLabelColor()));
        pnlTicker.add(pnlDaySummary);

        pnlStocks = new ColouredTextPanel();
        pnlStocks.setBackground(pnlTicker.getBackground());
        pnlStocks.setDisplayStyle(ColouredTextPanel.DISPLAY_STYLE.SCROLL);
        pnlTicker.add(pnlStocks);

        pnlRightDrag = new JPanel();
        pnlRightDrag.setBackground(pnlTicker.getBackground());
        pnlRightDrag.setPreferredSize(pnlLeftDrag.getPreferredSize());
        pnlRightDrag.setMaximumSize(pnlLeftDrag.getPreferredSize());
        pnlRightDrag.setMinimumSize(pnlLeftDrag.getPreferredSize());
        pnlTicker.add(pnlRightDrag);

        setType(Type.UTILITY);
        setUndecorated(true);
        setContentPane(pnlTicker);
    }
}
