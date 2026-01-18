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

/**
 * The main ticker bar UI class that displays stock prices and related information.
 */
@Slf4j
public class TickerBar extends JFrame implements CallbackInterface {

    public static final int STOCK_SEPARATION = 10;
    public static final int VALUE_SEPARATION = 5;
    public static final int UP_DOWN_SEPARATION = 0;

    private final SettingsManager settings = SettingsManager.getPersistentSettings();
    private final SymbolsManager symbols = new SymbolsManager();
    private final PricesManager prices = new PricesManager(settings);
    private final ExchangeRatesManager rates = new ExchangeRatesManager(settings);
    private final ArrayList<LivePrice> livePrices = new ArrayList<>();

    private JPanel pnlLeftDrag;
    private JPanel pnlRightDrag;
    private ColouredTextPanel pnlStocks;
    private ColouredTextPanel pnlSummary;
    private ColouredTextPanel pnlDaySummary;
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
     * @throws Exception if there is an error during initialization.
     */
    public TickerBar() throws Exception {
        createUIComponents();
        setupContextMenu();
        initializeUI();
        setupDragging();

        // Add all the symbols to the prices manager
        prices.replacePrices(symbols.getAllSymbolCodes(false));

        // Add all the currencies to the exchange rates manager
        rates.replaceExchangeRates(symbols.getAllCurrencyCodes(false));

        // Draw the ticker content
        drawTickerContent();
    }

    /**
     * Draw the content of the ticker
     */
    synchronized private void drawTickerContent() {
        drawLivePrices();
        drawSummary();
    }

    /**
     * Draw the content of the ticker
     */
    private void drawLivePrices() {

        // Get a fresh list of live prices to work with
        pnlStocks.cls();
        pnlStocks.setBackground(settings.getBackgroundColor());
        livePrices.clear();
        livePrices.addAll(LivePrice.getLivePrices(symbols, prices, rates, settings));

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
    }

    /**
     * Draws the summary panel on the ticker.
     */
    private void drawSummary() {
        pnlSummary.cls();
        if (settings.isShowSummary()) {
            pnlSummary.setBackground(settings.getBackgroundColor());
            pnlSummary.setFontColor(settings.getNormalTextColor());
            pnlSummary.setFont(new Font(settings.getFontName(), settings.getFontStyle(), settings.getFontSize()));

            // Create a summary stats object to calculate the summary data
            SummaryStats summaryStats = new SummaryStats(symbols, prices, rates, settings);
            double totalValue = summaryStats.calculateTotalValue();
            double totalCost = summaryStats.calculateTotalCost();
            double adjustedTotalValue = totalValue - settings.getTotalInvestment() - settings.getMargin();

            // Draw the summary data
            pnlSummary.setFontColor(settings.getLabelColor());
            pnlSummary.print("Summary:");
            pnlSummary.setFontColor(settings.getNormalTextColor());
            if (settings.isShowPortfolioProfitAndLoss()) {
                pnlSummary.print(Utils.formatCurrencyValue(totalValue, settings.getCurrencySymbol()));
                pnlSummary.setFontColor(adjustedTotalValue < totalCost ? settings.getDownColor() : adjustedTotalValue > totalCost ? settings.getUpColor() : settings.getNormalTextColor());
                pnlSummary.print(String.format(" (%s)", Utils.formatCurrencyValue(totalValue==0.0 ? 0 : (totalCost - adjustedTotalValue), settings.getCurrencySymbol())));
                pnlSummary.setCurrentX(pnlSummary.getCurrentX() + VALUE_SEPARATION);
            }
            if (settings.isShowPortfolioProfitAndLossPercent()) {
                pnlSummary.setFontColor(adjustedTotalValue < totalCost ? settings.getDownColor() : adjustedTotalValue > totalCost ? settings.getUpColor() : settings.getNormalTextColor());
                pnlSummary.print(String.format("%.2f%%", totalValue==0.0 ? 0 : (adjustedTotalValue - totalCost) / totalCost * 100));
                pnlSummary.setCurrentX(pnlSummary.getCurrentX() + VALUE_SEPARATION);
            }
            if (settings.isShowTotalCost()) {
                pnlSummary.setFontColor(settings.getNormalTextColor());
                pnlSummary.print(String.format("Cost:%s", Utils.formatCurrencyValue(totalCost, settings.getCurrencySymbol())));
                pnlSummary.setCurrentX(pnlSummary.getCurrentX() + VALUE_SEPARATION);
            }
            if (settings.isShowTotalValue()) {
                pnlSummary.setFontColor(settings.getNormalTextColor());
                pnlSummary.print(String.format("Value:%s", Utils.formatCurrencyValue(totalValue, settings.getCurrencySymbol())));
                pnlSummary.setCurrentX(pnlSummary.getCurrentX() + VALUE_SEPARATION);
            }
            pnlSummary.print(" ");
        }
    }

    /**
     * Draws a live price on the ticker.
     *
     * @param livePrice The live price data to draw.
     */
    private void drawLivePrice(LivePrice livePrice) {

        // Draw the price and other data
        boolean bShownOtherData = drawSymbolPrice(livePrice);

        // Draw the day changes
        drawSymbolDayChanges(livePrice, bShownOtherData);
    }

    /**
     * Draws the day changes for a symbol on the ticker.
     *
     * @param livePrice     The live price data for the symbol.
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
                boolean needSeparation = symbol.isShowDayChange() || symbol.isShowDayChangePercent() || symbol.isShowProfitLoss();
                pnlStocks.setCurrentX(pnlStocks.getCurrentX() + (needSeparation ? -UP_DOWN_SEPARATION : UP_DOWN_SEPARATION));
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
                boolean needSeparation = symbol.isShowPrice() || symbol.isShowChange() || symbol.isShowChangePercent() || symbol.isShowProfitLoss();
                pnlStocks.setCurrentX(pnlStocks.getCurrentX() + (needSeparation ? -UP_DOWN_SEPARATION : UP_DOWN_SEPARATION));
                pnlStocks.setFontColor(livePrice.isUp() ? settings.getUpArrowColor() : livePrice.isDown() ? settings.getDownArrowColor() : settings.getLabelColor());
                pnlStocks.print(livePrice.isUp() ? "↑" : livePrice.isDown() ? "↓" : "↕");
                pnlStocks.setFontBold((settings.getFontStyle() | Font.BOLD) > 0);
            }
        }
        return showData;
    }

    @Override
    public void changed(Component sourceForm) {

        // If there is no component, then this is a complete initialization
        switch (sourceForm) {
            case null -> {

                // Load all the changed settings from storage
                settings.loadFromStorage();
                setTicketSpeed(settings.getTickerSpeed());
                initializeUI();

                // Load all symbols, prices and exchange rates from storage
                symbols.loadFromStorage();
                prices.loadFromStorage();
                rates.loadFromStorage();

                // Reset the schedulers to pick up any changes
                prices.resetScheduler();
                rates.resetScheduler();

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
                prices.replacePrices(symbols.getAllSymbolCodes(false));
                prices.resetScheduler();

                rates.replaceExchangeRates(symbols.getAllCurrencyCodes(false));
                rates.resetScheduler();
                drawTickerContent();
            }
            default -> {
            }
        }
    }

    /**
     * Initializes the UI components and layout.
     */
    private void initializeUI() {

        // Position and size the main frame
        setAlwaysOnTop(settings.isAlwaysOnTop());
        setFont(new Font(settings.getFontName(), settings.getFontStyle(), settings.getFontSize()));
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
        pnlDaySummary.setVisible(settings.isShowDailyChange());
        pnlStocks.setScrollSpeed(settings.getTickerSpeed());
        pnlDaySummary.setBackground(pnlTicker.getBackground());
        pnlSummary.setBackground(pnlTicker.getBackground());
        pnlStocks.setBackground(pnlTicker.getBackground());
        pnlDaySummary.setForeground(pnlTicker.getForeground());
        pnlSummary.setForeground(pnlTicker.getForeground());
        pnlStocks.setForeground(pnlTicker.getForeground());

        // Finalize and display the frame
        setVisible(true);
    }

    /**
     * Sets up dragging functionality for the ticker panel.
     */
    private void setupDragging() {

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
                    x = x + getWidth() > Toolkit.getDefaultToolkit().getScreenSize().width
                            ? Toolkit.getDefaultToolkit().getScreenSize().width - getWidth() : x;
                    y = y + getHeight() > Toolkit.getDefaultToolkit().getScreenSize().height
                            ? Toolkit.getDefaultToolkit().getScreenSize().height - getHeight() : y;
                    setLocation(x, y);
                }
            }
        });

        // Set a timer to keep track of the mouse position for tooltips
        Timer timer = new Timer(500, e -> {
            Point mousePos = MouseInfo.getPointerInfo().getLocation();
            LivePrice price = getLivePriceAtPoint(mousePos);
            pnlTicker.setToolTipText(price != null ? price.getSymbolTransaction().getDisplayName() + " " + price.getSymbolTransaction().getDisplayTimestamp() : null);
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
     * @param point The point to check.
     * @return The LivePrice at the point, or null if none found.
     */
    private LivePrice getLivePriceAtPoint(Point point) {
        Point panelOnScreen = pnlStocks.getLocationOnScreen();
        point.x -= panelOnScreen.x;
        point.y -= panelOnScreen.y;
        for (LivePrice livePrice : livePrices) {
            Rectangle bounds = livePrice.getBounds();
            if (bounds != null && bounds.contains(point)) {
                log.debug("Found live price at point {}: {}", point, livePrice.getSymbolTransaction().getCode());
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
        refresh.addActionListener(e -> drawTickerContent());
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
        about.addActionListener(e -> {
            Utils.showTopmostMessage(VersionInfo.getVersionString(), "About", JOptionPane.INFORMATION_MESSAGE);
        });
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
            new SymbolsForm(settings, this, symbols);
        });
    }

    /**
     * Displays the settings dialog.
     */
    private void showSettingsDialog() {
        SwingUtilities.invokeLater(() -> {
            new SettingsForm(this, settings);
        });
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
     * Sets the font size for the ticker and updates settings.
     *
     * @param size The new font size.
     */
    private void setFontSize(int size) {
        Font currentFont = getFont();
        Font newFont = new Font(currentFont.getName(), currentFont.getStyle(), size);
        setFont(newFont);
        pnlDaySummary.setFont(newFont);
        pnlSummary.setFont(newFont);
        pnlStocks.setFont(newFont);
        pnlTicker.setPreferredSize(new Dimension(getWidth(), getFontMetrics(newFont).getHeight() + 2));
        setSize(new Dimension(getWidth(), getFontMetrics(newFont).getHeight() + 2));
        settings.setFontSize(size);
        fontSizeItemSmall.setSelected(size == SettingsManager.FONT_SIZE_SMALL);
        fontSizeItemMedium.setSelected(size == SettingsManager.FONT_SIZE_MEDIUM);
        fontSizeItemLarge.setSelected(size == SettingsManager.FONT_SIZE_LARGE);

        // Redraw all the ticker content
        drawTickerContent();
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
        pnlSummary.setBorder(BorderFactory.createMatteBorder(0, 0, 0, 1, Color.LIGHT_GRAY));
        pnlTicker.add(pnlSummary);

        pnlDaySummary = new ColouredTextPanel();
        pnlDaySummary.setBackground(pnlTicker.getBackground());
        pnlDaySummary.setDisplayStyle(ColouredTextPanel.DISPLAY_STYLE.FIT);
        pnlDaySummary.setBorder(BorderFactory.createMatteBorder(0, 0, 0, 1, Color.LIGHT_GRAY));
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

        setType(Window.Type.UTILITY);
        setUndecorated(true);
        setContentPane(pnlTicker);
    }
}
