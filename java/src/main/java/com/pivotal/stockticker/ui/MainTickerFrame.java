package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.model.Settings;
import com.pivotal.stockticker.model.Symbol;
import com.pivotal.stockticker.service.PreferencesService;
import com.pivotal.stockticker.service.StockDataService;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseMotionAdapter;
import java.util.List;

@Slf4j
/**
 * Main application frame for the stock ticker.
 */
public class MainTickerFrame extends JFrame {
    private final PreferencesService prefsService;
    private final StockDataService stockService;
    private Settings settings;
    private List<Symbol> symbols;
    private JPanel tickerPanel;
    private Timer scrollTimer;
    private Timer updateTimer;
    private int scrollPosition = 0;
    private String tickerText = "";
    private static final int SCROLL_SPEED = 2;
    private static final int SCROLL_DELAY = 30;

    /**
     * Constructs the main ticker frame, initializes UI components, and starts timers.
     */
    public MainTickerFrame() {
        this.prefsService = new PreferencesService();
        this.stockService = new StockDataService();
        this.settings = prefsService.loadSettings();
        this.symbols = prefsService.loadSymbols();
        initializeUI();
        startTimers();
    }

    /**
     * Initializes the user interface components.
     */
    private void initializeUI() {
        setTitle("Stock Ticker");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setUndecorated(true);
        setBounds(settings.getWindowX(), settings.getWindowY(), settings.getWindowWidth(), 25);
        setAlwaysOnTop(settings.isAlwaysOnTop());
        tickerPanel = new JPanel() {
            @Override
            protected void paintComponent(Graphics g) {
                super.paintComponent(g);
                drawTicker(g);
            }
        };
        tickerPanel.setBackground(Color.BLACK);
        tickerPanel.setPreferredSize(new Dimension(settings.getWindowWidth(), 25));
        setupContextMenu();
        setupDragging();
        add(tickerPanel);
        setVisible(true);
    }

    /**
     * Sets up the context menu for the ticker panel.
     */
    private void setupContextMenu() {
        JPopupMenu contextMenu = new JPopupMenu();
        JMenuItem symbolsItem = new JMenuItem("Edit Symbols...");
        symbolsItem.addActionListener(e -> showSymbolsDialog());
        contextMenu.add(symbolsItem);
        JMenuItem settingsItem = new JMenuItem("Settings...");
        settingsItem.addActionListener(e -> showSettingsDialog());
        contextMenu.add(settingsItem);
        contextMenu.addSeparator();
        JMenuItem exitItem = new JMenuItem("Exit");
        exitItem.addActionListener(e -> exitApplication());
        contextMenu.add(exitItem);
        tickerPanel.setComponentPopupMenu(contextMenu);
        tickerPanel.addMouseListener(new MouseAdapter() {
            @Override
            public void mousePressed(MouseEvent e) {
                if (SwingUtilities.isLeftMouseButton(e)) {
                    contextMenu.show(tickerPanel, e.getX(), e.getY());
                }
            }
        });
    }

    /**
     * Sets up dragging functionality for the ticker panel.
     */
    private void setupDragging() {
        final Point[] dragStart = {null};
        tickerPanel.addMouseListener(new MouseAdapter() {
            @Override
            public void mousePressed(MouseEvent e) {
                if (SwingUtilities.isRightMouseButton(e)) {
                    dragStart[0] = e.getPoint();
                }
            }

            @Override
            public void mouseReleased(MouseEvent e) {
                if (dragStart[0] != null) {
                    settings.setWindowX(getX());
                    settings.setWindowY(getY());
                    prefsService.saveSettings(settings);
                    dragStart[0] = null;
                }
            }
        });
        tickerPanel.addMouseMotionListener(new MouseMotionAdapter() {
            @Override
            public void mouseDragged(MouseEvent e) {
                if (dragStart[0] != null) {
                    Point current = e.getLocationOnScreen();
                    setLocation(current.x - dragStart[0].x, current.y - dragStart[0].y);
                }
            }
        });
    }

    /**
     * Starts the scrolling and update timers.
     */
    private void startTimers() {
        scrollTimer = new Timer(SCROLL_DELAY, e -> {
            scrollPosition -= SCROLL_SPEED;
            tickerPanel.repaint();
        });
        scrollTimer.start();
        updateTimer = new Timer(settings.getFrequency() * 1000, e -> updateStockData());
        updateTimer.start();
        updateStockData();
    }

    /**
     * Updates stock data for all symbols and rebuilds the ticker text.
     */
    private void updateStockData() {
        SwingUtilities.invokeLater(() -> {
            for (Symbol symbol : symbols) {
                if (!symbol.isDisabled()) {
                    stockService.updateSymbolPrice(symbol);
                }
            }
            buildTickerText();
        });
    }

    /**
     * Builds the ticker text based on the current symbols and settings.
     */
    private void buildTickerText() {
        StringBuilder sb = new StringBuilder();
        double totalValue = 0, totalCost = 0;
        for (Symbol symbol : symbols) {
            if (symbol.isDisabled()) {
                continue;
            }
            sb.append(symbol.getDisplayName()).append(": ");
            if (symbol.isShowPrice()) {
                sb.append(symbol.getFormattedValue()).append(" ");
            }
            if (symbol.isShowChange()) {
                sb.append("(").append(String.format("%+.2f", symbol.getCurrentPrice() - symbol.getPrice())).append(") ");
            }
            if (symbol.isShowChangePercent()) {
                sb.append(symbol.getFormattedPercentChange()).append(" ");
            }
            if (symbol.isShowChangeUpDown()) {
                if (symbol.getCurrentPrice() > symbol.getPrice()) {
                    sb.append("▲ ");
                }
                else if (symbol.getCurrentPrice() < symbol.getPrice()) {
                    sb.append("▼ ");
                }
            }
            if (symbol.isShowProfitLoss()) {
                sb.append("P/L: ").append(symbol.getFormattedProfitLoss()).append(" ");
            }
            if (symbol.isShowDayChange()) {
                sb.append("Day: ").append(String.format("%+.2f", symbol.getDayChange())).append(" ");
            }
            if (symbol.isShowDayChangePercent() && symbol.getDayStart() > 0) {
                double dayChangePercent = (symbol.getDayChange() / symbol.getDayStart()) * 100;
                sb.append(String.format("(%+.2f%%)", dayChangePercent)).append(" ");
            }
            sb.append("    ");
            if (!symbol.isExcludeFromSummary() && !symbol.isObserveOnly()) {
                totalValue += symbol.getCurrentPrice() * symbol.getShares();
                totalCost += symbol.getPrice() * symbol.getShares();
            }
        }
        if (settings.isShowTotal() || settings.isShowTotalPercent() || settings.isShowTotalCost() || settings.isShowTotalValue()) {
            sb.append("    ||  SUMMARY: ");
            if (settings.isShowTotalCost()) {
                sb.append("Cost: ").append(settings.getSummaryCurrencySymbol()).append(String.format("%.2f", totalCost)).append(" ");
            }
            if (settings.isShowTotalValue()) {
                sb.append("Value: ").append(settings.getSummaryCurrencySymbol()).append(String.format("%.2f", totalValue)).append(" ");
            }
            if (settings.isShowTotal()) {
                sb.append("P/L: ").append(settings.getSummaryCurrencySymbol()).append(String.format("%+.2f", totalValue - totalCost)).append(" ");
            }
            if (settings.isShowTotalPercent() && totalCost > 0) {
                sb.append(String.format("(%+.2f%%)", ((totalValue - totalCost) / totalCost) * 100)).append(" ");
            }
        }
        tickerText = sb.toString();
    }

    /**
     * Draws the ticker text on the panel.
     *
     * @param g Graphics context.
     */
    private void drawTicker(Graphics g) {
        Graphics2D g2d = (Graphics2D) g;
        g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g2d.setFont(settings.getTickerFont());
        FontMetrics fm = g2d.getFontMetrics();
        int textWidth = fm.stringWidth(tickerText);
        if (scrollPosition < -textWidth) {
            scrollPosition = getWidth();
        }
        g2d.setColor(settings.getNormalTextColor());
        int y = (getHeight() + fm.getAscent()) / 2;
        g2d.drawString(tickerText, scrollPosition, y);
    }

    /**
     * Shows the symbols editing dialog.
     */
    private void showSymbolsDialog() {
        SymbolsDialogX dialog = new SymbolsDialogX(this, symbols, prefsService);
        dialog.setVisible(true);
        symbols = dialog.getSymbols();
        updateStockData();
    }

    /**
     * Shows the settings dialog.
     */
    private void showSettingsDialog() {
        SettingsDialogX dialog = new SettingsDialogX(this, settings, prefsService);
        dialog.setVisible(true);
        settings = dialog.getSettings();
        setAlwaysOnTop(settings.isAlwaysOnTop());
        updateTimer.setDelay(settings.getFrequency() * 1000);
        tickerPanel.repaint();
    }

    /**
     * Exits the application, saving settings and stopping timers.
     */
    private void exitApplication() {
        settings.setWindowX(getX());
        settings.setWindowY(getY());
        prefsService.saveSettings(settings);
        if (scrollTimer != null) {
            scrollTimer.stop();
        }
        if (updateTimer != null) {
            updateTimer.stop();
        }
        System.exit(0);
    }

    /**
     * Main method to launch the application.
     *
     * @param args Command-line arguments.
     */
    public static void main(String[] args) {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        }
        catch (Exception e) {
            log.error("Cannot set look and feel", e);
        }
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                new MainTickerFrame();
            }
        });
    }
}
