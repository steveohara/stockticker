package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.StartupManager;
import com.pivotal.stockticker.Utils;
import com.pivotal.stockticker.VersionInfo;
import com.pivotal.stockticker.model.Settings;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseMotionAdapter;
import java.net.URI;

@Slf4j
public class TickerBar extends JFrame implements CallbackInterface {
    private final Settings settings = Settings.createProxy();
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

        drawTickerContent();

    }

    /**
     * Draw the content of the ticker
     */
    private void drawTickerContent() {
        pnlSummary.cls();
        pnlSummary.setFontColor(settings.getNormalTextColor());
        pnlSummary.setCurrentX(4);
        pnlSummary.print("this is to show the");
        pnlSummary.setFontBold(true);
        pnlSummary.setFontColor(settings.getUpColor());
        pnlSummary.print(" ↕ ");
        pnlSummary.setFontBold(false);
        pnlSummary.setFontColor(settings.getUpArrowColor());
        pnlSummary.print("summary panel is");
        pnlSummary.setFontColor(settings.getNormalTextColor());
        pnlSummary.setFontBold(true);
        pnlSummary.print(" ↓ ");
        pnlSummary.setFontBold(false);
        pnlSummary.setFontColor(settings.getDownColor());
        pnlSummary.print("growing");
        pnlSummary.setFontColor(settings.getNormalTextColor());
        pnlSummary.setFontBold(true);
        pnlSummary.print(" ↑");

        pnlDaySummary.cls();
        pnlDaySummary.setFontColor(settings.getNormalTextColor());
        pnlDaySummary.setCurrentX(10);
        pnlDaySummary.print("summary panel is growing ");

        pnlStocks.cls();
        pnlStocks.setFontColor(settings.getNormalTextColor());
        pnlStocks.setCurrentX(10);
        pnlStocks.print("hello");
        pnlStocks.setFontBold(true);
        pnlStocks.print(" steve");
        pnlStocks.setFontColor(settings.getUpColor());
        pnlStocks.setFontBold(false);
        pnlStocks.setCurrentX(pnlStocks.getCurrentX() + 25);
        pnlStocks.print("hello steve again");
        pnlStocks.paintImmediately(pnlStocks.getBounds());
    }

    @Override
    public void changed(Component c) {
        initializeUI();
        setFontSize(settings.getFontSize());
        setTicketSpeed(settings.getTickerSpeed());
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

        // Add mouse listeners for dragging the ticker
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

        JMenuItem fontSize = new JMenu("Font Size");
        fontSizeItemSmall = new JCheckBoxMenuItem("Small", settings.getFontSize() == Settings.FONT_SIZE_SMALL);
        fontSizeItemSmall.addActionListener(e -> setFontSize(Settings.FONT_SIZE_SMALL));
        fontSize.add(fontSizeItemSmall);
        fontSizeItemMedium = new JCheckBoxMenuItem("Normal", settings.getFontSize() == Settings.FONT_SIZE_MEDIUM);
        fontSizeItemMedium.addActionListener(e -> setFontSize(Settings.FONT_SIZE_MEDIUM));
        fontSize.add(fontSizeItemMedium);
        fontSizeItemLarge = new JCheckBoxMenuItem("Large", settings.getFontSize() == Settings.FONT_SIZE_LARGE);
        fontSizeItemLarge.addActionListener(e -> setFontSize(Settings.FONT_SIZE_LARGE));
        fontSize.add(fontSizeItemLarge);
        contextMenu.add(fontSize);
        contextMenu.addSeparator();

        JMenuItem scroll = new JMenu("Scroll");
        scrollItemSlow = new JCheckBoxMenuItem("Slow", settings.getTickerSpeed() == Settings.SCROLL_SPEED_SLOW);
        scrollItemSlow.addActionListener(e -> setTicketSpeed(Settings.SCROLL_SPEED_SLOW));
        scroll.add(scrollItemSlow);
        scrollItemNormal = new JCheckBoxMenuItem("Medium", settings.getTickerSpeed() == Settings.SCROLL_SPEED_MEDIUM);
        scrollItemNormal.addActionListener(e -> setTicketSpeed(Settings.SCROLL_SPEED_MEDIUM));
        scroll.add(scrollItemNormal);
        scrollItemFast = new JCheckBoxMenuItem("Fast", settings.getTickerSpeed() == Settings.SCROLL_SPEED_FAST);
        scrollItemFast.addActionListener(e -> setTicketSpeed(Settings.SCROLL_SPEED_FAST));
        scroll.add(scrollItemFast);
        contextMenu.add(scroll);
        contextMenu.addSeparator();

        JMenuItem refresh = new JMenuItem("Refresh");
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
        SymbolsForm dialog = new SymbolsForm(this, settings);
        dialog.setVisible(true);
    }

    /**
     * Displays the settings dialog.
     */
    private void showSettingsDialog() {
        SettingsForm dialog = new SettingsForm(this, settings);
        dialog.setVisible(true);
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
        drawTickerContent();
        fontSizeItemSmall.setSelected(size == Settings.FONT_SIZE_SMALL);
        fontSizeItemMedium.setSelected(size == Settings.FONT_SIZE_MEDIUM);
        fontSizeItemLarge.setSelected(size == Settings.FONT_SIZE_LARGE);
    }

    /**
     * Sets the ticker scroll speed and updates settings.
     *
     * @param speed The new scroll speed.
     */
    private void setTicketSpeed(int speed) {
        settings.setTickerSpeed(speed);
        pnlStocks.setScrollSpeed(speed);
        scrollItemSlow.setSelected(speed == Settings.SCROLL_SPEED_SLOW);
        scrollItemNormal.setSelected(speed == Settings.FONT_SIZE_MEDIUM);
        scrollItemFast.setSelected(speed == Settings.FONT_SIZE_LARGE);
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
