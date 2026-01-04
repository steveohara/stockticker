package com.pivotal.stockticker.ui;

import com.intellij.uiDesigner.core.GridConstraints;
import com.intellij.uiDesigner.core.GridLayoutManager;
import com.pivotal.stockticker.model.Settings;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseMotionAdapter;

@Slf4j
public class TickerBar extends JFrame {
    private final Settings settings = Settings.createProxy();
    private JPanel pnlLeftDrag;
    private JPanel pnlRightDrag;
    private ColouredTextPanel pnlStocks;
    private ColouredTextPanel pnlSummary;
    private ColouredTextPanel pnlDaySummary;
    private JPanel pnlTicker;

    private Point dragStart = null;
    private Point leftDragStart = null;
    private Point rightDragStart = null;
    private int left = 0;
    private int right = 0;

    public TickerBar() throws Exception {
        createUIComponents();
        initializeUI();
        setupDragging();
        setupContextMenu();

        drawTickerContent();

    }

    /**
     * Draw the content of the ticker
     */
    private void drawTickerContent() {
        pnlStocks.cls();
        pnlStocks.setFontColor(Color.RED);
        pnlStocks.setFontBold(true);
        pnlStocks.print("hello steve");
        pnlStocks.setFontColor(Color.YELLOW);
        pnlStocks.setFontBold(false);
        pnlStocks.print("hello steve again");
        pnlStocks.paintImmediately(pnlStocks.getBounds());
    }

    /**
     * Initializes the UI components and layout.
     */
    private void initializeUI() {

        // Position and size the main frame
        setUndecorated(true);
        setAlwaysOnTop(settings.isAlwaysOnTop());
        setContentPane(pnlTicker);
        setFont(settings.getTickerFont());
        setLocation(settings.getWindowX(), settings.getWindowY());
        pnlTicker.setSize(new Dimension(settings.getWindowWidth(), getFontMetrics(getFont()).getHeight() + 2));
        pnlTicker.setPreferredSize(new Dimension(settings.getWindowWidth(), getFontMetrics(getFont()).getHeight() + 2));
        setBackground(pnlTicker.getBackground());
        pack();

        // Apply font settings to all panels
        pnlDaySummary.setFont(getFont());
        pnlSummary.setFont(getFont());
        pnlStocks.setFont(getFont());

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
        });

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
            }

            @Override
            public void mouseReleased(MouseEvent e) {
                if (rightDragStart != null) {
                    settings.setWindowWidth(getWidth());
                    rightDragStart = null;
                }
            }
        });

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
//        symbolsItem.addActionListener(e -> showSymbolsDialog());
        contextMenu.add(symbolsItem);
        JMenuItem settingsItem = new JMenuItem("Settings...");
//        settingsItem.addActionListener(e -> showSettingsDialog());
        contextMenu.add(settingsItem);
        contextMenu.addSeparator();

        JMenuItem fontSize = new JMenu("Font Size");
        JMenuItem fontSizeItemSmall = new JMenuItem("Small");
        fontSizeItemSmall.addActionListener(e -> setFontSize(8));
        fontSize.add(fontSizeItemSmall);
        JMenuItem fontSizeItemMedium = new JMenuItem("Normal");
        fontSizeItemMedium.addActionListener(e -> setFontSize(11));
        fontSize.add(fontSizeItemMedium);
        JMenuItem fontSizeItemLarge = new JMenuItem("Large");
        fontSizeItemLarge.addActionListener(e -> setFontSize(14));
        fontSize.add(fontSizeItemLarge);
        contextMenu.add(fontSize);
        contextMenu.addSeparator();

        JMenuItem scroll = new JMenu("Scroll");
        JMenuItem scrollItemSlow = new JMenuItem("Slow");
        scroll.add(scrollItemSlow);
        JMenuItem scrollItemNormal = new JMenuItem("Normal");
        scroll.add(scrollItemNormal);
        JMenuItem scrollItemFast = new JMenuItem("Fast");
        scroll.add(scrollItemFast);
        contextMenu.add(scroll);
        contextMenu.addSeparator();

        JMenuItem refresh = new JMenuItem("Refresh");
        contextMenu.add(refresh);
        contextMenu.addSeparator();

        JMenuItem onTop = new JMenuItem("Keep the ticket on top of other windows");
        contextMenu.add(onTop);
        JMenuItem runAtStartup = new JMenuItem("Run at Startup");
        contextMenu.add(runAtStartup);
        contextMenu.addSeparator();

        JMenuItem help = new JMenuItem("Help");
        contextMenu.add(help);
        JMenuItem about = new JMenuItem("About...");
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
        settings.setTickerFont(newFont);
        drawTickerContent();
    }

    /**
     * Exits the application, saving settings and stopping timers.
     */
    private void exitApplication() {
//        if (scrollTimer != null) {
//            scrollTimer.stop();
//        }
//        if (updateTimer != null) {
//            updateTimer.stop();
//        }
        System.exit(0);
    }

    /**
     * Creates all the UI components
     */
    private void createUIComponents() {
        pnlTicker = new JPanel();
        pnlTicker.setLayout(new GridLayoutManager(1, 5, new Insets(0, 0, 0, 0), 0, 0));
        pnlTicker.setAlignmentX(0.0f);
        pnlTicker.setAlignmentY(0.0f);
        pnlTicker.setBackground(new Color(-16777216));
        pnlTicker.setMaximumSize(new Dimension(2147483647, 24));
        pnlTicker.setMinimumSize(new Dimension(100, 24));
        pnlTicker.setPreferredSize(new Dimension(-1, 24));
        pnlLeftDrag = new JPanel();
        pnlLeftDrag.setBackground(new Color(-13276875));
        pnlLeftDrag.setEnabled(true);
        pnlLeftDrag.setToolTipText("Drag to change the width of the ticker");
        pnlTicker.add(pnlLeftDrag, new GridConstraints(0, 0, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_BOTH, GridConstraints.SIZEPOLICY_FIXED, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_WANT_GROW, new Dimension(5, -1), new Dimension(5, -1), new Dimension(5, -1), 0, false));
        pnlRightDrag = new JPanel();
        pnlRightDrag.setBackground(new Color(-13276875));
        pnlRightDrag.setToolTipText("Drag to change the width of the ticker");
        pnlTicker.add(pnlRightDrag, new GridConstraints(0, 4, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_BOTH, GridConstraints.SIZEPOLICY_FIXED, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_WANT_GROW, new Dimension(5, -1), new Dimension(5, -1), new Dimension(5, -1), 0, false));
        pnlStocks = new ColouredTextPanel();
        pnlStocks.setBackground(new Color(-12500730));
        pnlTicker.add(pnlStocks, new GridConstraints(0, 3, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_BOTH, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_WANT_GROW, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, null, null, null, 0, false));
        pnlSummary = new ColouredTextPanel();
        pnlSummary.setBackground(new Color(-12514785));
        pnlTicker.add(pnlSummary, new GridConstraints(0, 1, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_BOTH, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_WANT_GROW, new Dimension(50, -1), new Dimension(50, -1), new Dimension(10000, -1), 0, false));
        pnlDaySummary = new ColouredTextPanel();
        pnlDaySummary.setBackground(new Color(-15659455));
        pnlTicker.add(pnlDaySummary, new GridConstraints(0, 2, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_BOTH, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_WANT_GROW, new Dimension(50, -1), new Dimension(50, -1), new Dimension(1000, -1), 0, false));
    }

}
