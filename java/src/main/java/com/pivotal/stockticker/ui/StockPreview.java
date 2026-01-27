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
import com.pivotal.stockticker.ui.components.ColouredTextPanel;
import com.pivotal.stockticker.ui.components.SettingsLabel;
import com.pivotal.stockticker.utils.CallbackInterface;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.time.Duration;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Form showing the stock preview
 */
@Slf4j
public class StockPreview extends JDialog implements CallbackInterface {

    private static final String AGENT_NAME = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36";
    private final HttpClient client = HttpClient.newBuilder().followRedirects(HttpClient.Redirect.ALWAYS).connectTimeout(Duration.ofSeconds(10)).build();

    // Where to get a chart from
    private static final String CHART_URL = "https://www.reuters.wallst.com/enhancements/chartapi/index_chart_api.asp?symbol=%s&duration=%d&headerType=quote&width=%d&height=%d";

    private ColouredTextPanel pnlSummary;
    private SettingsLabel lblGraph;
    private final TickerBar tickerBar;
    private LivePrice livePrice;
    private GRAPH_TYPE graphType = GRAPH_TYPE.DAY;

    // Types of graph durations to display
    private enum GRAPH_TYPE {
        DAY(1),
        WEEK(7);

        int durationDays = 1;

        GRAPH_TYPE(int durationDays) {
            this.durationDays = durationDays;
        }
    }

    /**
     * Creates new form SettingsForm
     *
     * @param tickerBar Parent callback interface
     */
    public StockPreview(TickerBar tickerBar) {
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
            // the same symbol, hide the preview
            Point mousePos = MouseInfo.getPointerInfo().getLocation();
            Point screenLocation = getLocationOnScreen();
            Rectangle bounds = new Rectangle(screenLocation.x, screenLocation.y, getWidth(), getHeight());
            if (!bounds.contains(mousePos)) {
                LivePrice price = tickerBar.getLivePriceAtPoint(mousePos);
                if (price == null || !price.getSymbol().equalsIgnoreCase(livePrice.getSymbol())) {
                    log.info("Closing stock preview for {} as mouse moved to {}", livePrice, price);
                    setVisible(false);
                }
            }
        });
        timer.start();
    }

    /**
     * Sets up component listeners
     */
    private void initListeners() {

        // Add a listner to the graph label for double-clicks to change the graph type
        lblGraph.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                if (e.getClickCount() == 2) {
                    graphType = graphType == GRAPH_TYPE.DAY ? GRAPH_TYPE.WEEK : GRAPH_TYPE.DAY;
                    displayGraph();
                }
            }
        });
    }

    /**
     * Displays the graph for the current live price and graph type
     */
    private void displayGraph() {

        // Check the cache first
        lblGraph.setIcon(null);
        ImageIcon imageIcon = ImageCache.getImageFromCache(livePrice.getSymbol(), graphType);
        if (imageIcon != null) {
            lblGraph.setIcon(imageIcon);
            return;
        }

        // Run a background thread to load the image
        SwingWorker<ImageIcon, Void> worker = new SwingWorker<>() {
            @Override
            protected ImageIcon doInBackground() throws Exception {
                lblGraph.setText("Loading graph...");

                // Build request to fetch graph image
                String url = String.format(CHART_URL, livePrice.getSymbol(), graphType.durationDays, lblGraph.getWidth(), lblGraph.getHeight());
                HttpRequest request = HttpRequest.newBuilder()
                        .uri(URI.create(url))
                        .GET()
                        .header("Accept", "image/png")
                        .header("User-Agent", AGENT_NAME)
                        .build();

                // Loop round at most 3 times to get a valid image
                HttpResponse<byte[]> response = null;
                int loops = 1;
                do {
                    try {

                        // Get raw bytes of the image
                        response = client.send(request, HttpResponse.BodyHandlers.ofByteArray());
                        if (response.statusCode() != 200) {
                            log.error("Failed to fetch graph for {}: HTTP [{}]", livePrice.getSymbol(), response.statusCode());
                            Thread.sleep(500);
                        }
                        else {
                            return new ImageIcon(response.body());
                        }
                    }
                    catch (Exception e) {
                        log.error("Error fetching graph from {} - {}", url, e.getMessage());
                    }
                } while (loops++ < 3 && response != null && response.statusCode() != 200);
                log.error("Failed to get graph from {} - after {} attempts", url, loops - 1);
                lblGraph.setText("Failed to load graph from external source");
                return null;
            }

            @Override
            protected void done() {
                try {
                    ImageIcon icon = get();
                    if (icon != null) {
                        ImageCache.add(livePrice.getSymbol(), graphType, icon);
                        lblGraph.setIcon(icon);
                        lblGraph.setText("");
                    }
                }
                catch (Exception e) {
                    lblGraph.setText("Error loading image");
                    log.error("Error loading image", e);
                }
            }
        };
        worker.execute();
    }

    /**
     * Shows the stock preview for the given live price at the given location
     *
     * @param livePrice Live price to show
     * @param location  Location to show the preview at
     */
    public void showSymbol(LivePrice livePrice, Point location) {

        // If this is the same symbol as currently shown, do nothing
        if (this.livePrice != null && this.livePrice.equals(livePrice)) {
            return;
        }
        this.livePrice = livePrice;

        // Reset the gra[ph display to be a day
        graphType = GRAPH_TYPE.DAY;

        // Position the form near the ticker bar and the selected live price
        Rectangle screen = Utils.getAllScreensBounds();
        int x = (int) location.getX();
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
        displayGraph();

        pnlSummary.cls();
        pnlSummary.print(livePrice.getSymbol());
    }

    @Override
    public void changed(Object source) {
        SettingsManager settings = SettingsManager.getInstance();
        setBackground(settings.getBackgroundColor());
        pnlSummary.setBackground(settings.getBackgroundColor());
    }

    /**
     * Initializes all the UI components
     */
    private void initComponents() {

        setAlwaysOnTop(true);
        setResizable(false);
        setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);

        setSize(700, 350);
        setPreferredSize(getSize());
        setMaximumSize(getSize());
        setMinimumSize(getSize());

        setResizable(false);
        getContentPane().setLayout(null);

        pnlSummary = ColouredTextPanel.create().withDimensions(200, getHeight()).atRight(getWidth()).to(getContentPane());
        lblGraph = SettingsLabel.create("").withDimensions(getWidth() - pnlSummary.getWidth(), getHeight()).atRight(pnlSummary.getX()).atTop(0).setAlignment(SwingConstants.CENTER).to(getContentPane());
        lblGraph.setBorder(BorderFactory.createLineBorder(Color.DARK_GRAY));

        SettingsManager settings = SettingsManager.getInstance();
        setBackground(settings.getBackgroundColor());
        pnlSummary.setBackground(settings.getBackgroundColor());

        setType(Window.Type.UTILITY);
        setUndecorated(true);
    }


    /**
     * Cache manager for stock images
     */
    private static class ImageCache {

        public static final int EXPIRY_DURATION_MS = 10 * 60 * 1000; // 10 minutes

        private static final Map<String, ImageCache> imageCache = new ConcurrentHashMap<>();

        public final String symbol;
        public final GRAPH_TYPE graphType;
        public final ImageIcon imageIcon;
        public final long timestamp = System.currentTimeMillis();

        /**
         * Constructor
         *
         * @param symbol    Stock symbol
         * @param graphType Graph type
         * @param imageIcon Image icon
         */
        private ImageCache(String symbol, GRAPH_TYPE graphType, ImageIcon imageIcon) {
            this.symbol = symbol;
            this.graphType = graphType;
            this.imageIcon = imageIcon;
        }

        /**
         * Returns true if the cache entry has expired
         *
         * @return True if it hasn't expired
         */
        private boolean isCurrent() {
            return System.currentTimeMillis() - timestamp <= EXPIRY_DURATION_MS;
        }

        /**
         * Adds an image to the cache
         *
         * @param symbol    Stock symbol
         * @param graphType Graph type
         * @param imageIcon Image icon
         */
        public static void add(String symbol, GRAPH_TYPE graphType, ImageIcon imageIcon) {
            imageCache.put(symbol + "_" + graphType.toString(), new ImageCache(symbol, graphType, imageIcon));
        }

        /**
         * Retrieves an image from the cache
         *
         * @param symbol    Stock symbol
         * @param graphType Graph type
         * @return Image icon or null if not found or expired
         */
        public static ImageIcon getImageFromCache(String symbol, GRAPH_TYPE graphType) {
            ImageCache entry = imageCache.get(symbol + "_" + graphType.toString());
            if (entry != null && entry.isCurrent()) {
                return entry.imageIcon;
            }
            return null;
        }

    }

}
