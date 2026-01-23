/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.model;

import com.pivotal.stockticker.service.PersistanceManager;
import lombok.AccessLevel;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.awt.*;
import java.util.prefs.Preferences;

/**
 * Application settings and configuration
 * These changes will be persisted to settings.json automatically
 */
@Slf4j
@Getter
@Setter
public class SettingsManager extends PersistanceManager {

    public static String BROWSER_STOCK_LAUNCH_URL = "https://finance.yahoo.com/quote";

    public static final float FONT_SIZE_SMALL = 11.1f;
    public static final float FONT_SIZE_MEDIUM = 12.49f;
    public static final float FONT_SIZE_LARGE = 14.3f;

    public static final int SCROLL_SPEED_SLOW = 1;
    public static final int SCROLL_SPEED_MEDIUM = 2;
    public static final int SCROLL_SPEED_FAST = 4;

    private String proxyServer = null;
    @Setter(AccessLevel.NONE)
    private int frequency = 60;
    private int exchangeRateFrequency = 3600;
    private String currencyCode = null;
    private String currencySymbol = null;
    private double totalInvestment = 0.0;
    private double margin = 0.0;
    private Color upColor = new Color(0, 255, 0);
    private Color downColor = new Color(255, 0, 0);
    private Color backgroundColor = Color.BLACK;
    private Color normalTextColor = Color.WHITE;
    private Color upArrowColor = new Color(0, 255, 0);
    private Color downArrowColor = new Color(255, 0, 0);
    private Color labelColor = Color.LIGHT_GRAY;
    private String fontName = "Calibri";
    private boolean fontBold = false;
    private boolean fontItalic = false;
    private float fontSize = FONT_SIZE_MEDIUM;
    private int tickerSpeed = SCROLL_SPEED_MEDIUM;
    private boolean showPortfolioProfitAndLoss = true;
    private boolean showPortfolioProfitAndLossPercent = true;
    private boolean showTotalCost = false;
    private boolean showTotalValue = false;
    private boolean showDailySummary = true;
    private boolean showUniqueSymbols = true;
    private boolean alwaysOnTop = true;
    private boolean hideDisabledSymbols = false;
    private String highAlarmWaveFile = null;
    private String lowAlarmWaveFile = null;
    private String alphaVantageToken = null;
    private String marketStackToken = null;
    private String twelveDataToken = null;
    private String finhubToken = null;
    private String tiingoToken = null;
    private String freeCurrencyToken = null;
    private int windowX = 100;
    private int windowY = 100;
    private int windowWidth = 800;
    private String summarySortColumn = "Code";
    private String summarySortOrder = "ascending";
    private String daySortColumn = "Code";
    private String daySortOrder = "ascending";

    /**
     * Sets the update frequency, ensuring it is within valid bounds (1 to 600 seconds).
     *
     * @param frequency Update frequency in seconds.
     */
    public void setFrequency(int frequency) {
        this.frequency = Math.max(1, Math.min(600, frequency));
    }

    /**
     * Creates an auto-saving proxy instance of this class so that we can intercept method calls
     * and automatically save the values to persistent storage.
     *
     * @return A proxy instance of this class.
     * @throws Exception if proxy creation fails.
     */
    public static SettingsManager getPersistentSettings() throws Exception {
        return createProxyInstance(SettingsManager.class, Preferences.userRoot().node(ROOT_NODE + SettingsManager.class.getSimpleName()), true);
    }

    /**
     * Returns the font style based on the bold and italic settings.
     * @return The font style as an integer constant from the Font class.
     */
    public int getFontStyle() {
        return (fontBold ? Font.BOLD : Font.PLAIN) |
                (fontItalic ? Font.ITALIC : Font.PLAIN);
    }

    /**
     * Determines if any summary information is set to be displayed.
     *
     * @return true if any summary display options are enabled, false otherwise.
     */
    public boolean isShowSummary() {
        return showPortfolioProfitAndLoss ||
               showPortfolioProfitAndLossPercent ||
               showTotalCost ||
               showTotalValue;
    }

    /**
     * Reloads the settings from persistent storage.
     */
    public void loadFromStorage() {
        loadFromStorage(Preferences.userRoot().node(PersistanceManager.ROOT_NODE + SettingsManager.class.getSimpleName()));
    }

    /**
     * Gets the Font object based on the current font settings.
     *
     * @return The Font object configured with the current font name, style, and size.
     */
    public Font getFont() {
        Font font = new Font(getFontName(), getFontStyle(), 1);
        return font.deriveFont(getFontSize());
    }
}
