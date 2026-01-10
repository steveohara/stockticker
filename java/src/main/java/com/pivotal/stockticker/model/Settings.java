package com.pivotal.stockticker.model;

import lombok.AccessLevel;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.awt.*;

/**
 * Application settings and configuration
 * These changes will be persisted to settings.json automatically
 */
@Slf4j
@Getter
@Setter
public class Settings extends PersistanceManager {

    public static final int FONT_SIZE_SMALL = 11;
    public static final int FONT_SIZE_MEDIUM = 13;
    public static final int FONT_SIZE_LARGE = 16;

    public static final int SCROLL_SPEED_SLOW = 1;
    public static final int SCROLL_SPEED_MEDIUM = 2;
    public static final int SCROLL_SPEED_FAST = 4;

    private String proxyServer = null;
    @Setter(AccessLevel.NONE)
    private int frequency = 60;
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
    private String fontName = "Calibri";
    private boolean fontBold = false;
    private boolean fontItalic = false;
    private int fontSize = FONT_SIZE_MEDIUM;
    private int tickerSpeed = SCROLL_SPEED_MEDIUM;
    private boolean showPortfolioProfitAndLoss = true;
    private boolean showPortfolioProfitAndLossPercent = true;
    private boolean showTotalCost = false;
    private boolean showTotalValue = false;
    private boolean showDailyChange = true;
    private boolean showUniqueSymbols = true;
    private boolean alwaysOnTop = true;
    private String highAlarmWaveFile = null;
    private String lowAlarmWaveFile = null;
    private String iexToken = null;
    private String alphaVantageToken = null;
    private String marketStackToken = null;
    private String twelveDataToken = null;
    private String finhubToken = null;
    private String tiingoToken = null;
    private String freeCurrencyToken = null;
    private int windowX = 100;
    private int windowY = 100;
    private int windowWidth = 800;

    /**
     * Sets the update frequency, ensuring it is within valid bounds (1 to 600 seconds).
     *
     * @param frequency Update frequency in seconds.
     */
    public void setFrequency(int frequency) {
        this.frequency = Math.max(1, Math.min(600, frequency));
    }

    /**
     * Creates a proxy instance of this class so that we can intercept method calls.
     *
     * @return A proxy instance of this class.
     * @throws Exception if proxy creation fails.
     */
    public static Settings createProxy() throws Exception {
        return createProxy(Settings.class);
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
}
