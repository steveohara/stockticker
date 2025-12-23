package com.pivotal.stockticker.model;

import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.awt.Color;
import java.awt.Font;

/**
 * Application settings and configuration.
 */
@Slf4j
@Getter
@Setter
public class Settings {
    private String proxy = "";
    private int frequency = 60;
    private String summaryCurrency = "USD";
    private String summaryCurrencySymbol = "$";
    private double summaryTotal = 0.0;
    private double summaryMargin = 0.0;
    private Color upColor = new Color(0, 255, 0);
    private Color downColor = new Color(255, 0, 0);
    private Color normalTextColor = Color.WHITE;
    private Color upArrowColor = new Color(0, 255, 0);
    private Color downArrowColor = new Color(255, 0, 0);
    private Font tickerFont = new Font("Arial", Font.PLAIN, 14);
    private boolean showTotal = false;
    private boolean showTotalPercent = false;
    private boolean showTotalCost = false;
    private boolean showTotalValue = false;
    private boolean showDailyChange = false;
    private boolean showPrice = false;
    private boolean showCostBase = false;
    private boolean alwaysOnTop = true;
    private String highAlarmWaveFile = "";
    private String lowAlarmWaveFile = "";
    private String alphaVantageToken = "";
    private String marketStackToken = "";
    private String twelveDataToken = "";
    private String finhubToken = "";
    private String tiingoToken = "";
    private String freeCurrencyToken = "";
    private int windowX = 100;
    private int windowY = 100;
    private int windowWidth = 800;
    private int windowHeight = 60;

    /**
     * Sets the update frequency, ensuring it is within valid bounds (1 to 600 seconds).
     *
     * @param frequency Update frequency in seconds.
     */
    public void setFrequency(int frequency) {
        this.frequency = Math.max(1, Math.min(600, frequency));
    }
}
