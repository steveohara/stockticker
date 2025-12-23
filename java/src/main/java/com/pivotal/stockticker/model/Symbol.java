package com.pivotal.stockticker.model;

import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.time.LocalDateTime;
import java.util.Objects;

/**
 * Represents a stock symbol with all its display and configuration properties.
 * Equivalent to cSymbol class in VB6.
 */
@Slf4j
@Getter
@Setter
public class Symbol {
    private String regKey;
    private String code;
    private String alias;
    private double price;
    private String currencyName;
    private String currencySymbol;
    private double shares;
    private boolean showPrice;
    private boolean showChange;
    private boolean showChangePercent;
    private boolean showChangeUpDown;
    private boolean showProfitLoss;
    private boolean showDayChange;
    private boolean showDayChangePercent;
    private boolean showDayChangeUpDown;
    private boolean excludeFromSummary;
    private boolean observeOnly;
    private boolean disabled;
    private double currentPrice;
    private double dayStart;
    private double dayChange;
    private double dayHigh;
    private double dayLow;
    private String errorDescription;

    // Alarm properties
    private boolean lowAlarmEnabled;
    private double lowAlarmValue;
    private boolean lowAlarmIsPercent;
    private boolean lowAlarmSoundEnabled;
    private boolean highAlarmEnabled;
    private double highAlarmValue;
    private boolean highAlarmIsPercent;
    private boolean highAlarmSoundEnabled;
    private boolean alarmShowing;

    private LocalDateTime lastUpdate;
    private String source;

    // Constructor initializing default values
    public Symbol() {
        this.regKey = String.valueOf(System.currentTimeMillis());
        this.currencySymbol = "$";
        this.currencyName = "USD";
        this.showPrice = true;
    }

    /**
     * Returns the display name, using alias if available, otherwise the code.
     *
     * @return Display name of the symbol.
     */
    public String getDisplayName() {
        return (alias == null || alias.isEmpty()) ? code : alias;
    }

    /**
     * Calculates the percentage change from the original price to the current price.
     *
     * @return Percentage change.
     */
    public double getPercentChange() {
        if (price == 0) {
            return 0;
        }
        return ((currentPrice - price) * 100) / price;
    }

    /**
     * Returns the formatted percentage change as a string with two decimal places.
     *
     * @return Formatted percentage change.
     */
    public String getFormattedPercentChange() {
        return String.format("%.2f%%", getPercentChange());
    }

    /**
     * Returns the formatted current price as a currency string.
     *
     * @return Formatted current price.
     */
    public String getFormattedValue() {
        return formatCurrencyValue(currentPrice);
    }

    /**
     * Returns the formatted total value (current price * shares) as a currency string.
     *
     * @return Formatted total value.
     */
    public String getFormattedTotalValue() {
        return formatCurrencyValue(currentPrice * shares);
    }

    /**
     * Returns the formatted cost price as a currency string.
     *
     * @return Formatted cost price.
     */
    public String getFormattedCost() {
        return formatCurrencyValue(price);
    }

    /**
     * Returns the formatted total cost (cost price * shares) as a currency string.
     *
     * @return Formatted total cost.
     */
    public String getFormattedTotalCost() {
        return formatCurrencyValue(price * shares);
    }

    /**
     * Calculates the profit or loss based on current price and original price.
     *
     * @return Profit or loss amount.
     */
    public double getProfitLoss() {
        return (currentPrice - price) * shares;
    }

    /**
     * Returns the formatted profit or loss as a currency string.
     *
     * @return Formatted profit or loss.
     */
    public String getFormattedProfitLoss() {
        return formatCurrencyValue(getProfitLoss());
    }

    /**
     * Formats a given value as a currency string with the appropriate currency symbol.
     *
     * @param value Value to format.
     * @return Formatted currency string.
     */
    private String formatCurrencyValue(double value) {
        return String.format("%s%.2f", currencySymbol, Math.abs(value));
    }

    /**
     * Generates a sort key based on the code and registration key.
     *
     * @return Sort key string.
     */
    public String getSortKey() {
        return String.format("%-10s%-20s", code, regKey);
    }

    /**
     * Sets the stock code, ensuring it is trimmed and uppercase.
     *
     * @param code Stock code.
     */
    public void setCode(String code) {
        this.code = code != null ? code.trim().toUpperCase() : null;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        Symbol symbol = (Symbol) o;
        return Objects.equals(regKey, symbol.regKey);
    }

    @Override
    public int hashCode() {
        return Objects.hash(regKey);
    }
}
