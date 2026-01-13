package com.pivotal.stockticker.model;

import com.pivotal.stockticker.Utils;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.time.Instant;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.prefs.Preferences;

/**
 * Represents a stock symbol with all its display and configuration properties.
 * Equivalent to cSymbol class in VB6.
 */
@Slf4j
@Getter
@Setter
public class SymbolTransaction extends PersistanceManager {
    private String key = String.valueOf(System.currentTimeMillis());
    private String code = "YAHOO";
    private String alias;
    private boolean disabled;
    private double pricePaid;
    private String currencyCode = "USD";
    private String currencySymbol = "$";
    private double sharesBought;
    private boolean showPrice = true;
    private boolean showChange;
    private boolean showChangePercent;
    private boolean showChangeUpDown;
    private boolean showProfitLoss;
    private boolean showDayChange;
    private boolean showDayChangePercent;
    private boolean showDayChangeUpDown;
    private boolean excludeFromSummary;

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

    // Transient properties (not persisted)
    private transient boolean edited = false;
    private transient boolean added = false;

    /**
     * Default constructor initializing default values.
     */
    public SymbolTransaction() {
        super();
    }

    /**
     * Creates a new proxy instance of this class loaded from persistent storage.
     *
     * @return A proxy instance of this class.
     * @throws Exception if proxy creation fails.
     */
    public static SymbolTransaction getSymbolTransaction() throws Exception {
        String key = String.valueOf(System.currentTimeMillis());
        return createProxyInstance(SymbolTransaction.class, Preferences.userRoot().node(ROOT_NODE + SymbolTransaction.class.getSimpleName() + '/' + key), false);
    }

    /**
     * Creates a proxy instance of this class loaded from persistent storage.
     *
     * @param key Unique key for the symbol transaction.
     * @return A proxy instance of this class.
     * @throws Exception if proxy creation fails.
     */
    public static SymbolTransaction getSymbolTransaction(String key) throws Exception {
        SymbolTransaction symbol = createProxyInstance(SymbolTransaction.class, Preferences.userRoot().node(ROOT_NODE + SymbolTransaction.class.getSimpleName() + '/' + key), false);
        symbol.setKey(key);
        return symbol;
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
     * Returns the formatted cost price as a currency string.
     *
     * @return Formatted cost price.
     */
    public String getFormattedCost() {
        return Utils.formatCurrencyValue(pricePaid, currencyCode);
    }

    /**
     * Returns the formatted total cost (cost price * shares) as a currency string.
     *
     * @return Formatted total cost.
     */
    public String getFormattedTotalCost() {
        return Utils.formatCurrencyValue(pricePaid * sharesBought, currencyCode);
    }

    /**
     * Generates a sort key based on the code and registration key.
     *
     * @return Sort key string.
     */
    public String getSortKey() {
        return String.format("%-20s%-20s", code, key);
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
    public String toString() {
        return code + '(' + key + ')';
    }

    /**
     * Returns a human-readable timestamp derived from the key.
     *
     * @return Formatted timestamp string.
     */
    public String getDisplayTimestamp() {
        Instant instant = Instant.ofEpochMilli(Long.parseLong(key));
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm");
        return formatter.format(LocalDateTime.ofInstant(instant, java.time.ZoneId.systemDefault()));
    }
}
