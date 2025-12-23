package com.pivotal.stockticker.service;

import com.pivotal.stockticker.model.Settings;
import com.pivotal.stockticker.model.Symbol;
import lombok.extern.slf4j.Slf4j;

import java.awt.Color;
import java.awt.Font;
import java.util.ArrayList;
import java.util.List;
import java.util.prefs.BackingStoreException;
import java.util.prefs.Preferences;

/**
 * Service for persisting application settings and symbols using Java Preferences API.
 */
@Slf4j
public class PreferencesService {
    private static final String SETTINGS_NODE = "StockTicker/Settings";
    private static final String SYMBOLS_NODE = "StockTicker/Symbols";
    private final Preferences settingsPrefs;
    private final Preferences symbolsPrefs;

    public PreferencesService() {
        this.settingsPrefs = Preferences.userRoot().node(SETTINGS_NODE);
        this.symbolsPrefs = Preferences.userRoot().node(SYMBOLS_NODE);
    }

    public void saveSettings(Settings settings) {
        settingsPrefs.put("proxy", settings.getProxy());
        settingsPrefs.putInt("frequency", settings.getFrequency());
        settingsPrefs.put("summaryCurrency", settings.getSummaryCurrency());
        settingsPrefs.put("summaryCurrencySymbol", settings.getSummaryCurrencySymbol());
        settingsPrefs.putDouble("summaryTotal", settings.getSummaryTotal());
        settingsPrefs.putDouble("summaryMargin", settings.getSummaryMargin());
        settingsPrefs.putInt("upColor", settings.getUpColor().getRGB());
        settingsPrefs.putInt("downColor", settings.getDownColor().getRGB());
        settingsPrefs.putInt("normalTextColor", settings.getNormalTextColor().getRGB());
        settingsPrefs.putInt("upArrowColor", settings.getUpArrowColor().getRGB());
        settingsPrefs.putInt("downArrowColor", settings.getDownArrowColor().getRGB());
        Font font = settings.getTickerFont();
        settingsPrefs.put("fontName", font.getName());
        settingsPrefs.putInt("fontStyle", font.getStyle());
        settingsPrefs.putInt("fontSize", font.getSize());
        settingsPrefs.putBoolean("showTotal", settings.isShowTotal());
        settingsPrefs.putBoolean("showTotalPercent", settings.isShowTotalPercent());
        settingsPrefs.putBoolean("showTotalCost", settings.isShowTotalCost());
        settingsPrefs.putBoolean("showTotalValue", settings.isShowTotalValue());
        settingsPrefs.putBoolean("showDailyChange", settings.isShowDailyChange());
        settingsPrefs.putBoolean("showPrice", settings.isShowPrice());
        settingsPrefs.putBoolean("showCostBase", settings.isShowCostBase());
        settingsPrefs.putBoolean("alwaysOnTop", settings.isAlwaysOnTop());
        settingsPrefs.put("highAlarmWaveFile", settings.getHighAlarmWaveFile());
        settingsPrefs.put("lowAlarmWaveFile", settings.getLowAlarmWaveFile());
        settingsPrefs.put("alphaVantageToken", settings.getAlphaVantageToken());
        settingsPrefs.put("marketStackToken", settings.getMarketStackToken());
        settingsPrefs.put("twelveDataToken", settings.getTwelveDataToken());
        settingsPrefs.put("finhubToken", settings.getFinhubToken());
        settingsPrefs.put("tiingoToken", settings.getTiingoToken());
        settingsPrefs.put("freeCurrencyToken", settings.getFreeCurrencyToken());
        settingsPrefs.putInt("windowX", settings.getWindowX());
        settingsPrefs.putInt("windowY", settings.getWindowY());
        settingsPrefs.putInt("windowWidth", settings.getWindowWidth());
        settingsPrefs.putInt("windowHeight", settings.getWindowHeight());
        try {
            settingsPrefs.flush();
        }
        catch (BackingStoreException e) {
            e.printStackTrace();
        }
    }

    public Settings loadSettings() {
        Settings settings = new Settings();
        settings.setProxy(settingsPrefs.get("proxy", ""));
        settings.setFrequency(settingsPrefs.getInt("frequency", 60));
        settings.setSummaryCurrency(settingsPrefs.get("summaryCurrency", "USD"));
        settings.setSummaryCurrencySymbol(settingsPrefs.get("summaryCurrencySymbol", "$"));
        settings.setSummaryTotal(settingsPrefs.getDouble("summaryTotal", 0.0));
        settings.setSummaryMargin(settingsPrefs.getDouble("summaryMargin", 0.0));
        settings.setUpColor(new Color(settingsPrefs.getInt("upColor", new Color(0, 255, 0).getRGB())));
        settings.setDownColor(new Color(settingsPrefs.getInt("downColor", new Color(255, 0, 0).getRGB())));
        settings.setNormalTextColor(new Color(settingsPrefs.getInt("normalTextColor", Color.WHITE.getRGB())));
        settings.setUpArrowColor(new Color(settingsPrefs.getInt("upArrowColor", new Color(0, 255, 0).getRGB())));
        settings.setDownArrowColor(new Color(settingsPrefs.getInt("downArrowColor", new Color(255, 0, 0).getRGB())));
        String fontName = settingsPrefs.get("fontName", "Arial");
        int fontStyle = settingsPrefs.getInt("fontStyle", Font.PLAIN);
        int fontSize = settingsPrefs.getInt("fontSize", 14);
        settings.setTickerFont(new Font(fontName, fontStyle, fontSize));
        settings.setShowTotal(settingsPrefs.getBoolean("showTotal", false));
        settings.setShowTotalPercent(settingsPrefs.getBoolean("showTotalPercent", false));
        settings.setShowTotalCost(settingsPrefs.getBoolean("showTotalCost", false));
        settings.setShowTotalValue(settingsPrefs.getBoolean("showTotalValue", false));
        settings.setShowDailyChange(settingsPrefs.getBoolean("showDailyChange", false));
        settings.setShowPrice(settingsPrefs.getBoolean("showPrice", false));
        settings.setShowCostBase(settingsPrefs.getBoolean("showCostBase", false));
        settings.setAlwaysOnTop(settingsPrefs.getBoolean("alwaysOnTop", true));
        settings.setHighAlarmWaveFile(settingsPrefs.get("highAlarmWaveFile", ""));
        settings.setLowAlarmWaveFile(settingsPrefs.get("lowAlarmWaveFile", ""));
        settings.setAlphaVantageToken(settingsPrefs.get("alphaVantageToken", ""));
        settings.setMarketStackToken(settingsPrefs.get("marketStackToken", ""));
        settings.setTwelveDataToken(settingsPrefs.get("twelveDataToken", ""));
        settings.setFinhubToken(settingsPrefs.get("finhubToken", ""));
        settings.setTiingoToken(settingsPrefs.get("tiingoToken", ""));
        settings.setFreeCurrencyToken(settingsPrefs.get("freeCurrencyToken", ""));
        settings.setWindowX(settingsPrefs.getInt("windowX", 100));
        settings.setWindowY(settingsPrefs.getInt("windowY", 100));
        settings.setWindowWidth(settingsPrefs.getInt("windowWidth", 800));
        settings.setWindowHeight(settingsPrefs.getInt("windowHeight", 60));
        return settings;
    }

    public void saveSymbol(Symbol symbol) {
        String key = symbol.getRegKey();
        Preferences symbolNode = symbolsPrefs.node(key);
        symbolNode.put("code", symbol.getCode());
        symbolNode.put("alias", symbol.getAlias() != null ? symbol.getAlias() : "");
        symbolNode.putDouble("price", symbol.getPrice());
        symbolNode.put("currencyName", symbol.getCurrencyName());
        symbolNode.put("currencySymbol", symbol.getCurrencySymbol());
        symbolNode.putDouble("shares", symbol.getShares());
        symbolNode.putBoolean("showPrice", symbol.isShowPrice());
        symbolNode.putBoolean("showChange", symbol.isShowChange());
        symbolNode.putBoolean("showChangePercent", symbol.isShowChangePercent());
        symbolNode.putBoolean("showChangeUpDown", symbol.isShowChangeUpDown());
        symbolNode.putBoolean("showProfitLoss", symbol.isShowProfitLoss());
        symbolNode.putBoolean("showDayChange", symbol.isShowDayChange());
        symbolNode.putBoolean("showDayChangePercent", symbol.isShowDayChangePercent());
        symbolNode.putBoolean("showDayChangeUpDown", symbol.isShowDayChangeUpDown());
        symbolNode.putBoolean("excludeFromSummary", symbol.isExcludeFromSummary());
        symbolNode.putBoolean("observeOnly", symbol.isObserveOnly());
        symbolNode.putBoolean("disabled", symbol.isDisabled());
        symbolNode.putBoolean("lowAlarmEnabled", symbol.isLowAlarmEnabled());
        symbolNode.putDouble("lowAlarmValue", symbol.getLowAlarmValue());
        symbolNode.putBoolean("lowAlarmIsPercent", symbol.isLowAlarmIsPercent());
        symbolNode.putBoolean("lowAlarmSoundEnabled", symbol.isLowAlarmSoundEnabled());
        symbolNode.putBoolean("highAlarmEnabled", symbol.isHighAlarmEnabled());
        symbolNode.putDouble("highAlarmValue", symbol.getHighAlarmValue());
        symbolNode.putBoolean("highAlarmIsPercent", symbol.isHighAlarmIsPercent());
        symbolNode.putBoolean("highAlarmSoundEnabled", symbol.isHighAlarmSoundEnabled());
        try {
            symbolNode.flush();
        }
        catch (BackingStoreException e) {
            e.printStackTrace();
        }
    }

    public void deleteSymbol(String regKey) {
        try {
            symbolsPrefs.node(regKey).removeNode();
            symbolsPrefs.flush();
        }
        catch (BackingStoreException e) {
            e.printStackTrace();
        }
    }

    public List<Symbol> loadSymbols() {
        List<Symbol> symbols = new ArrayList<>();
        try {
            String[] keys = symbolsPrefs.childrenNames();
            for (String key : keys) {
                Preferences symbolNode = symbolsPrefs.node(key);
                Symbol symbol = new Symbol();
                symbol.setRegKey(key);
                symbol.setCode(symbolNode.get("code", ""));
                symbol.setAlias(symbolNode.get("alias", ""));
                symbol.setPrice(symbolNode.getDouble("price", 0.0));
                symbol.setCurrencyName(symbolNode.get("currencyName", "USD"));
                symbol.setCurrencySymbol(symbolNode.get("currencySymbol", "$"));
                symbol.setShares(symbolNode.getDouble("shares", 0.0));
                symbol.setShowPrice(symbolNode.getBoolean("showPrice", true));
                symbol.setShowChange(symbolNode.getBoolean("showChange", false));
                symbol.setShowChangePercent(symbolNode.getBoolean("showChangePercent", false));
                symbol.setShowChangeUpDown(symbolNode.getBoolean("showChangeUpDown", false));
                symbol.setShowProfitLoss(symbolNode.getBoolean("showProfitLoss", false));
                symbol.setShowDayChange(symbolNode.getBoolean("showDayChange", false));
                symbol.setShowDayChangePercent(symbolNode.getBoolean("showDayChangePercent", false));
                symbol.setShowDayChangeUpDown(symbolNode.getBoolean("showDayChangeUpDown", false));
                symbol.setExcludeFromSummary(symbolNode.getBoolean("excludeFromSummary", false));
                symbol.setObserveOnly(symbolNode.getBoolean("observeOnly", false));
                symbol.setDisabled(symbolNode.getBoolean("disabled", false));
                symbol.setLowAlarmEnabled(symbolNode.getBoolean("lowAlarmEnabled", false));
                symbol.setLowAlarmValue(symbolNode.getDouble("lowAlarmValue", 0.0));
                symbol.setLowAlarmIsPercent(symbolNode.getBoolean("lowAlarmIsPercent", false));
                symbol.setLowAlarmSoundEnabled(symbolNode.getBoolean("lowAlarmSoundEnabled", false));
                symbol.setHighAlarmEnabled(symbolNode.getBoolean("highAlarmEnabled", false));
                symbol.setHighAlarmValue(symbolNode.getDouble("highAlarmValue", 0.0));
                symbol.setHighAlarmIsPercent(symbolNode.getBoolean("highAlarmIsPercent", false));
                symbol.setHighAlarmSoundEnabled(symbolNode.getBoolean("highAlarmSoundEnabled", false));
                symbols.add(symbol);
            }
        }
        catch (BackingStoreException e) {
            e.printStackTrace();
        }
        return symbols;
    }
}
