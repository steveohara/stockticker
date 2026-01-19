package com.pivotal.stockticker.service;

import com.pivotal.stockticker.Utils;
import com.pivotal.stockticker.model.SettingsManager;
import com.pivotal.stockticker.model.SymbolTransaction;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import net.bytebuddy.ByteBuddy;
import net.bytebuddy.implementation.MethodDelegation;
import net.bytebuddy.implementation.bind.annotation.*;
import net.bytebuddy.matcher.ElementMatchers;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.*;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.nio.charset.MalformedInputException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Base64;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.Callable;
import java.util.prefs.BackingStoreException;
import java.util.prefs.InvalidPreferencesFormatException;
import java.util.prefs.Preferences;

/**
 * Application settings and configuration
 * These changes will be persisted to settings.json automatically
 */
@Slf4j
abstract public class PersistanceManager {

    // Root node for all preferences
    public static final String ROOT_NODE_NAME = "/stockticker";
    public static final String ROOT_NODE = ROOT_NODE_NAME + '/';

    // Cache of constructors that have been created for proxy instances
    private static final Map<String, Constructor> constructors = new HashMap<>();

    // Preferences instance
    private Preferences prefs;

    // File chooser for backup/restore operations
    private static OverwritePromptChooser chooser = null;

    // Auto-save flag - if true, changes are automatically saved to preferences
    @Getter
    @Setter
    private boolean autoSave = true;

    /**
     * Loads the field value from preferences.
     *
     * @param field The field to save.
     */
    private void loadField(Field field) {

        // Ignore non-setable fields
        if (Modifier.isStatic(field.getModifiers()) ||
                Modifier.isFinal(field.getModifiers()) ||
                Modifier.isTransient(field.getModifiers())) {
            return;
        }

        field.setAccessible(true);
        try {
            Class<?> type = field.getType();
            String key = field.getName();
            Object target = field.getDeclaringClass().cast(this);

            if (type == String.class) {
                String value = prefs.get(key, field.get(target) == null ? "" : field.get(target).toString());
                field.set(target, value);
            }
            else if (type == int.class || type == Integer.class) {
                int value = prefs.getInt(key, field.get(target) == null ? 0 : field.getInt(target));
                field.set(target, value);
            }
            else if (type == long.class || type == Long.class) {
                long value = prefs.getLong(key, field.get(target) == null ? 0 : field.getLong(target));
                field.set(target, value);
            }
            else if (type == boolean.class || type == Boolean.class) {
                boolean value = prefs.getBoolean(key, field.get(target) != null && field.getBoolean(target));
                field.set(target, value);
            }
            else if (type == double.class || type == Double.class) {
                double value = prefs.getDouble(key, field.get(target) == null ? 0.0 : field.getDouble(target));
                field.set(target, value);
            }
            else if (type == Color.class || type == Font.class) {
                Object value = deserializeObject(prefs.get(key, null));
                if (value != null) {
                    field.set(target, value);
                }
            }
        }
        catch (IllegalAccessException e) {
            throw new RuntimeException("Failed to set field: " + field.getName(), e);
        }
        catch (NumberFormatException e) {
            throw new RuntimeException("Invalid default value for field: " + field.getName(), e);
        }
    }

    /**
     * Saves the field value to preferences.
     *
     * @param field The field to save.
     * @param value Value to save
     */
    protected void saveField(Field field, Object value) {

        // Ignore non-setable fields
        if (Modifier.isStatic(field.getModifiers()) ||
                Modifier.isFinal(field.getModifiers()) ||
                Modifier.isTransient(field.getModifiers())) {
            return;
        }

        field.setAccessible(true);
        try {
            Class<?> type = field.getType();
            String key = field.getName();

            if (value == null && type != String.class) {
                prefs.remove(key);
            }
            else if (type == String.class) {
                prefs.put(key, value == null ? "" : field.get(this).toString());
            }
            else if (type == int.class || type == Integer.class) {
                prefs.putInt(key, (int) value);
            }
            else if (type == long.class || type == Long.class) {
                prefs.putLong(key, (long) value);
            }
            else if (type == boolean.class || type == Boolean.class) {
                prefs.putBoolean(key, (boolean) value);
            }
            else if (type == double.class || type == Double.class) {
                prefs.putDouble(key, (double) value);
            }
            else if (type == Color.class || type == Font.class) {
                String serialized = serializeObject((Serializable) value);
                prefs.put(key, serialized);
            }
        }
        catch (IllegalAccessException e) {
            throw new RuntimeException("Failed to save field: " + field.getName(), e);
        }
        catch (NumberFormatException e) {
            throw new RuntimeException("Invalid default value for field: " + field.getName(), e);
        }
    }

    /**
     * Serializes an object to a Base64 encoded string.
     *
     * @param obj The object to serialize.
     * @return The Base64 encoded string.
     */
    private static String serializeObject(Serializable obj) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ObjectOutputStream oos = new ObjectOutputStream(baos);
            oos.writeObject(obj);
            oos.close();
            return Base64.getEncoder().encodeToString(baos.toByteArray());
        }
        catch (IOException e) {
            throw new RuntimeException("Failed to serialize object", e);
        }
    }

    /**
     * Deserializes an object from a Base64 encoded string.
     *
     * @param str The Base64 encoded string.
     * @return The deserialized object.
     */
    private static Object deserializeObject(String str) {
        if (str == null || str.isEmpty()) {
            return null;
        }
        try {
            byte[] data = Base64.getDecoder().decode(str);
            ObjectInputStream ois = new ObjectInputStream(new ByteArrayInputStream(data));
            Object obj = ois.readObject();
            ois.close();
            return obj;
        }
        catch (IOException | ClassNotFoundException e) {
            log.error("Failed to deserialize object", e);
            return null;
        }
    }

    /**
     * Checks if a key exists in preferences.
     *
     * @param key The key to check.
     * @return True if the key exists, false otherwise.
     */
    public boolean keyExists(String key) {
        try {
            return Arrays.asList(prefs.childrenNames()).contains(key);
        }
        catch (BackingStoreException e) {
            return false;
        }
    }

    /**
     * Creates a proxy instance of this class so that we can intercept method calls.
     *
     * @param clazz    The class to create a proxy for.
     * @param prefs    Preferences to save the values to/from.
     * @param autoSave Indicates whether to load existing values from storage upon creation and enable auto-saving on changes.
     * @return A proxy instance of this class.
     */
    @SuppressWarnings("unchecked")
    protected static <T> T createProxyInstance(Class<T> clazz, Preferences prefs, boolean autoSave) throws Exception {
        log.debug("Getting proxy for {} from cache", clazz.getSimpleName());
        Constructor<T> constructor;
        synchronized (constructors) {
            constructor = constructors.get(clazz.getName());
            if (constructor == null) {
                log.debug("Creating proxy for {}", clazz.getSimpleName());
                constructor = (Constructor<T>) new ByteBuddy()
                        .subclass(clazz)
                        .method(ElementMatchers.nameStartsWith("set"))
                        .intercept(MethodDelegation.to(ChangeTrackingInterceptor.class))
                        .make()
                        .load(clazz.getClassLoader())
                        .getLoaded()
                        .getDeclaredConstructor();
                constructors.put(clazz.getName(), constructor);
            }
        }
        T instance = constructor.newInstance();

        // Initialize from storage
        log.debug("Loading {} instance {} from storage", clazz.getSimpleName(), instance);
        ((PersistanceManager) instance).prefs = prefs;
        ((PersistanceManager) instance).setAutoSave(autoSave);
        for (Field field : instance.getClass().getSuperclass().getDeclaredFields()) {
            ((PersistanceManager) instance).loadField(field);
        }
        log.debug("Initialised {} instance {} from storage", clazz.getSimpleName(), instance);
        return instance;
    }

    /**
     * Saves all fields to storage into the current instance.
     */
    public void saveToStorage() {
        for (Field field : getClass().getSuperclass().getDeclaredFields()) {
            try {
                // Ignore non-setable fields
                if (!Modifier.isStatic(field.getModifiers()) &&
                        !Modifier.isFinal(field.getModifiers()) &&
                        !Modifier.isTransient(field.getModifiers())) {
                    field.setAccessible(true);
                    this.saveField(field, field.get(this));
                }
            }
            catch (IllegalAccessException e) {
                throw new RuntimeException("Failed to save field: " + field.getName(), e);
            }
        }
    }

    /**
     * Loads all fields from storage into the current instance.
     */
    protected void loadFromStorage(Preferences prefs) {
        this.prefs = prefs;
        for (Field field : getClass().getSuperclass().getDeclaredFields()) {
            this.loadField(field);
        }
    }

    /**
     * Backs up the given Preferences subtree to a user-selected file.
     *
     * @param dialog The parent dialog for the file chooser.
     */
    public static void backupPreferences(Dialog dialog) {

        // Use JFileChooser to let user pick the file location
        if (chooser == null) {
            chooser = new OverwritePromptChooser();
            chooser.setFileFilter(new FileNameExtensionFilter("Backup Files (*.bck)", "bck"));
        }
        chooser.setChooserType(OverwritePromptChooser.CHOOSER_TYPE.SAVE);
        chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        chooser.setAcceptAllFileFilterUsed(false);
        chooser.setMultiSelectionEnabled(false);
        chooser.setDialogType(JFileChooser.SAVE_DIALOG);
        chooser.setDialogTitle("Backup Settings");
        chooser.setApproveButtonText("Save");
        int userSelection = chooser.showSaveDialog(dialog);

        // If user approved, export the preferences to the selected file
        if (userSelection == JFileChooser.APPROVE_OPTION) {

            // Get the file and make sure it has an extension
            File selectedFile = chooser.getSelectedFile();
            if (!selectedFile.getName().toLowerCase().endsWith(".bck")) {
                selectedFile = new File(selectedFile.getAbsolutePath() + ".bck");
            }

            // Get the file from the chooser and export the preferences
            try (OutputStream os = Files.newOutputStream(selectedFile.toPath())) {
                Preferences prefs = Preferences.userRoot().node(ROOT_NODE_NAME);
                prefs.exportSubtree(os);
                Utils.showTopmostMessage("Settings backed up successfully!", "Backup Successful", JOptionPane.INFORMATION_MESSAGE);
                log.info("Preferences backed up to {}", selectedFile);
            }
            catch (IOException | BackingStoreException ex) {
                Utils.showTopmostMessage("Export failed: " + ex.getMessage(), "Backup Failed", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    /**
     * Restores Preferences subtree from a user-selected file.
     *
     * @param dialog The parent dialog for the file chooser.
     * @return true if restore was successful, false otherwise.
     */
    public static boolean restorePreferences(JDialog dialog) {

        // Use JFileChooser to let user pick the file location
        if (chooser == null) {
            chooser = new OverwritePromptChooser();
            chooser.setFileFilter(new FileNameExtensionFilter("Backup Files (*.bck)", "bck"));
        }
        chooser.setChooserType(OverwritePromptChooser.CHOOSER_TYPE.OPEN);
        chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        chooser.setAcceptAllFileFilterUsed(false);
        chooser.setMultiSelectionEnabled(false);
        chooser.setDialogType(JFileChooser.OPEN_DIALOG);
        chooser.setDialogTitle("Restore Settings");
        chooser.setApproveButtonText("Open");
        int userSelection = chooser.showSaveDialog(dialog);

        // If user approved, import the preferences from the selected file
        if (userSelection == JFileChooser.APPROVE_OPTION) {

            // Get the file and make sure it has an extension
            File selectedFile = chooser.getSelectedFile();
            if (selectedFile == null || !selectedFile.exists()) {
                Utils.showTopmostMessage("Selected file does not exist.", "Restore Failed", JOptionPane.ERROR_MESSAGE);
                return false;
            }

            // Load the preferences from the file
            try {
                loadPreferencesFromFile(selectedFile);
                Utils.showTopmostMessage("Settings restored successfully!", "Restore Successful", JOptionPane.INFORMATION_MESSAGE);
                log.info("Preferences restored from {}", chooser.getSelectedFile());
                return true;
            }
            catch (Exception ex) {
                log.error("Restore from {} failed", selectedFile, ex);
                Utils.showTopmostMessage("Restore failed: " + ex.getMessage(), "Restore Failed", JOptionPane.ERROR_MESSAGE);
            }
        }
        return false;
    }

    /**
     * Loads preferences from the given file, determining the format automatically.
     *
     * @param file The file to load preferences from.
     */
    private static void loadPreferencesFromFile(File file) throws Exception {

        // Read the file contents in one go
        String content;
        try {
            content = Files.readString(Path.of(file.getAbsolutePath()), StandardCharsets.UTF_8);
        }
        catch (MalformedInputException mie) {
            content = Files.readString(Path.of(file.getAbsolutePath()), StandardCharsets.ISO_8859_1);
        }

        // Get some transient values that need to be maintained
        SettingsManager settings = SettingsManager.getPersistentSettings();
        int left = settings.getWindowX();
        int top = settings.getWindowY();
        int width = settings.getWindowWidth();

        // Determine the format and load accordingly
        if (content.startsWith("Windows Registry Editor Version 5.00")) {
            loadPreferencesFromRegistryString(content);
        }
        else {
            loadPreferencesFromBackupString(content);
        }

        // Replace the positioning settings to maintain window position
        settings = SettingsManager.getPersistentSettings();
        settings.setWindowX(left);
        settings.setWindowY(top);
        settings.setWindowWidth(width);
    }

    /**
     * Loads preferences from a backup string.
     *
     * @param content The backup string to load preferences from.
     */
    private static void loadPreferencesFromBackupString(String content) throws IOException, BackingStoreException, InvalidPreferencesFormatException {

        // Get the file from the chooser and export the preferences
        try (InputStream is = new ByteArrayInputStream(content.getBytes())) {

            // Load the preferences from the content
            Preferences prefs = Preferences.userRoot().node(ROOT_NODE_NAME);
            prefs.removeNode();
            Preferences.importPreferences(is);
        }
    }

    /**
     * Loads preferences from a Windows Registry export string.
     *
     * @param content The registry export string to load preferences from.
     * @throws Exception if an error occurs during loading.
     */
    private static void loadPreferencesFromRegistryString(String content) throws Exception {

        // Delete the preferences node first
        Preferences prefs = Preferences.userRoot().node(ROOT_NODE_NAME);
        prefs.removeNode();

        // Now create new managers to load the data into
        SettingsManager settings = SettingsManager.getPersistentSettings();
        SymbolsManager symbolsManager = new SymbolsManager();
        PricesManager pricesManager = new PricesManager(settings);
        ExchangeRatesManager exchangeRatesManager = new ExchangeRatesManager(settings);

        // Split the content into lines
        String[] lines = content.split("\\r?\\n");
        for (int line = 0; line < lines.length; line++) {

            // If this is the settings section, load settings
            if (lines[line].equalsIgnoreCase("[HKEY_CURRENT_USER\\SOFTWARE\\Pivotal\\StockTicker\\Settings]")) {
                line = loadSettingsFromRegistryLines(settings, lines, line);
            }

            // If this is a symbol section, load symbol
            else if (lines[line].startsWith("[HKEY_CURRENT_USER\\SOFTWARE\\Pivotal\\StockTicker\\Symbols\\")) {
                line = loadSymbolFromRegistryLines(symbolsManager, pricesManager, exchangeRatesManager, lines, line);
            }
        }
    }

    /**
     * Loads settings from registry lines starting at the given line index.
     *
     * @param settings The Settings instance to load into.
     * @param lines    The registry lines.
     * @param line     The starting line index.
     * @return The next line index after processing.
     */
    private static int loadSettingsFromRegistryLines(SettingsManager settings, String[] lines, int line) {

        // Loop through all the lines until we hit another section
        int row;
        for (row = line + 1; row < lines.length && !lines[row].startsWith("["); row++) {

            // Check we have something to process
            if (lines[row].trim().isEmpty()) {
                continue;
            }

            // Get the key and value
            String key = lines[row].split("=", 2)[0].replace("\"", "");
            String value = lines[row].split("=", 2)[1].replace("\"", "");

            // Create a new Settings entry
            switch (key.toLowerCase()) {
                case "proxy":
                    settings.setProxyServer(value);
                    break;
                case "frequency":
                    settings.setFrequency(Integer.parseInt(value));
                    break;
                case "alwaysontop":
                    settings.setAlwaysOnTop(Boolean.parseBoolean(value));
                    break;

                case "background colour":
                    settings.setBackgroundColor(vb6ColorToJavaColor(value));
                    break;
                case "text colour":
                    settings.setNormalTextColor(vb6ColorToJavaColor(value));
                    break;
                case "up colour":
                    settings.setUpColor(vb6ColorToJavaColor(value));
                    break;
                case "down colour":
                    settings.setDownColor(vb6ColorToJavaColor(value));
                    break;
                case "up arrow colour":
                    settings.setUpArrowColor(vb6ColorToJavaColor(value));
                    break;
                case "down arrow colour":
                    settings.setDownArrowColor(vb6ColorToJavaColor(value));
                    break;

                case "font":
                    settings.setFontName(value);
                    break;
                case "bold":
                    settings.setFontBold(Boolean.parseBoolean(value));
                    break;
                case "italic":
                    settings.setFontItalic(Boolean.parseBoolean(value));
                    break;
                case "font size":
                    int size = Integer.parseInt(value);
                    if (size < SettingsManager.FONT_SIZE_SMALL) {
                        size = SettingsManager.FONT_SIZE_SMALL;
                    }
                    else if (size < SettingsManager.FONT_SIZE_MEDIUM) {
                        size = SettingsManager.FONT_SIZE_MEDIUM;
                    }
                    else {
                        size = SettingsManager.FONT_SIZE_LARGE;
                    }
                    settings.setFontSize(size);
                    break;

                case "show total profit and loss":
                    settings.setShowPortfolioProfitAndLoss(Boolean.parseBoolean(value));
                    break;
                case "show total profit and loss as percent":
                    settings.setShowPortfolioProfitAndLossPercent(Boolean.parseBoolean(value));
                    break;
                case "show total cost":
                    settings.setShowTotalCost(Boolean.parseBoolean(value));
                    break;
                case "show total value":
                    settings.setShowTotalValue(Boolean.parseBoolean(value));
                    break;

                case "show daily change":
                    settings.setShowDailyChange(Boolean.parseBoolean(value));
                    break;
                case "summarise":
                    settings.setShowUniqueSymbols(Boolean.parseBoolean(value));
                    break;

                case "currency":
                    settings.setCurrencyCode(value);
                    break;
                case "currency symbol":
                    settings.setCurrencySymbol(value);
                    break;

                case "total investment":
                    settings.setTotalInvestment(Integer.parseInt(value));
                    break;
                case "margin":
                    settings.setMargin(Integer.parseInt(value));
                    break;

                case "day summary sort order":
                    settings.setDaySortOrder(value);
                    break;
                case "day summary sort column":
                    settings.setDaySortColumn(value);
                    break;
                case "summary sort order":
                    settings.setSummarySortOrder(value);
                    break;
                case "summary sort column":
                    settings.setSummarySortColumn(value);
                    break;

                case "high alarm wave file":
                    settings.setHighAlarmWaveFile(value);
                    break;
                case "low alarm wave file":
                    settings.setHighAlarmWaveFile(value);
                    break;

                case "iex key":
                    settings.setIexToken(value);
                    break;
                case "alphavantage key":
                    settings.setAlphaVantageToken(value);
                    break;
                case "marketstack key":
                    settings.setMarketStackToken(value);
                    break;
                case "twelvedata key":
                    settings.setTwelveDataToken(value);
                    break;
                case "finhub key":
                    settings.setFinhubToken(value);
                    break;
                case "tiingo key":
                    settings.setTiingoToken(value);
                    break;
                case "freecurrency key":
                    settings.setFreeCurrencyToken(value);
                    break;

                case "scroll speed":
                    int speed = Integer.parseInt(value.replaceFirst("^[^,]+,", ""));
                    if (speed < SettingsManager.SCROLL_SPEED_SLOW) {
                        speed = SettingsManager.SCROLL_SPEED_SLOW;
                    }
                    else if (speed < SettingsManager.SCROLL_SPEED_MEDIUM) {
                        speed = SettingsManager.SCROLL_SPEED_MEDIUM;
                    }
                    else {
                        speed = SettingsManager.SCROLL_SPEED_FAST;
                    }
                    settings.setTickerSpeed(speed);
                    break;
            }
        }
        return row - 1;
    }

    /**
     * Loads a symbol from registry lines starting at the given line index.
     *
     * @param symbolsManager The SymbolsManager to add the symbol to.
     * @param pricesManager  The PricesManager to add prices to.
     * @param ratesManager   The ExchangeRatesManager to add exchange rates to.
     * @param lines          The registry lines.
     * @param line           The starting line index.
     * @return The next line index after processing.
     */
    private static int loadSymbolFromRegistryLines(SymbolsManager symbolsManager, PricesManager pricesManager, ExchangeRatesManager ratesManager, String[] lines, int line) {

        // Create a new symbol instance
        String timestamp = lines[line].replaceAll("(^.+\\\\)|(])", "");

        // We need to convert this timestamp to a unix version for the key
        // this timestamp is actually the number of seconds since 1st Jan 2008
        timestamp = (Long.parseLong(timestamp) + 1199145600L) * 1000 + ""; // Convert to milliseconds
        SymbolTransaction symbol = symbolsManager.createNewSymbolTransaction(timestamp);

        // Loop through all the lines until we hit another section
        int row;
        for (row = line + 1; row < lines.length && !lines[row].startsWith("["); row++) {

            // Check we have something to process
            if (lines[row].trim().isEmpty()) {
                continue;
            }

            // Get the key and value
            String key = lines[row].split("=", 2)[0].replace("\"", "");
            String value = lines[row].split("=", 2)[1].replace("\"", "");

            // Create a new SymbolTransaction entry
            switch (key.toLowerCase()) {
                case "symbol":
                    symbol.setCode(value);
                    break;
                case "alias":
                    symbol.setAlias(value);
                    break;
                case "disabled":
                    symbol.setDisabled(Boolean.parseBoolean(value));
                    break;
                case "price":
                    symbol.setPricePaid(Double.parseDouble(value));
                    break;
                case "currency":
                    symbol.setCurrencyCode(value);
                    break;
                case "currency symbol":
                    symbol.setCurrencySymbol(value);
                    break;
                case "shares":
                    symbol.setSharesBought(Double.parseDouble(value));
                    break;

                case "show price":
                    symbol.setShowPrice(Boolean.parseBoolean(value));
                    break;
                case "exclude from summary":
                    symbol.setExcludeFromSummary(Boolean.parseBoolean(value));
                    break;
                case "show change":
                    symbol.setShowChange(Boolean.parseBoolean(value));
                    break;
                case "show change percent":
                    symbol.setShowChangePercent(Boolean.parseBoolean(value));
                    break;
                case "show change indicator":
                    symbol.setShowChangeUpDown(Boolean.parseBoolean(value));
                    break;
                case "show profit and loss":
                    symbol.setShowProfitLoss(Boolean.parseBoolean(value));
                    break;

                case "show day change":
                    symbol.setShowDayChange(Boolean.parseBoolean(value));
                    break;
                case "show day change percent":
                    symbol.setShowDayChangePercent(Boolean.parseBoolean(value));
                    break;
                case "show day change indicator":
                    symbol.setShowDayChangeUpDown(Boolean.parseBoolean(value));
                    break;

                case "low alarm enabled":
                    symbol.setLowAlarmEnabled(Boolean.parseBoolean(value));
                    break;
                case "low alarm as percent":
                    symbol.setLowAlarmIsPercent(Boolean.parseBoolean(value));
                    break;
                case "low alarm sound enabled":
                    symbol.setLowAlarmSoundEnabled(Boolean.parseBoolean(value));
                    break;
                case "low alarm value":
                    symbol.setLowAlarmValue(Double.parseDouble(value));
                    break;

                case "high alarm enabled":
                    symbol.setHighAlarmEnabled(Boolean.parseBoolean(value));
                    break;
                case "high alarm as percent":
                    symbol.setHighAlarmIsPercent(Boolean.parseBoolean(value));
                    break;
                case "high alarm sound enabled":
                    symbol.setHighAlarmSoundEnabled(Boolean.parseBoolean(value));
                    break;
                case "high alarm value":
                    symbol.setHighAlarmValue(Double.parseDouble(value));
                    break;
            }
        }

        // Save the symbol
        log.info("Added symbol from registry: {}", symbol);
        pricesManager.addPrice(symbol.getCode());
        ratesManager.addRate(symbol.getCurrencySymbol());
        symbolsManager.persistChanges();
        return row - 1;
    }

    /**
     * Converts a VB6 OLE color (BBGGRR) to a java.awt.Color.
     */
    private static Color vb6ColorToJavaColor(String vb6ColorString) {
        int vb6Color = Integer.parseInt(vb6ColorString);
        int r = vb6Color & 0xFF;
        int g = (vb6Color >> 8) & 0xFF;
        int b = (vb6Color >> 16) & 0xFF;
        return new Color(r, g, b);
    }

    /**
     * SettingsManager class to handle persisting settings changes
     */
    public static class ChangeTrackingInterceptor {
        @RuntimeType
        public static Object intercept(@This Object self,
                                       @Origin Method method,
                                       @AllArguments Object[] args,
                                       @SuperCall Callable<?> zuper) throws Exception {

            // Set the field value first
            Object result = zuper.call();

            // If we are auto-saving, persist the change
            if (((PersistanceManager) self).isAutoSave()) {
                if (args != null && method.getName().startsWith("set")) {
                    String fieldName = method.getName().substring(3);

                    // Find the corresponding field name
                    for (Field field : self.getClass().getSuperclass().getDeclaredFields()) {
                        if (field.getName().equalsIgnoreCase(fieldName)) {
                            field.setAccessible(true);
                            ((PersistanceManager) self).saveField(field, args[0]);
                            break;
                        }
                    }
                    log.debug("Field changed: {}", fieldName);
                }
            }
            return result;
        }
    }

    /**
     * Custom JFileChooser that prompts for overwrite confirmation
     */
    @Getter
    @Setter
    private static class OverwritePromptChooser extends JFileChooser {
        public enum CHOOSER_TYPE {
            OPEN,
            SAVE
        }

        private CHOOSER_TYPE chooserType = CHOOSER_TYPE.OPEN;

        @Override
        public void approveSelection() {
            File file = getSelectedFile();

            // If we are saving, check if the file exists
            if (chooserType == CHOOSER_TYPE.SAVE) {
                if (file != null && file.exists()) {
                    int answer = JOptionPane.showConfirmDialog(
                            this,
                            "File \"" + file.getName() + "\" already exists.\nOverwrite?",
                            "Confirm Overwrite",
                            JOptionPane.YES_NO_OPTION,
                            JOptionPane.QUESTION_MESSAGE);
                    if (answer != JOptionPane.YES_OPTION) {
                        return;
                    }
                }
            }

            // If we are loading, then confirm overwrite of settings
            else {
                int answer = JOptionPane.showConfirmDialog(
                        this,
                        "You are about to overwrite your settings from \"" + file.getName() + "\".\nAre you sure?",
                        "Confirm Overwrite",
                        JOptionPane.YES_NO_OPTION,
                        JOptionPane.QUESTION_MESSAGE);
                if (answer != JOptionPane.YES_OPTION) {
                    return;
                }
            }
            super.approveSelection();
        }
    }
}
