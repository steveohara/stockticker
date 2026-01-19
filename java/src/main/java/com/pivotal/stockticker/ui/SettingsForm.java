package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.Utils;
import com.pivotal.stockticker.model.SettingsManager;
import com.pivotal.stockticker.service.PersistanceManager;
import com.pivotal.stockticker.ui.components.*;
import com.pivotal.stockticker.utils.CallbackInterface;
import com.pivotal.stockticker.utils.StartupManager;

import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.util.Currency;
import java.util.Objects;
import java.util.Set;

/**
 * Settings form for the Stock Ticker application
 */
public class SettingsForm extends JDialog implements CallbackInterface {

    private CapableTextField txtCurrencySymbol, txtMargin, txtTotalInvestment;
    private SettingsButton btnBackground, btnBackup, btnCancel, btnDownArrowColour, btnDownColour, btnLabelColour, btnHighAlarm, btnLowAlarm, btnNormalText, btnOk, btnRestore, btnUpArrowColour, btnUpColour;
    private SettingsCheckbox chkBold, chkItalic, chkShowDailyChange, chkShowTotalCost, chkShowTotalProfit, chkShowTotalProfitPercentage, chkShowTotalValue, chkShowUniqueSymbols;
    private SettingsComboBox<String> lstFont, lstCurrencyCode;
    private SettingsSpinner spnTickerUpdate;
    private SettingsTextField txtAlphaVantageToken, txtFinHubToken, txtFreeCurrencyToken, txtHighAlarm, txtIexToken, txtLowAlarm, txtMarketStackToken, txtProxyServer, txtTiingoToken, txtTwelveDataToken;

    private final SettingsManager settings;
    private final CallbackInterface caller;

    /**
     * Creates new form SettingsForm
     *
     * @param caller   Parent callback interface
     * @param settings Settings object to load and save data
     */
    public SettingsForm(CallbackInterface caller, SettingsManager settings) {
        this.caller = caller;
        this.settings = settings;
        initComponents();
        setTitle("Settings New");
        setModal(true);
        setAlwaysOnTop(true);
        setLocationRelativeTo(null);
        setResizable(false);

        // Init the settings from storage
        loadFromSettings(settings);

        // Initialize listeners
        initListeners();

        // Show the form and select the first button
        setVisible(true);
    }

    /**
     * Handles colour button clicks to open a color chooser dialog.
     *
     * @param e ActionEvent triggered by button click
     */
    private void colourButtonClicked(ActionEvent e) {
        JButton button = (JButton) e.getSource();
        Color selectedColor = JColorChooser.showDialog(button, "Select a Color", button.getBackground());
        if (selectedColor != null) {
            button.setBackground(selectedColor);
            btnOk.setEnabled(true);
        }
    }

    /**
     * Initializes all listeners for the form components.
     */
    private void initListeners() {
        getRootPane().setDefaultButton(btnOk);
        btnOk.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onOK();
            }
        });
        btnCancel.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onCancel();
            }
        });

        // call onCancel() when cross is clicked
        setDefaultCloseOperation(DO_NOTHING_ON_CLOSE);
        addWindowListener(new WindowAdapter() {
            public void windowClosing(WindowEvent e) {
                onCancel();
            }
        });

        // call onCancel() on ESCAPE
        getRootPane().registerKeyboardAction(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onCancel();
            }
        }, KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0), JComponent.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT);

        btnBackground.addActionListener(this::colourButtonClicked);
        btnNormalText.addActionListener(this::colourButtonClicked);
        btnUpColour.addActionListener(this::colourButtonClicked);
        btnDownColour.addActionListener(this::colourButtonClicked);
        btnUpArrowColour.addActionListener(this::colourButtonClicked);
        btnDownArrowColour.addActionListener(this::colourButtonClicked);
        btnLabelColour.addActionListener(this::colourButtonClicked);

        btnLowAlarm.addActionListener(this::selectAudioFile);
        btnHighAlarm.addActionListener(this::selectAudioFile);

        btnBackup.addActionListener(e -> {
            PersistanceManager.backupPreferences(this);
        });
        btnRestore.addActionListener(e -> {
            if (PersistanceManager.restorePreferences(this)) {
                caller.changed(null);
                loadFromSettings(settings);
                btnOk.setEnabled(false);
            }
        });

        // Listen for changes
        Utils.attachChangeListeners(getContentPane(), this);
    }

    /**
     * Set-up the display with data from storage
     *
     * @param settings Settings object to load data from
     */
    private void loadFromSettings(SettingsManager settings) {

        // Colours
        btnBackground.setBackground(settings.getBackgroundColor());
        btnNormalText.setBackground(settings.getNormalTextColor());
        btnUpColour.setBackground(settings.getUpColor());
        btnDownColour.setBackground(settings.getDownColor());
        btnUpArrowColour.setBackground(settings.getUpArrowColor());
        btnDownArrowColour.setBackground(settings.getDownArrowColor());
        btnLabelColour.setBackground(settings.getLabelColor());

        // Fonts
        lstFont.setSelectedItem(settings.getFontName());
        chkBold.setSelected(settings.isFontBold());
        chkItalic.setSelected(settings.isFontItalic());

        // Alarms
        txtHighAlarm.setText(settings.getHighAlarmWaveFile());
        txtLowAlarm.setText(settings.getLowAlarmWaveFile());

        // Display
        spnTickerUpdate.setValue(settings.getFrequency());
        chkShowTotalProfit.setSelected(settings.isShowPortfolioProfitAndLoss());
        chkShowTotalProfitPercentage.setSelected(settings.isShowPortfolioProfitAndLossPercent());
        chkShowTotalCost.setSelected(settings.isShowTotalCost());
        chkShowTotalValue.setSelected(settings.isShowTotalValue());
        chkShowDailyChange.setSelected(settings.isShowDailyChange());
        chkShowUniqueSymbols.setSelected(settings.isShowUniqueSymbols());

        // API Keys
        txtIexToken.setText(settings.getIexToken());
        txtAlphaVantageToken.setText(settings.getAlphaVantageToken());
        txtMarketStackToken.setText(settings.getMarketStackToken());
        txtTwelveDataToken.setText(settings.getTwelveDataToken());
        txtFinHubToken.setText(settings.getFinhubToken());
        txtTiingoToken.setText(settings.getTiingoToken());
        txtFreeCurrencyToken.setText(settings.getFreeCurrencyToken());

        // Other
        txtProxyServer.setText(settings.getProxyServer());
        lstCurrencyCode.setSelectedItem(settings.getCurrencyCode());

        txtCurrencySymbol.setText(settings.getCurrencySymbol());
        txtTotalInvestment.setText(settings.getTotalInvestment());
        txtMargin.setText(settings.getMargin());
    }

    /**
     * Saves the current settings from the form to the Settings object.
     */
    private void saveSettings() {
        // Colours
        settings.setBackgroundColor(btnBackground.getBackground());
        settings.setNormalTextColor(btnNormalText.getBackground());
        settings.setUpColor(btnUpColour.getBackground());
        settings.setDownColor(btnDownColour.getBackground());
        settings.setUpArrowColor(btnUpArrowColour.getBackground());
        settings.setDownArrowColor(btnDownArrowColour.getBackground());
        settings.setLabelColor(btnLabelColour.getBackground());

        // Fonts
        settings.setFontName((String) lstFont.getSelectedItem());
        settings.setFontBold(chkBold.isSelected());
        settings.setFontItalic(chkItalic.isSelected());

        // Alarms
        settings.setHighAlarmWaveFile(txtHighAlarm.getText().trim());
        settings.setLowAlarmWaveFile(txtLowAlarm.getText().trim());

        // Display
        settings.setFrequency((Integer) spnTickerUpdate.getValue());
        settings.setShowPortfolioProfitAndLoss(chkShowTotalProfit.isSelected());
        settings.setShowPortfolioProfitAndLossPercent(chkShowTotalProfitPercentage.isSelected());
        settings.setShowTotalCost(chkShowTotalCost.isSelected());
        settings.setShowTotalValue(chkShowTotalValue.isSelected());
        settings.setShowDailyChange(chkShowDailyChange.isSelected());
        settings.setShowUniqueSymbols(chkShowUniqueSymbols.isSelected());

        // API Keys
        settings.setIexToken(txtIexToken.getText().trim());
        settings.setAlphaVantageToken(txtAlphaVantageToken.getText().trim());
        settings.setMarketStackToken(txtMarketStackToken.getText().trim());
        settings.setTwelveDataToken(txtTwelveDataToken.getText().trim());
        settings.setFinhubToken(txtFinHubToken.getText().trim());
        settings.setTiingoToken(txtTiingoToken.getText().trim());
        settings.setFreeCurrencyToken(txtFreeCurrencyToken.getText().trim());

        // Other
        settings.setProxyServer(txtProxyServer.getText().trim());
        settings.setCurrencyCode(Objects.requireNonNull(lstCurrencyCode.getSelectedItem()).toString());
        settings.setCurrencySymbol(txtCurrencySymbol.getText().trim());
        settings.setMargin(txtMargin.getValue());
        settings.setTotalInvestment(txtTotalInvestment.getValue());
    }

    /**
     * Handles the OK button click event.
     */
    private void onOK() {
        saveSettings();
        caller.changed(this);
        dispose();
    }

    /**
     * Handles the Cancel button click event.
     */
    private void onCancel() {
        if (btnOk.isEnabled()) {
            if (JOptionPane.showConfirmDialog(this, "Discard changes?", "Confirm", JOptionPane.YES_NO_OPTION) == JOptionPane.NO_OPTION) {
                return;
            }
        }
        dispose();
    }

    /**
     * Handles the selection of audio files for alarm sounds.
     *
     * @param e ActionEvent triggered by button click
     */
    private void selectAudioFile(ActionEvent e) {
        JButton button = (JButton) e.getSource();
        JTextField textField = button == btnHighAlarm ? txtHighAlarm : txtLowAlarm;
        String file = Utils.selectAudioFile(this, String.format("Select %s Alarm Sound File", button == btnHighAlarm ? "High" : "Low"), textField.getText());
        if (file != null) {
            textField.setText(file);
        }
    }

    @Override
    public void changed(Component c) {
        btnOk.setEnabled(true);
    }

    /**
     * Initialises all the UI components
     */
    private void initComponents() {

        setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);

        int vGap = 6;
        int hGap = 10;
        int lblGap = 5;
        int width = StartupManager.isWindows() ? 460 :  450;

        setSize(width, 800);
        setPreferredSize(getSize());
        setMaximumSize(getSize());
        setMinimumSize(getSize());

        setResizable(false);
        getContentPane().setLayout(null);

        // Normal Settings
        SettingsLabel jLabel1 = SettingsLabel.create("Proxy Server").atPosition(hGap, vGap).to(getContentPane());
        txtProxyServer = SettingsTextField.create("", "The address of a proxy server to use e.g. www.proxy.com:8989 etc.")
                        .tail(jLabel1, lblGap).withWidth(getWidth() - (jLabel1.getRight() + lblGap + hGap * 3)).to(getContentPane());

        SettingsLabel jLabel2 = SettingsLabel.create("Update Every").below(jLabel1, vGap).to(getContentPane());
        spnTickerUpdate = SettingsSpinner.create(30, 30, 600, 10).tail(jLabel2, lblGap).setTooltip("How often to retrieve prices data (30-600)").to(getContentPane());
        JLabel jLabel3 = SettingsLabel.create("Seconds").setAlignment(Label.LEFT).tail(spnTickerUpdate, hGap).to(getContentPane());

        // Divider
        SettingsSeparator jSeparator1 = SettingsSeparator.create().below(jLabel2, vGap * 3).withWidth(txtProxyServer.getRight() - jLabel2.getX()).to(getContentPane());
        JLabel jLabel30 = SettingsLabel.create("<html><p style='font-weight:bold;color:#808080'>&nbsp;&nbsp;Colours, Fonts & Sounds</p></html>")
                .setAlignment(SwingConstants.LEFT)
                .withWidth(170)
                .atTop(jSeparator1.getY() - vGap - 2)
                .atLeft(jSeparator1.getX() + hGap * 3)
                .setBackColor(getContentPane().getBackground())
                .to(getContentPane());
        jLabel30.setOpaque(true);
        getContentPane().setComponentZOrder(jLabel30, 0);

        // Colour buttons
        SettingsLabel jLabel4 = SettingsLabel.create("Background").below(jLabel30, vGap).atLeft(jLabel1).withDimensions(jLabel1).to(getContentPane());
        btnBackground = SettingsButton.create("").tail(jLabel4, lblGap).withWidth(25).withHeight(jLabel4).setBackColor(new java.awt.Color(0, 0, 0)).to(getContentPane());

        SettingsLabel jLabel5 = SettingsLabel.create("Up Colour").tail(btnBackground, vGap).withHeight(jLabel4).withWidth(75).to(getContentPane());
        btnUpColour = SettingsButton.create("").tail(jLabel5, lblGap).withDimensions(btnBackground).setBackColor(new java.awt.Color(0, 204, 204)).to(getContentPane());

        SettingsLabel jLabel6 = SettingsLabel.create("Up Arrow Colour").tail(btnUpColour, vGap).withHeight(jLabel4).withWidth(115).withHeight(jLabel4).to(getContentPane());
        btnUpArrowColour = SettingsButton.create("").tail(jLabel6, lblGap).withDimensions(btnBackground).setBackColor(new java.awt.Color(0, 204, 51)).to(getContentPane());

        SettingsLabel jLabel7 = SettingsLabel.create("Normal Text").below(jLabel4, vGap).withDimensions(jLabel4).to(getContentPane());
        btnNormalText = SettingsButton.create("").tail(jLabel7, lblGap).withDimensions(btnBackground).setBackColor(new java.awt.Color(153, 153, 153)).to(getContentPane());

        SettingsLabel jLabel8 = SettingsLabel.create("Down Colour").below(jLabel5, vGap).withDimensions(jLabel5).to(getContentPane());
        btnDownColour = SettingsButton.create("").tail(jLabel8, lblGap).withDimensions(btnBackground).setBackColor(new java.awt.Color(255, 102, 102)).to(getContentPane());

        SettingsLabel jLabel9 = SettingsLabel.create("Down Arrow Colour").below(jLabel6, vGap).withDimensions(jLabel6).to(getContentPane());
        btnDownArrowColour = SettingsButton.create("").tail(jLabel9, lblGap).withDimensions(btnBackground).setBackColor(new java.awt.Color(0, 204, 204)).to(getContentPane());

        SettingsLabel jLabel10 = SettingsLabel.create("Label Text").below(jLabel7, vGap).withDimensions(jLabel4).to(getContentPane());
        btnLabelColour = SettingsButton.create("").tail(jLabel10, lblGap).withDimensions(btnBackground).setBackColor(new java.awt.Color(153, 153, 153)).to(getContentPane());

        // Font settings
        SettingsLabel jLabel11 = SettingsLabel.create("Font").below(jLabel10, vGap).to(getContentPane());
        lstFont = SettingsComboBox.<String>create().tail(jLabel11, lblGap).withWidth(150).to(getContentPane());
        chkBold = SettingsCheckbox.create("Bold").tail(lstFont, hGap * 2).withWidth(60).to(getContentPane());
        chkItalic = SettingsCheckbox.create("Italic").tail(chkBold, hGap).withWidth(chkBold).to(getContentPane());

        // Set the font list
        lstFont.addItem("");
        String[] fonts = GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames();
        for (String font : fonts) {
            lstFont.addItem(font.toUpperCase());
            if (font.equalsIgnoreCase(settings.getFontName())) {
                lstFont.setSelectedItem(font);
            }
        }

        SettingsLabel jLabel12 = SettingsLabel.create("High alarm").below(jLabel11, vGap).withDimensions(jLabel1).to(getContentPane());
        txtHighAlarm = SettingsTextField.create().tail(jLabel12, lblGap).withWidth(getWidth() - (jLabel12.getRight() + lblGap + 70)).to(getContentPane());
        btnHighAlarm = SettingsButton.create("...").tail(txtHighAlarm, lblGap).withWidth(25).withHeight(txtHighAlarm).setBackColor(new java.awt.Color(204, 204, 204)).to(getContentPane());

        SettingsLabel jLabel13 = SettingsLabel.create("Low alarm").below(jLabel12, vGap).withDimensions(jLabel1).to(getContentPane());
        txtLowAlarm = SettingsTextField.create().tail(jLabel13, lblGap).withDimensions(txtHighAlarm).to(getContentPane());
        btnLowAlarm = SettingsButton.create("...").tail(txtLowAlarm, lblGap).withDimensions(btnHighAlarm).setBackColor(new java.awt.Color(204, 204, 204)).to(getContentPane());

        // Divider
        SettingsSeparator jSeparator2 = SettingsSeparator.create().below(jLabel13, vGap * 3).withDimensions(jSeparator1).to(getContentPane());
        JLabel jLabel14 = SettingsLabel.create("<html><p style='font-weight:bold;color:#808080'>&nbsp;&nbsp;Summary</p></html>")
                .setAlignment(SwingConstants.LEFT)
                .withWidth(75)
                .atTop(jSeparator2.getY() - vGap - 2)
                .atLeft(jSeparator2.getX() + hGap * 3)
                .setBackColor(getContentPane().getBackground())
                .to(getContentPane());
        jLabel14.setOpaque(true);
        getContentPane().setComponentZOrder(jLabel14, 0);

        // Summary settings
        chkShowTotalProfit = SettingsCheckbox.create("Show Portfolio Profit & Loss", "Displays the total cost of the portfolio (including cash investment)")
                .below(jLabel14, vGap).atLeft(50).withWidth(300).to(getContentPane());
        chkShowTotalProfitPercentage = SettingsCheckbox.create("Show Portfolio Profit & Loss as Percentage", "Show the overall portfolio position as a percentage of the total cost")
                .below(chkShowTotalProfit, vGap / 3).withWidth(chkShowTotalProfit).to(getContentPane());
        chkShowTotalCost = SettingsCheckbox.create("Show Total Portfolio Cost", "Displays the total cost of the portfolio (including cash investment)")
                .below(chkShowTotalProfitPercentage, vGap / 3).withWidth(chkShowTotalProfit).to(getContentPane());
        chkShowTotalValue = SettingsCheckbox.create("Show Total Portfolio Value", "Display the current value of the portfolio (minus cash investment)")
                .below(chkShowTotalCost, vGap / 3).withWidth(chkShowTotalProfit).to(getContentPane());
        chkShowDailyChange = SettingsCheckbox.create("Show Daily Change", "Show the daily summary and day position of the portfolio")
                .below(chkShowTotalValue, vGap / 3).withWidth(chkShowTotalProfit).to(getContentPane());
        chkShowUniqueSymbols = SettingsCheckbox.create("Show Unique Stock Symbols", "Shows a single stock for multiple trades of the same symbol and aggregates the costs (Base Cost)")
                .below(chkShowDailyChange, vGap / 3).withWidth(chkShowTotalProfit).to(getContentPane());

        // Divider
        SettingsSeparator jSeparator3 = SettingsSeparator.create().below(chkShowUniqueSymbols, vGap * 3).atLeft(jSeparator1).withDimensions(jSeparator1).to(getContentPane());
        JLabel jLabel15 = SettingsLabel.create("<html><p style='font-weight:bold;color:#808080'>&nbsp;&nbsp;Currency Conversion & Investment</p></html>")
                .setAlignment(SwingConstants.LEFT)
                .withWidth(240)
                .atTop(jSeparator3.getY() - vGap - 2)
                .atLeft(jSeparator3.getX() + hGap * 3)
                .setBackColor(getContentPane().getBackground())
                .to(getContentPane());
        jLabel15.setOpaque(true);
        getContentPane().setComponentZOrder(jLabel15, 0);

        SettingsLabel jLabel16 = SettingsLabel.create("Currency Code").below(jLabel15, vGap).withDimensions(jLabel1).to(getContentPane());
        lstCurrencyCode = SettingsComboBox.<String>create().setTooltip("ISO Currency to convert summary values into e.g. GBP, USD etc.").tail(jLabel16, lblGap).withWidth(70).to(getContentPane());
        lstCurrencyCode.addItem("");
        Set<Currency> currencies = Currency.getAvailableCurrencies();
        currencies.stream()
            .map(Currency::getCurrencyCode)
            .sorted()
            .forEach(code -> lstCurrencyCode.addItem(code));

        SettingsLabel jLabel17 = SettingsLabel.create("Currency Symbol").tail(lstCurrencyCode, hGap).withDimensions(jLabel1).withWidth(120).to(getContentPane());
        txtCurrencySymbol = CapableTextField.create("", "The display symbol to use. Note: if the symbol is a letter, it is assumed it appends otherwise it prepends e.g. £/$ would go at the front where as c/p will go at the back")
                .tail(jLabel17, lblGap).withWidth(30).atRight(txtProxyServer.getRight()).setConversionType(CapableTextField.CONVERSION_TYPE.UPPER).to(getContentPane());
        jLabel17.atRight(txtCurrencySymbol.getX() - lblGap);

        SettingsLabel jLabel19 = SettingsLabel.create("Total Investment").below(jLabel16, vGap).withDimensions(jLabel1).to(getContentPane());
        txtTotalInvestment = CapableTextField.create("", "Total amount invested (cash paid) in stocks in local currency").tail(jLabel19, lblGap).withWidth(90).setConversionType(CapableTextField.CONVERSION_TYPE.NUMERIC).to(getContentPane());

        SettingsLabel jLabel20 = SettingsLabel.create("Margin").tail(txtTotalInvestment, hGap).withDimensions(jLabel1).withWidth(70).to(getContentPane());
        txtMargin = CapableTextField.create("", "Amount of money in debit (margin) account").tail(jLabel20, lblGap).withDimensions(txtTotalInvestment).atRight(txtProxyServer.getRight()).setConversionType(CapableTextField.CONVERSION_TYPE.NUMERIC).to(getContentPane());
        jLabel20.atRight(txtMargin.getX() - lblGap);

        // Divider
        SettingsSeparator jSeparator5 = SettingsSeparator.create().below(txtTotalInvestment, vGap * 3).atLeft(jSeparator1).withDimensions(jSeparator1).to(getContentPane());
        JLabel jLabel21 = SettingsLabel.create("<html><p style='font-weight:bold;color:#808080'>&nbsp;&nbsp;API Tokens</p></html>")
                .setAlignment(SwingConstants.LEFT)
                .withWidth(90)
                .atTop(jSeparator5.getY() - vGap - 2)
                .atLeft(jSeparator5.getX() + hGap * 3)
                .setBackColor(getContentPane().getBackground())
                .to(getContentPane());
        jLabel21.setOpaque(true);
        getContentPane().setComponentZOrder(jLabel21, 0);

        txtIexToken = SettingsTextField.create().below(jLabel21, vGap / 2).withWidth(300).atRight(txtProxyServer.getRight()).to(getContentPane());
        SettingsLabel jLabel22 = SettingsLabel.create("IEX Token").atTop(txtIexToken).withWidth(130).atRight(txtIexToken.getX() - lblGap).to(getContentPane());

        SettingsLabel jLabel23 = SettingsLabel.create("AlphaVantage Token").below(jLabel22, vGap).to(getContentPane());
        txtAlphaVantageToken = SettingsTextField.create().below(txtIexToken, vGap).to(getContentPane());

        SettingsLabel jLabel24 = SettingsLabel.create("MarketStack Token").below(jLabel23, vGap).to(getContentPane());
        txtMarketStackToken = SettingsTextField.create().below(txtAlphaVantageToken, vGap).to(getContentPane());

        SettingsLabel jLabel25 = SettingsLabel.create("TwelveData Token").below(jLabel24, vGap).to(getContentPane());
        txtTwelveDataToken = SettingsTextField.create().below(txtMarketStackToken, vGap).to(getContentPane());

        SettingsLabel jLabel26 = SettingsLabel.create("Finhub Token").below(jLabel25, vGap).to(getContentPane());
        txtFinHubToken = SettingsTextField.create().below(txtTwelveDataToken, vGap).to(getContentPane());

        SettingsLabel jLabel27 = SettingsLabel.create("Tiingo Token").below(jLabel26, vGap).withDimensions(jLabel22).to(getContentPane());
        txtTiingoToken = SettingsTextField.create().below(txtFinHubToken, vGap).to(getContentPane());

        SettingsLabel jLabel28 = SettingsLabel.create("FreeCurrency Token").below(jLabel27, vGap).withDimensions(jLabel22).to(getContentPane());
        txtFreeCurrencyToken = SettingsTextField.create().below(txtTiingoToken, vGap).to(getContentPane());

        // Buttons
        btnBackup = SettingsButton.create("Backup").below(jLabel28, vGap * 3).atLeft(hGap * 2).withWidth(80).to(getContentPane());
        btnRestore = SettingsButton.create("Restore").tail(btnBackup, hGap).withDimensions(btnBackup).to(getContentPane());

        btnCancel = SettingsButton.create("Cancel").atTop(btnBackup).withDimensions(btnBackup).atRight(txtProxyServer.getRight()).to(getContentPane());
        btnOk = SettingsButton.create("Ok").atTop(btnBackup).withDimensions(btnBackup).atRight(btnCancel.getX() - hGap).to(getContentPane());

        setSize(width, btnOk.getBottom() + (StartupManager.isWindows() ? 50 : 40));
        setPreferredSize(getSize());
        setMaximumSize(getSize());
        setMinimumSize(getSize());

        btnOk.setEnabled(false);
    }

}
