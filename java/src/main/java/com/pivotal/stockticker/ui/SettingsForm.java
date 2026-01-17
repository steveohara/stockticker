package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.Utils;
import com.pivotal.stockticker.model.Settings;
import com.pivotal.stockticker.service.PersistanceManager;
import com.pivotal.stockticker.ui.components.*;
import com.pivotal.stockticker.utils.CallbackInterface;
import com.pivotal.stockticker.utils.StartupManager;

import javax.swing.*;
import java.awt.*;
import java.awt.event.*;

/**
 * Settings form for the Stock Ticker application
 */
public class SettingsForm extends JDialog implements CallbackInterface {

    private CapableTextField txtCurrencyCode, txtMargin, txtTotalInvestment;
    private SettingsButton btnBackground, btnBackup, btnCancel, btnDownArrowColour, btnDownColour, btnLabelColour, btnHighAlarm, btnLowAlarm, btnNormalText, btnOk, btnRestore, btnUpArrowColour, btnUpColour;
    private SettingsCheckbox chkBold, chkItalic, chkShowDailyChange, chkShowTotalCost, chkShowTotalProfit, chkShowTotalProfitPercentage, chkShowTotalValue, chkShowUniqueSymbols;
    private JComboBox<String> lstFont;
    private SettingsSpinner spnTickerUpdate;
    private SettingsTextField txtAlphaVantagToken, txtCurrencySymbol, txtFinHubToken, txtFreeCurrencyToken, txtHighAlarm, txtIexToken, txtLowAlarm, txtMarketStackToken, txtProxyServer, txtTiingoToken, txtTwelveDataToken;

    private final Settings settings;
    private final CallbackInterface caller;

    /**
     * Creates new form SettingsForm
     *
     * @param caller   Parent callback interface
     * @param settings Settings object to load and save data
     */
    public SettingsForm(CallbackInterface caller, Settings settings) {
        this.caller = caller;
        this.settings = settings;
        initComponents();
        setTitle("Settings New");
//      setModal(true);
        setAlwaysOnTop(true);
        setLocationRelativeTo(null);
//        setResizable(false);

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

        btnBackup.addActionListener(e -> {
            setAlwaysOnTop(false);
            PersistanceManager.backupPreferences();
            setAlwaysOnTop(true);
        });
        btnRestore.addActionListener(e -> {
            setAlwaysOnTop(false);
            if (PersistanceManager.restorePreferences()) {
                caller.changed(null);
                loadFromSettings(settings);
                btnOk.setEnabled(false);
            }
            setAlwaysOnTop(true);
        });

        // Listen for changes
        Utils.attachChangeListeners(getContentPane(), this);
    }

    /**
     * Set-up the display with data from storage
     *
     * @param settings Settings object to load data from
     */
    private void loadFromSettings(Settings settings) {

        // Colours
        btnBackground.setBackground(settings.getBackgroundColor());
        btnNormalText.setBackground(settings.getNormalTextColor());
        btnUpColour.setBackground(settings.getUpColor());
        btnDownColour.setBackground(settings.getDownColor());
        btnUpArrowColour.setBackground(settings.getUpArrowColor());
        btnDownArrowColour.setBackground(settings.getDownArrowColor());

        // Fonts
        String[] fonts = GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames();
        for (String font : fonts) {
            lstFont.addItem(font);
            if (font.equalsIgnoreCase(settings.getFontName())) {
                lstFont.setSelectedItem(font);
            }
        }
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
        txtAlphaVantagToken.setText(settings.getAlphaVantageToken());
        txtMarketStackToken.setText(settings.getMarketStackToken());
        txtTwelveDataToken.setText(settings.getTwelveDataToken());
        txtFinHubToken.setText(settings.getFinhubToken());
        txtTiingoToken.setText(settings.getTiingoToken());
        txtFreeCurrencyToken.setText(settings.getFreeCurrencyToken());

        // Other
        txtProxyServer.setText(settings.getProxyServer());
        txtCurrencyCode.setText(settings.getCurrencyCode());
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
        settings.setAlphaVantageToken(txtAlphaVantagToken.getText().trim());
        settings.setMarketStackToken(txtMarketStackToken.getText().trim());
        settings.setTwelveDataToken(txtTwelveDataToken.getText().trim());
        settings.setFinhubToken(txtFinHubToken.getText().trim());
        settings.setTiingoToken(txtTiingoToken.getText().trim());
        settings.setFreeCurrencyToken(txtFreeCurrencyToken.getText().trim());

        // Other
        settings.setProxyServer(txtProxyServer.getText().trim());
        settings.setCurrencyCode(txtCurrencyCode.getText().trim());
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

    @Override
    public void changed(Component c) {
        btnOk.setEnabled(true);
    }

    /**
     * Initialises all the UI components
     */
    private void initComponents() {

        setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);

        setSize(450, 800);
        setPreferredSize(getSize());
        setMaximumSize(getSize());
        setMinimumSize(getSize());

        int vGap = 8;
        int hGap = 10;
        int lblGap = 5;
        int stdWidth = 100;
        int stdHeight = 20;
        int width = StartupManager.isWindows() ? 500 :  450;
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
        SettingsSeparator jSeparator1 = SettingsSeparator.create().below(jLabel2, vGap * 2).withWidth(txtProxyServer.getRight() - jLabel2.getX()).to(getContentPane());
        JLabel jLabel30 = SettingsLabel.create("<html><p style='font-weight:bold;color:#808080'>&nbsp;&nbsp;Colours, Fonts & Sounds</p></html>")
                .setAlignment(SwingConstants.LEFT)
                .withWidth(145)
                .atTop(jSeparator1.getY() - vGap)
                .atLeft(jSeparator1.getX() + hGap * 3)
                .setBackColor(getContentPane().getBackground())
                .to(getContentPane());
        jLabel30.setOpaque(true);
        getContentPane().setComponentZOrder(jLabel30, 0);


        // Colour buttons
        SettingsLabel jLabel4 = SettingsLabel.create("Background").below(jLabel30, vGap).atLeft(jLabel2).withHeight(16).to(getContentPane());
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
        SettingsLabel jLabel11 = SettingsLabel.create("Font").below(jLabel10, vGap).withDimensions(jLabel4).to(getContentPane());


        SettingsLabel jLabel12 = SettingsLabel.create("High alarm").below(jLabel11, vGap).withDimensions(jLabel1).to(getContentPane());
        txtHighAlarm = SettingsTextField.create().tail(jLabel12, lblGap).withWidth(getWidth() - (jLabel12.getRight() + lblGap + 70)).to(getContentPane());


        SettingsLabel jLabel13 = SettingsLabel.create("Low alarm").below(jLabel12, vGap).withDimensions(jLabel1).to(getContentPane());
        txtLowAlarm = SettingsTextField.create().tail(jLabel13, lblGap).withDimensions(txtHighAlarm).to(getContentPane());





        btnOk = SettingsButton.create();

        btnCancel = SettingsButton.create();
        btnRestore = SettingsButton.create();
        JPanel jPanel1 = new JPanel();
        btnDownColour = SettingsButton.create();
        btnUpArrowColour = SettingsButton.create();
        btnDownArrowColour = SettingsButton.create();
        lstFont = new JComboBox<>();
        chkBold = SettingsCheckbox.create("");
        chkItalic = SettingsCheckbox.create("");
        SettingsLabel jLabel26 = SettingsLabel.create();
        btnBackground = SettingsButton.create();
        JPanel jPanel2 = new JPanel();
        chkShowTotalProfitPercentage = SettingsCheckbox.create("");
        chkShowTotalCost = SettingsCheckbox.create("");
        chkShowTotalValue = SettingsCheckbox.create("");
        chkShowDailyChange = SettingsCheckbox.create("");
        chkShowTotalProfit = SettingsCheckbox.create("");
        chkShowUniqueSymbols = SettingsCheckbox.create("");
        JPanel jPanel3 = new JPanel();
        SettingsLabel jLabel15 = SettingsLabel.create();
        txtIexToken = SettingsTextField.create();
        SettingsLabel jLabel16 = SettingsLabel.create();
        txtAlphaVantagToken = SettingsTextField.create();
        SettingsLabel jLabel17 = SettingsLabel.create();
        txtMarketStackToken = SettingsTextField.create();
        SettingsLabel jLabel18 = SettingsLabel.create();
        txtTwelveDataToken = SettingsTextField.create();
        SettingsLabel jLabel19 = SettingsLabel.create();
        txtFinHubToken = SettingsTextField.create();
        SettingsLabel jLabel20 = SettingsLabel.create();
        txtTiingoToken = SettingsTextField.create();
        SettingsLabel jLabel21 = SettingsLabel.create();
        txtFreeCurrencyToken = SettingsTextField.create();
        JPanel jPanel4 = new JPanel();
        SettingsLabel jLabel14 = SettingsLabel.create();
        SettingsLabel jLabel22 = SettingsLabel.create();
        txtCurrencySymbol = SettingsTextField.create();
        txtCurrencyCode = new CapableTextField(CapableTextField.CONVERSION_TYPE.UPPER);
        SettingsLabel jLabel23 = SettingsLabel.create();
        txtMargin = new CapableTextField(CapableTextField.CONVERSION_TYPE.NUMERIC);
        txtTotalInvestment = new CapableTextField(CapableTextField.CONVERSION_TYPE.NUMERIC);
        SettingsLabel jLabel24 = SettingsLabel.create();
        SettingsLabel jLabel25 = SettingsLabel.create();
        JSeparator jSeparator2 = new JSeparator();
        btnBackup = SettingsButton.create();
        JSeparator jSeparator4 = new JSeparator();

        btnOk.setText("OK");
        btnOk.setEnabled(false);

        btnCancel.setText("Cancel");

        btnRestore.setText("Restore");
        btnRestore.setToolTipText("Restore settings and symbols from a local file");


        btnDownArrowColour.setBackground(java.awt.Color.red);

        chkBold.setText("Bold");

        chkItalic.setText("Italic");

//        btnHighAlarm.setBackground(new java.awt.Color(204, 204, 204));
//        btnHighAlarm.setText("...");
//
//        btnLowAlarm.setBackground(new java.awt.Color(204, 204, 204));
//        btnLowAlarm.setText("...");

        jLabel26.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel26.setText("Down Arrow Colour");



        chkShowTotalProfitPercentage.setText("Show Portfolio Profit & Loss as Percentage");
        chkShowTotalProfitPercentage.setToolTipText("Show the overall portfolio position as a percentage of the total cost");

        chkShowTotalCost.setText("Show Total Portfolio Cost");
        chkShowTotalCost.setToolTipText("Displays the total cost of the portfolio (including cash investment)");

        chkShowTotalValue.setText("Show Total Portfolio Value");
        chkShowTotalValue.setToolTipText("Display the current value of the portfolio (minus cash investment)");

        chkShowDailyChange.setText("Show Daily Change");
        chkShowDailyChange.setToolTipText("Show the daily summary and day position of the portfolio");

        chkShowTotalProfit.setText("Show Portfolio Profit & Loss");
        chkShowTotalProfit.setToolTipText("Show an overall position of the portfolio");

        chkShowUniqueSymbols.setText("Show Unique Stock Symbols");
        chkShowUniqueSymbols.setToolTipText("Shows a single stock for multiple trades of the same symbol and aggregates the costs (Base Cost)");

        jLabel2.setHorizontalAlignment(SwingConstants.RIGHT);


        spnTickerUpdate.setToolTipText("How often to retrieve prices data (30-600)");

        jLabel3.setHorizontalAlignment(SwingConstants.LEFT);
        jLabel3.setText("Seconds");


        jLabel15.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel15.setText("IEX Token");

        jLabel16.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel16.setText("AlphaVantage Token");

        jLabel17.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel17.setText("MarketStack Token");

        jLabel18.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel18.setText("TwelveData Token");

        jLabel19.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel19.setText("Finhub Token");

        jLabel20.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel20.setText("Tiingo Token");

        jLabel21.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel21.setText("FreeCurrency Token");

        jLabel11.setFont(new java.awt.Font("Helvetica Neue", 1, 13)); // NOI18N
        jLabel11.setForeground(new java.awt.Color(204, 204, 204));
        jLabel11.setHorizontalAlignment(SwingConstants.CENTER);
        jLabel11.setText("Currency Conversion");

        jLabel14.setText("Currency Code");

        jLabel22.setText("Currency Symbol");

        txtCurrencyCode.setToolTipText("ISO Currency to convert summary values into e.g. GBP, USD etc.");

        jLabel23.setText("Margin");

        txtMargin.setToolTipText("Ampunt of money in debit (margin) account");

        txtTotalInvestment.setToolTipText("Total amount invested in stocks in local currency");

        jLabel24.setFont(new java.awt.Font("Helvetica Neue", 1, 13)); // NOI18N
        jLabel24.setForeground(new java.awt.Color(204, 204, 204));
        jLabel24.setHorizontalAlignment(SwingConstants.CENTER);
        jLabel24.setText("Investment");

        jLabel25.setText("Total Investment");


        txtProxyServer.setToolTipText("The address of a proxy server to use e.g. www.proxy.com:8989 etc.");
        txtProxyServer.setName(""); // NOI18N


        btnBackup.setText("Backup");
        btnBackup.setToolTipText("Backup all settings and symbols to a local file");

    }

}
