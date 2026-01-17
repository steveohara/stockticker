package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.Utils;
import com.pivotal.stockticker.model.Settings;
import com.pivotal.stockticker.service.PersistanceManager;
import com.pivotal.stockticker.ui.components.CapableTextField;
import com.pivotal.stockticker.ui.components.SettingsLabel;
import com.pivotal.stockticker.ui.components.SettingsSpinner;
import com.pivotal.stockticker.ui.components.SettingsTextField;
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
    private JButton btnBackground, btnBackup, btnCancel, btnDownArrowColour, btnDownColour, btnHighAlarm, btnLowAlarm, btnNormalText, btnOk, btnRestore, btnUpArrowColour, btnUpColour;
    private JCheckBox chkBold, chkItalic, chkShowDailyChange, chkShowTotalCost, chkShowTotalProfit, chkShowTotalProfitPercentage, chkShowTotalValue, chkShowUniqueSymbols;
    private JComboBox<String> lstFont;
    private JSpinner spnTickerUpdate;
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

        int vGap = 10;
        int hGap = 10;
        int lblGap = 5;
        int stdWidth = 100;
        int stdHeight = 20;
        int width = StartupManager.isWindows() ? 460 :  450;
        setResizable(false);
        getContentPane().setLayout(null);

        SettingsLabel jLabel1 = SettingsLabel.create("Proxy Server").atPosition(hGap, vGap).to(getContentPane());
        txtProxyServer = SettingsTextField.create("", "The address of a proxy server to use e.g. www.proxy.com:8989 etc.")
                        .beside(jLabel1, lblGap).withWidth(getWidth() - jLabel1.getRight() - hGap - lblGap).to(getContentPane());

        SettingsLabel jLabel2 = SettingsLabel.create("Update Every").below(jLabel1, vGap).to(getContentPane());
        spnTickerUpdate = SettingsSpinner.create(30, 30, 600, 10).beside(jLabel2, lblGap).setTooltip("How often to retrieve prices data (30-600)").to(getContentPane());
        JLabel jLabel3 = SettingsLabel.create("Seconds").setAlignment(Label.LEFT).beside(spnTickerUpdate, hGap).to(getContentPane());









        btnOk = new JButton();

        btnCancel = new JButton();
        btnRestore = new JButton();
        JPanel jPanel1 = new JPanel();
        btnNormalText = new JButton();
        btnUpColour = new JButton();
        btnDownColour = new JButton();
        btnUpArrowColour = new JButton();
        btnDownArrowColour = new JButton();
        SettingsLabel jLabel10 = SettingsLabel.create();
        lstFont = new JComboBox<>();
        chkBold = new JCheckBox();
        chkItalic = new JCheckBox();
        SettingsLabel jLabel12 = SettingsLabel.create();
        SettingsLabel jLabel4 = SettingsLabel.create();
        txtHighAlarm = SettingsTextField.create();
        SettingsLabel jLabel5 = SettingsLabel.create();
        btnHighAlarm = new JButton();
        SettingsLabel jLabel6 = SettingsLabel.create();
        SettingsLabel jLabel13 = SettingsLabel.create();
        SettingsLabel jLabel7 = SettingsLabel.create();
        txtLowAlarm = SettingsTextField.create();
        SettingsLabel jLabel8 = SettingsLabel.create();
        btnLowAlarm = new JButton();
        SettingsLabel jLabel26 = SettingsLabel.create();
        btnBackground = new JButton();
        JPanel jPanel2 = new JPanel();
        chkShowTotalProfitPercentage = new JCheckBox();
        chkShowTotalCost = new JCheckBox();
        chkShowTotalValue = new JCheckBox();
        chkShowDailyChange = new JCheckBox();
        chkShowTotalProfit = new JCheckBox();
        chkShowUniqueSymbols = new JCheckBox();
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
        SettingsLabel jLabel11 = SettingsLabel.create();
        SettingsLabel jLabel14 = SettingsLabel.create();
        SettingsLabel jLabel22 = SettingsLabel.create();
        txtCurrencySymbol = SettingsTextField.create();
        txtCurrencyCode = new CapableTextField(CapableTextField.CONVERSION_TYPE.UPPER);
        SettingsLabel jLabel23 = SettingsLabel.create();
        txtMargin = new CapableTextField(CapableTextField.CONVERSION_TYPE.NUMERIC);
        txtTotalInvestment = new CapableTextField(CapableTextField.CONVERSION_TYPE.NUMERIC);
        SettingsLabel jLabel24 = SettingsLabel.create();
        SettingsLabel jLabel25 = SettingsLabel.create();
        JSeparator jSeparator1 = new JSeparator();
        JSeparator jSeparator2 = new JSeparator();
        btnBackup = new JButton();
        JSeparator jSeparator4 = new JSeparator();

        btnOk.setText("OK");
        btnOk.setEnabled(false);

        btnCancel.setText("Cancel");

        btnRestore.setText("Restore");
        btnRestore.setToolTipText("Restore settings and syymbols from a local file");

        btnNormalText.setBackground(new java.awt.Color(153, 153, 153));

        btnUpColour.setBackground(new java.awt.Color(0, 204, 204));

        btnDownColour.setBackground(new java.awt.Color(255, 102, 102));

        btnUpArrowColour.setBackground(new java.awt.Color(0, 204, 51));

        btnDownArrowColour.setBackground(java.awt.Color.red);

        jLabel10.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel10.setText("Font");

        chkBold.setText("Bold");

        chkItalic.setText("Italic");

        jLabel12.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel12.setText("High Alarm");

        jLabel4.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel4.setText("Background");

        jLabel5.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel5.setText("Up Colour");

        btnHighAlarm.setBackground(new java.awt.Color(204, 204, 204));
        btnHighAlarm.setText("...");

        jLabel6.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel6.setText("Up Arrow Colour");

        jLabel13.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel13.setText("Low Alarm");

        jLabel7.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel7.setText("Normal Text");

        jLabel8.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel8.setText("Down Colour");

        btnLowAlarm.setBackground(new java.awt.Color(204, 204, 204));
        btnLowAlarm.setText("...");

        jLabel26.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel26.setText("Down Arrow Colour");

        btnBackground.setBackground(new java.awt.Color(0, 0, 0));


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
