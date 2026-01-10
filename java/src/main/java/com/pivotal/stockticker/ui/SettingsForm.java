package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.Utils;
import com.pivotal.stockticker.model.Settings;

import javax.swing.*;
import java.awt.*;
import java.awt.event.*;

/**
 * Settings form for the Stock Ticker application
 */
public class SettingsForm extends JDialog implements CallbackInterface {

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
        setTitle("Settings");
        setModal(true);
        setAlwaysOnTop(true);
        setLocationRelativeTo(null);
        setResizable(false);

        // Init the settings from storage
        loadFromSettings(settings);

        // Initialize listeners
        initListeners();
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
        txtTotalInvestment.setText(String.valueOf(settings.getTotalInvestment()));
        txtMargin.setText(String.valueOf(settings.getMargin()));
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
        settings.setMargin(Utils.parseDouble(txtMargin.getText(), 0));
        settings.setTotalInvestment(Utils.parseDouble(txtTotalInvestment.getText(), 0));
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
        btnOk = new JButton();
        btnCancel = new JButton();
        btnRestore = new JButton();
        jTabbedPane1 = new JTabbedPane();
        jPanel1 = new JPanel();
        btnNormalText = new JButton();
        btnUpColour = new JButton();
        btnDownColour = new JButton();
        btnUpArrowColour = new JButton();
        btnDownArrowColour = new JButton();
        jLabel10 = new JLabel();
        lstFont = new JComboBox<>();
        chkBold = new JCheckBox();
        chkItalic = new JCheckBox();
        jLabel12 = new JLabel();
        jLabel4 = new JLabel();
        txtHighAlarm = new JTextField();
        jLabel5 = new JLabel();
        btnHighAlarm = new JButton();
        jLabel6 = new JLabel();
        jLabel13 = new JLabel();
        jLabel7 = new JLabel();
        txtLowAlarm = new JTextField();
        jLabel8 = new JLabel();
        btnLowAlarm = new JButton();
        jLabel26 = new JLabel();
        btnBackground = new JButton();
        jPanel2 = new JPanel();
        chkShowTotalProfitPercentage = new JCheckBox();
        chkShowTotalCost = new JCheckBox();
        chkShowTotalValue = new JCheckBox();
        chkShowDailyChange = new JCheckBox();
        chkShowTotalProfit = new JCheckBox();
        chkShowUniqueSymbols = new JCheckBox();
        jLabel2 = new JLabel();
        spnTickerUpdate = new JSpinner(new SpinnerNumberModel(30, 30, 600, 10));
        jLabel3 = new JLabel();
        jPanel3 = new JPanel();
        jLabel15 = new JLabel();
        txtIexToken = new JTextField();
        jLabel16 = new JLabel();
        txtAlphaVantagToken = new JTextField();
        jLabel17 = new JLabel();
        txtMarketStackToken = new JTextField();
        jLabel18 = new JLabel();
        txtTwelveDataToken = new JTextField();
        jLabel19 = new JLabel();
        txtFinHubToken = new JTextField();
        jLabel20 = new JLabel();
        txtTiingoToken = new JTextField();
        jLabel21 = new JLabel();
        txtFreeCurrencyToken = new JTextField();
        jPanel4 = new JPanel();
        jLabel11 = new JLabel();
        jLabel14 = new JLabel();
        jLabel22 = new JLabel();
        txtCurrencySymbol = new JTextField();
        txtCurrencyCode = new JTextField();
        jLabel23 = new JLabel();
        txtMargin = new JTextField();
        txtTotalInvestment = new JTextField();
        jLabel24 = new JLabel();
        jLabel25 = new JLabel();
        jSeparator1 = new JSeparator();
        jLabel1 = new JLabel();
        txtProxyServer = new JTextField();
        jSeparator2 = new JSeparator();
        btnBackup = new JButton();
        jSeparator4 = new JSeparator();

        setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);

        btnOk.setText("OK");
        btnOk.setEnabled(false);

        btnCancel.setText("Cancel");

        btnRestore.setText("Restore");
        btnRestore.setToolTipText("Restore settings and syymbols from a local file");

        jTabbedPane1.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);

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

        GroupLayout jPanel1Layout = new GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
                jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.TRAILING)
                                        .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING, false)
                                                .addGroup(jPanel1Layout.createSequentialGroup()
                                                        .addComponent(jLabel12, GroupLayout.PREFERRED_SIZE, 75, GroupLayout.PREFERRED_SIZE)
                                                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                        .addComponent(txtHighAlarm, GroupLayout.PREFERRED_SIZE, 281, GroupLayout.PREFERRED_SIZE)
                                                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                        .addComponent(btnHighAlarm, GroupLayout.PREFERRED_SIZE, 24, GroupLayout.PREFERRED_SIZE))
                                                .addGroup(jPanel1Layout.createSequentialGroup()
                                                        .addComponent(jLabel13, GroupLayout.PREFERRED_SIZE, 75, GroupLayout.PREFERRED_SIZE)
                                                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                        .addComponent(txtLowAlarm, GroupLayout.PREFERRED_SIZE, 281, GroupLayout.PREFERRED_SIZE)
                                                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                        .addComponent(btnLowAlarm, GroupLayout.PREFERRED_SIZE, 24, GroupLayout.PREFERRED_SIZE)))
                                        .addGroup(jPanel1Layout.createSequentialGroup()
                                                .addComponent(jLabel10, GroupLayout.PREFERRED_SIZE, 104, GroupLayout.PREFERRED_SIZE)
                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                .addComponent(lstFont, GroupLayout.PREFERRED_SIZE, 146, GroupLayout.PREFERRED_SIZE)
                                                .addGap(18, 18, 18)
                                                .addComponent(chkBold)
                                                .addPreferredGap(LayoutStyle.ComponentPlacement.UNRELATED)
                                                .addComponent(chkItalic)
                                                .addGap(35, 35, 35))
                                        .addGroup(jPanel1Layout.createSequentialGroup()
                                                .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.TRAILING)
                                                        .addGroup(jPanel1Layout.createSequentialGroup()
                                                                .addComponent(jLabel4, GroupLayout.PREFERRED_SIZE, 87, GroupLayout.PREFERRED_SIZE)
                                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                                .addComponent(btnBackground, GroupLayout.PREFERRED_SIZE, 24, GroupLayout.PREFERRED_SIZE))
                                                        .addGroup(jPanel1Layout.createSequentialGroup()
                                                                .addComponent(jLabel7, GroupLayout.PREFERRED_SIZE, 87, GroupLayout.PREFERRED_SIZE)
                                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                                .addComponent(btnNormalText, GroupLayout.PREFERRED_SIZE, 24, GroupLayout.PREFERRED_SIZE)))
                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                                                        .addGroup(GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                                                .addComponent(jLabel5, GroupLayout.PREFERRED_SIZE, 94, GroupLayout.PREFERRED_SIZE)
                                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                                .addComponent(btnUpColour, GroupLayout.PREFERRED_SIZE, 24, GroupLayout.PREFERRED_SIZE))
                                                        .addGroup(jPanel1Layout.createSequentialGroup()
                                                                .addComponent(jLabel8, GroupLayout.PREFERRED_SIZE, 94, GroupLayout.PREFERRED_SIZE)
                                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                                .addComponent(btnDownColour, GroupLayout.PREFERRED_SIZE, 24, GroupLayout.PREFERRED_SIZE)))
                                                .addGap(18, 18, 18)
                                                .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                                                        .addGroup(GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                                                .addComponent(jLabel6, GroupLayout.PREFERRED_SIZE, 111, GroupLayout.PREFERRED_SIZE)
                                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                                .addComponent(btnUpArrowColour, GroupLayout.PREFERRED_SIZE, 24, GroupLayout.PREFERRED_SIZE))
                                                        .addGroup(jPanel1Layout.createSequentialGroup()
                                                                .addComponent(jLabel26)
                                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                                .addComponent(btnDownArrowColour, GroupLayout.PREFERRED_SIZE, 24, GroupLayout.PREFERRED_SIZE)))))
                                .addGap(0, 26, Short.MAX_VALUE))
        );
        jPanel1Layout.setVerticalGroup(
                jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGap(24, 24, 24)
                                .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                                        .addGroup(jPanel1Layout.createSequentialGroup()
                                                .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING, false)
                                                        .addComponent(btnUpColour, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                                        .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                                                                .addComponent(jLabel5)
                                                                .addComponent(jLabel4)
                                                                .addComponent(btnBackground, GroupLayout.PREFERRED_SIZE, 16, GroupLayout.PREFERRED_SIZE)))
                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                                                        .addComponent(jLabel8)
                                                        .addComponent(btnDownColour, GroupLayout.PREFERRED_SIZE, 17, GroupLayout.PREFERRED_SIZE)
                                                        .addComponent(jLabel7)
                                                        .addComponent(btnNormalText, GroupLayout.PREFERRED_SIZE, 17, GroupLayout.PREFERRED_SIZE)))
                                        .addGroup(jPanel1Layout.createSequentialGroup()
                                                .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING, false)
                                                        .addComponent(jLabel6, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                                        .addComponent(btnUpArrowColour, GroupLayout.PREFERRED_SIZE, 17, GroupLayout.PREFERRED_SIZE))
                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                                                        .addComponent(btnDownArrowColour, GroupLayout.PREFERRED_SIZE, 17, GroupLayout.PREFERRED_SIZE)
                                                        .addComponent(jLabel26))))
                                .addGap(33, 33, 33)
                                .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                                        .addComponent(jLabel10)
                                        .addComponent(lstFont, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                                        .addComponent(chkBold)
                                        .addComponent(chkItalic))
                                .addGap(18, 18, 18)
                                .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                                        .addComponent(jLabel12)
                                        .addComponent(txtHighAlarm, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                                        .addComponent(btnHighAlarm))
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                                        .addComponent(jLabel13)
                                        .addComponent(txtLowAlarm, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                                        .addComponent(btnLowAlarm))
                                .addContainerGap(55, Short.MAX_VALUE))
        );

        jTabbedPane1.addTab("Colours, Fonts & Sounds", jPanel1);

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
        jLabel2.setText("Update Every");

        spnTickerUpdate.setToolTipText("How often to retrieve prices data (30-600)");

        jLabel3.setHorizontalAlignment(SwingConstants.LEFT);
        jLabel3.setText("Seconds");

        GroupLayout jPanel2Layout = new GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
                jPanel2Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                        .addGroup(GroupLayout.Alignment.TRAILING, jPanel2Layout.createSequentialGroup()
                                .addContainerGap(46, Short.MAX_VALUE)
                                .addGroup(jPanel2Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                                        .addGroup(jPanel2Layout.createSequentialGroup()
                                                .addGap(21, 21, 21)
                                                .addGroup(jPanel2Layout.createParallelGroup(GroupLayout.Alignment.TRAILING)
                                                        .addComponent(chkShowUniqueSymbols, GroupLayout.PREFERRED_SIZE, 343, GroupLayout.PREFERRED_SIZE)
                                                        .addComponent(chkShowTotalProfitPercentage, GroupLayout.PREFERRED_SIZE, 343, GroupLayout.PREFERRED_SIZE)
                                                        .addComponent(chkShowTotalProfit, GroupLayout.PREFERRED_SIZE, 343, GroupLayout.PREFERRED_SIZE)
                                                        .addComponent(chkShowTotalCost, GroupLayout.PREFERRED_SIZE, 343, GroupLayout.PREFERRED_SIZE)
                                                        .addComponent(chkShowTotalValue, GroupLayout.PREFERRED_SIZE, 343, GroupLayout.PREFERRED_SIZE)
                                                        .addComponent(chkShowDailyChange, GroupLayout.PREFERRED_SIZE, 343, GroupLayout.PREFERRED_SIZE)))
                                        .addGroup(jPanel2Layout.createSequentialGroup()
                                                .addComponent(jLabel2, GroupLayout.PREFERRED_SIZE, 104, GroupLayout.PREFERRED_SIZE)
                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                .addComponent(spnTickerUpdate, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                .addComponent(jLabel3, GroupLayout.PREFERRED_SIZE, 62, GroupLayout.PREFERRED_SIZE)))
                                .addGap(37, 37, 37))
        );
        jPanel2Layout.setVerticalGroup(
                jPanel2Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel2Layout.createSequentialGroup()
                                .addGap(25, 25, 25)
                                .addGroup(jPanel2Layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                                        .addComponent(jLabel2)
                                        .addComponent(spnTickerUpdate, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                                        .addComponent(jLabel3))
                                .addGap(18, 18, 18)
                                .addComponent(chkShowTotalProfit)
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(chkShowTotalProfitPercentage)
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(chkShowTotalCost)
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(chkShowTotalValue)
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(chkShowDailyChange)
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(chkShowUniqueSymbols)
                                .addContainerGap(23, Short.MAX_VALUE))
        );

        jTabbedPane1.addTab("Display", jPanel2);

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

        GroupLayout jPanel3Layout = new GroupLayout(jPanel3);
        jPanel3.setLayout(jPanel3Layout);
        jPanel3Layout.setHorizontalGroup(
                jPanel3Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel3Layout.createSequentialGroup()
                                .addGap(30, 30, 30)
                                .addGroup(jPanel3Layout.createParallelGroup(GroupLayout.Alignment.TRAILING, false)
                                        .addComponent(jLabel20, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addComponent(jLabel18, GroupLayout.Alignment.LEADING, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addComponent(jLabel17, GroupLayout.Alignment.LEADING, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addComponent(jLabel16, GroupLayout.Alignment.LEADING, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addGroup(jPanel3Layout.createSequentialGroup()
                                                .addGap(0, 0, Short.MAX_VALUE)
                                                .addComponent(jLabel15, GroupLayout.PREFERRED_SIZE, 104, GroupLayout.PREFERRED_SIZE))
                                        .addComponent(jLabel19, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addComponent(jLabel21, GroupLayout.PREFERRED_SIZE, 119, GroupLayout.PREFERRED_SIZE))
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addGroup(jPanel3Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                                        .addComponent(txtTiingoToken)
                                        .addComponent(txtFinHubToken)
                                        .addComponent(txtTwelveDataToken)
                                        .addComponent(txtFreeCurrencyToken, GroupLayout.Alignment.TRAILING)
                                        .addComponent(txtAlphaVantagToken)
                                        .addComponent(txtMarketStackToken)
                                        .addComponent(txtIexToken))
                                .addGap(36, 36, 36))
        );
        jPanel3Layout.setVerticalGroup(
                jPanel3Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel3Layout.createSequentialGroup()
                                .addGap(18, 18, 18)
                                .addGroup(jPanel3Layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                                        .addComponent(jLabel15)
                                        .addComponent(txtIexToken, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addGroup(jPanel3Layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                                        .addComponent(jLabel16)
                                        .addComponent(txtAlphaVantagToken, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addGroup(jPanel3Layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                                        .addComponent(jLabel17)
                                        .addComponent(txtMarketStackToken, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addGroup(jPanel3Layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                                        .addComponent(jLabel18)
                                        .addComponent(txtTwelveDataToken, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addGroup(jPanel3Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                                        .addComponent(txtFinHubToken, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                                        .addGroup(jPanel3Layout.createSequentialGroup()
                                                .addGap(3, 3, 3)
                                                .addComponent(jLabel19, GroupLayout.PREFERRED_SIZE, 14, GroupLayout.PREFERRED_SIZE)))
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addGroup(jPanel3Layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                                        .addComponent(jLabel20)
                                        .addComponent(txtTiingoToken, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addGroup(jPanel3Layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                                        .addComponent(jLabel21)
                                        .addComponent(txtFreeCurrencyToken, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
                                .addContainerGap(GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        jTabbedPane1.addTab("API Keys", jPanel3);

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

        jLabel1.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel1.setText("Proxy Server");

        txtProxyServer.setToolTipText("The address of a proxy server to use e.g. www.proxy.com:8989 etc.");
        txtProxyServer.setName(""); // NOI18N

        GroupLayout jPanel4Layout = new GroupLayout(jPanel4);
        jPanel4.setLayout(jPanel4Layout);
        jPanel4Layout.setHorizontalGroup(
                jPanel4Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel4Layout.createSequentialGroup()
                                .addGap(24, 24, 24)
                                .addComponent(jLabel1, GroupLayout.PREFERRED_SIZE, 98, GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(txtProxyServer, GroupLayout.PREFERRED_SIZE, 270, GroupLayout.PREFERRED_SIZE)
                                .addGap(0, 49, Short.MAX_VALUE))
                        .addGroup(jPanel4Layout.createSequentialGroup()
                                .addGroup(jPanel4Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                                        .addGroup(jPanel4Layout.createSequentialGroup()
                                                .addGap(23, 23, 23)
                                                .addGroup(jPanel4Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                                                        .addGroup(jPanel4Layout.createSequentialGroup()
                                                                .addGap(12, 12, 12)
                                                                .addGroup(jPanel4Layout.createParallelGroup(GroupLayout.Alignment.TRAILING)
                                                                        .addComponent(jSeparator1)
                                                                        .addGroup(GroupLayout.Alignment.LEADING, jPanel4Layout.createSequentialGroup()
                                                                                .addComponent(jLabel14)
                                                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                                                .addComponent(txtCurrencyCode, GroupLayout.PREFERRED_SIZE, 59, GroupLayout.PREFERRED_SIZE)
                                                                                .addGap(38, 38, 38)
                                                                                .addComponent(jLabel22)
                                                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                                                .addComponent(txtCurrencySymbol, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)
                                                                                .addGap(0, 0, Short.MAX_VALUE))))
                                                        .addComponent(jSeparator2)
                                                        .addGroup(jPanel4Layout.createSequentialGroup()
                                                                .addComponent(jLabel25)
                                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                                .addComponent(txtTotalInvestment, GroupLayout.PREFERRED_SIZE, 98, GroupLayout.PREFERRED_SIZE)
                                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                                                .addComponent(jLabel23)
                                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                                .addComponent(txtMargin, GroupLayout.PREFERRED_SIZE, 74, GroupLayout.PREFERRED_SIZE)
                                                                .addGap(44, 44, 44))))
                                        .addGroup(jPanel4Layout.createSequentialGroup()
                                                .addContainerGap()
                                                .addComponent(jLabel24, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                                        .addGroup(jPanel4Layout.createSequentialGroup()
                                                .addContainerGap()
                                                .addComponent(jLabel11, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                                .addContainerGap())
        );
        jPanel4Layout.setVerticalGroup(
                jPanel4Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel4Layout.createSequentialGroup()
                                .addGap(24, 24, 24)
                                .addGroup(jPanel4Layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                                        .addComponent(jLabel1)
                                        .addComponent(txtProxyServer, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
                                .addGap(18, 18, 18)
                                .addComponent(jSeparator2, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jLabel11)
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addGroup(jPanel4Layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                                        .addComponent(jLabel14)
                                        .addComponent(jLabel22)
                                        .addComponent(txtCurrencySymbol, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                                        .addComponent(txtCurrencyCode, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
                                .addGap(24, 24, 24)
                                .addComponent(jSeparator1, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                                .addGap(7, 7, 7)
                                .addComponent(jLabel24)
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addGroup(jPanel4Layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                                        .addComponent(jLabel25)
                                        .addComponent(jLabel23)
                                        .addComponent(txtMargin, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                                        .addComponent(txtTotalInvestment, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
                                .addContainerGap(45, Short.MAX_VALUE))
        );

        jTabbedPane1.addTab("Misc.", jPanel4);

        btnBackup.setText("Backup");
        btnBackup.setToolTipText("Backup all settings and symbols to a local file");

        GroupLayout layout = new GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
                layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                        .addGroup(layout.createSequentialGroup()
                                .addGap(16, 16, 16)
                                .addComponent(btnBackup)
                                .addPreferredGap(LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(btnRestore)
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(btnOk)
                                .addPreferredGap(LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(btnCancel)
                                .addGap(20, 20, 20))
                        .addComponent(jTabbedPane1)
                        .addGroup(GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                                .addComponent(jSeparator4)
                                .addContainerGap())
        );
        layout.setVerticalGroup(
                layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                        .addGroup(layout.createSequentialGroup()
                                .addComponent(jTabbedPane1, GroupLayout.PREFERRED_SIZE, 280, GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(jSeparator4, GroupLayout.PREFERRED_SIZE, 10, GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                .addGroup(layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                                        .addComponent(btnOk)
                                        .addComponent(btnCancel)
                                        .addComponent(btnBackup)
                                        .addComponent(btnRestore))
                                .addContainerGap(14, Short.MAX_VALUE))
        );
        pack();

        // Make numeric text fields only accept numbers
        Utils.makeTextFieldNumeric(txtTotalInvestment);
        Utils.makeTextFieldNumeric(txtMargin);
    }

    private JButton btnBackground;
    private JButton btnBackup;
    private JButton btnCancel;
    private JButton btnDownArrowColour;
    private JButton btnDownColour;
    private JButton btnHighAlarm;
    private JButton btnLowAlarm;
    private JButton btnNormalText;
    private JButton btnOk;
    private JButton btnRestore;
    private JButton btnUpArrowColour;
    private JButton btnUpColour;
    private JCheckBox chkBold;
    private JCheckBox chkItalic;
    private JCheckBox chkShowUniqueSymbols;
    private JCheckBox chkShowDailyChange;
    private JCheckBox chkShowTotalCost;
    private JCheckBox chkShowTotalProfit;
    private JCheckBox chkShowTotalProfitPercentage;
    private JCheckBox chkShowTotalValue;
    private JLabel jLabel1;
    private JLabel jLabel10;
    private JLabel jLabel11;
    private JLabel jLabel12;
    private JLabel jLabel13;
    private JLabel jLabel14;
    private JLabel jLabel15;
    private JLabel jLabel16;
    private JLabel jLabel17;
    private JLabel jLabel18;
    private JLabel jLabel19;
    private JLabel jLabel2;
    private JLabel jLabel20;
    private JLabel jLabel21;
    private JLabel jLabel22;
    private JLabel jLabel23;
    private JLabel jLabel24;
    private JLabel jLabel25;
    private JLabel jLabel3;
    private JLabel jLabel4;
    private JLabel jLabel5;
    private JLabel jLabel6;
    private JLabel jLabel7;
    private JLabel jLabel8;
    private JPanel jPanel1;
    private JPanel jPanel2;
    private JPanel jPanel3;
    private JPanel jPanel4;
    private JSeparator jSeparator1;
    private JSeparator jSeparator2;
    private JSeparator jSeparator4;
    private JTabbedPane jTabbedPane1;
    private JComboBox<String> lstFont;
    private JLabel jLabel26;
    private JSpinner spnTickerUpdate;
    private JTextField txtAlphaVantagToken;
    private JTextField txtCurrencyCode;
    private JTextField txtCurrencySymbol;
    private JTextField txtFinHubToken;
    private JTextField txtFreeCurrencyToken;
    private JTextField txtHighAlarm;
    private JTextField txtIexToken;
    private JTextField txtLowAlarm;
    private JTextField txtMargin;
    private JTextField txtMarketStackToken;
    private JTextField txtProxyServer;
    private JTextField txtTiingoToken;
    private JTextField txtTotalInvestment;
    private JTextField txtTwelveDataToken;

}
