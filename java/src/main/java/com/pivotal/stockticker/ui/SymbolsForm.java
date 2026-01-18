/*
 *
 * Copyright (c) 2026, 4NG and/or its affiliates. All rights reserved.
 * 4NG PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.Utils;
import com.pivotal.stockticker.model.Settings;
import com.pivotal.stockticker.model.SymbolTransaction;
import com.pivotal.stockticker.service.SymbolsManager;
import com.pivotal.stockticker.ui.components.*;
import com.pivotal.stockticker.utils.CallbackInterface;
import com.pivotal.stockticker.utils.StartupManager;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import javax.swing.border.LineBorder;
import javax.swing.event.ListSelectionEvent;
import java.awt.*;
import java.awt.event.*;

/**
 * Form to manage stock symbols
 */
@Slf4j
public class SymbolsForm extends JDialog implements CallbackInterface {

    private SettingsButton btnAdd, btnCancel, btnDelete, btnOk;
    private SettingsCheckbox chkHideDisabled, chkAlarmHighPercent, chkAlarmHighPlaySound, chkAlarmLowPercent, chkAlarmLowPlaySound, chkDisabled, chkExcludeFromSummary, chkShowChange, chkShowChangePercent, chkShowDayChange, chkShowDayChangePercent, chkShowDayUpDown, chkShowPrice, chkShowProfitLoss, chlShowUpDown;
    private CheckBoxFrame pnlAlarmLow, pnlAlarmHigh;
    private SymbolsList lstSymbols;
    private CapableTextField txtAlarmHigh, txtAlarmLow, txtCurrencyCode, txtPricePaid, txtSharesBought, txtSymbol;
    private SettingsTextField txtCurrencySymbol, txtDisplayName;
    private JLabel lblTransactionTimestamp;

    private final Settings settings;
    private final SymbolsManager symbolsManager;
    private final CallbackInterface caller;
    private boolean ignoreChanges = false;

    /**
     * Creates new form Symbols
     */
    public SymbolsForm(Settings settings, CallbackInterface caller, SymbolsManager symbolsManager) {
        this.settings = settings;
        this.caller = caller;
        this.symbolsManager = symbolsManager;
        initComponents();
        setTitle("Symbols");
        setModal(true);
        setAlwaysOnTop(true);
        setLocationRelativeTo(null);
        setResizable(false);

        // Initialize listeners
        initListeners();

        // Init the settings from storage
        loadFromStorage();
        disableForm(this);

        // Show the form and select the first item in the list
        if (lstSymbols.getModel().getSize() > 0) {
            lstSymbols.setSelectedIndex(0);
        }
        setVisible(true);
    }

    /**
     * Recursively sets the enabled state of all components within the form.
     *
     * @param c The container whose components' enabled state is to be set.
     */
    private void disableForm(Container c) {
        for (Component comp : c.getComponents()) {
            if (comp != btnAdd && comp != btnOk && comp != btnCancel && comp != lstSymbols && comp != chkHideDisabled) {
                comp.setEnabled(lstSymbols.getSelectedListItem() != null);
            }
            if (comp instanceof Container) {
                disableForm((Container) comp);
            }
        }
    }

    /**
     * Initializes all listeners for the form components.
     */
    private void initListeners() {

        // Set default buttons, escape to cancel and window close handling
        getRootPane().setDefaultButton(btnOk);
        btnOk.addActionListener(e -> onOK());
        btnCancel.addActionListener(e -> onCancel());
        setDefaultCloseOperation(DO_NOTHING_ON_CLOSE);
        getRootPane().registerKeyboardAction(e -> onCancel(), KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0), JComponent.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT);
        addWindowListener(new WindowAdapter() {
            public void windowClosing(WindowEvent e) {
                onCancel();
            }
        });

        // Handle the symbols
        btnAdd.addActionListener(e -> addNewSymbolTransaction());
        btnDelete.addActionListener(this::deleteSymbolTransaction);
        lstSymbols.addListSelectionListener(this::selectSymbolTransaction);
        lstSymbols.addKeyListener(new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent e) {
                if (e.getKeyCode() == KeyEvent.VK_DELETE) {
                    deleteSymbolTransaction(null);
                }
            }
        });

        // Handle hide disabled symbols
        chkHideDisabled.addActionListener( e -> {
            settings.setHideDisabledSymbols(chkHideDisabled.isSelected());
            lstSymbols.hideDisabled(chkHideDisabled.isSelected());
        });
        chkDisabled.addActionListener( e -> {
            int index = lstSymbols.getSelectedIndex();
            lstSymbols.hideDisabled(lstSymbols.isHideDisabled());
//            if (lstSymbols.isHideDisabled() && chkDisabled.isSelected()) {
//                if (!lstSymbols.getModel().isEmpty()) {
//                    lstSymbols.clearSelection();
//                    lstSymbols.setSelectedIndex(30);
////                    SwingUtilities.invokeLater( () -> {
////                        lstSymbols.setSelectedIndex(index < lstSymbols.getModel().getSize() ? index : lstSymbols.getModel().getSize() - 1);
////                    });
//                }
//            }
        });

        // Listen for changes
        Utils.attachChangeListeners(getContentPane(), this, chkHideDisabled);
    }

    /**
     * Handles the selection of a symbol transaction from the list.
     *
     * @param e ListSelectionEvent
     */
    private void selectSymbolTransaction(ListSelectionEvent e) {
        if (!e.getValueIsAdjusting()) {
            SymbolTransaction symbol = lstSymbols.getSelectedListItem();
            if (symbol != null) {
                SwingUtilities.invokeLater(() -> {
                    disableForm(this);
                    displaySymbolTransaction(symbol);
                });
            }
            else {
                disableForm(this);
            }
            clearDisplay();
        }
    }

    /**
     * Clear the display fields
     */
    private void clearDisplay() {
        ignoreChanges = true;
        txtSymbol.setText("");
        chkDisabled.setSelected(false);
        txtDisplayName.setText("");
        txtPricePaid.setText("");
        txtSharesBought.setText("");
        txtCurrencyCode.setText("");
        txtCurrencySymbol.setText("");

        chkShowPrice.setSelected(false);
        chkExcludeFromSummary.setSelected(false);
        chkShowChange.setSelected(false);
        chkShowChangePercent.setSelected(false);
        chlShowUpDown.setSelected(false);
        chkShowProfitLoss.setSelected(false);

        chkShowDayChangePercent.setSelected(false);
        chkShowDayChange.setSelected(false);
        chkShowDayUpDown.setSelected(false);

        pnlAlarmLow.setSelected(false);
        chkAlarmLowPercent.setSelected(false);
        chkAlarmLowPlaySound.setSelected(false);
        txtAlarmLow.setText("");

        pnlAlarmHigh.setSelected(false);
        chkAlarmHighPercent.setSelected(false);
        chkAlarmHighPlaySound.setSelected(false);
        txtAlarmHigh.setText("");
        lblTransactionTimestamp.setText("");
        lblTransactionTimestamp.setToolTipText("");
        ignoreChanges = false;
    }

    /**
     * Display the selected symbol transaction details in the form fields
     *
     * @param symbol SymbolTransaction to display
     */
    private void displaySymbolTransaction(SymbolTransaction symbol) {
        ignoreChanges = true;
        txtSymbol.setText(symbol.getCode());
        chkDisabled.setSelected(symbol.isDisabled());
        txtDisplayName.setText(symbol.getAlias());
        txtPricePaid.setText(symbol.getPricePaid());
        txtSharesBought.setText(symbol.getSharesBought());
        txtCurrencyCode.setText(symbol.getCurrencyCode());
        txtCurrencySymbol.setText(symbol.getCurrencySymbol());

        chkShowPrice.setSelected(symbol.isShowPrice());
        chkExcludeFromSummary.setSelected(symbol.isExcludeFromSummary());
        chkShowChange.setSelected(symbol.isShowChange());
        chkShowChangePercent.setSelected(symbol.isShowChangePercent());
        chlShowUpDown.setSelected(symbol.isShowChangeUpDown());
        chkShowProfitLoss.setSelected(symbol.isShowProfitLoss());

        chkShowDayChangePercent.setSelected(symbol.isShowDayChangePercent());
        chkShowDayChange.setSelected(symbol.isShowDayChange());
        chkShowDayUpDown.setSelected(symbol.isShowDayChangeUpDown());

        pnlAlarmLow.setSelected(symbol.isLowAlarmEnabled());
        chkAlarmLowPercent.setSelected(symbol.isLowAlarmIsPercent());
        chkAlarmLowPlaySound.setSelected(symbol.isLowAlarmSoundEnabled());
        txtAlarmLow.setText(symbol.getLowAlarmValue());

        pnlAlarmHigh.setSelected(symbol.isHighAlarmEnabled());
        chkAlarmHighPercent.setSelected(symbol.isHighAlarmIsPercent());
        chkAlarmHighPlaySound.setSelected(symbol.isHighAlarmSoundEnabled());
        txtAlarmHigh.setText(symbol.getHighAlarmValue());
        lblTransactionTimestamp.setText(String.format("<html><p style='color:#c0c0c0;font-size:0.9em'>%s</p></html>", symbol.getDisplayTimestamp()));
        lblTransactionTimestamp.setToolTipText(String.format("<html><b>Added: </b>%s</html>", symbol.getFullDisplayTimestamp()));
        ignoreChanges = false;
    }

    /**
     * Delete the selected symbol transaction from the list
     * and selects the next one
     */
    private void deleteSymbolTransaction(ActionEvent e) {
        SymbolTransaction symbol = lstSymbols.getSelectedListItem();
        if (symbol != null) {
            symbolsManager.markSymbolTransactionAsDeleted(symbol);
            int index = lstSymbols.getSelectedIndex();
            lstSymbols.removeItem(symbol);
            if (!lstSymbols.getModel().isEmpty()) {
                lstSymbols.setSelectedIndex(index < lstSymbols.getModel().getSize() ? index : lstSymbols.getModel().getSize() - 1);
            }
            btnOk.setEnabled(true);
            log.debug("Deleted symbol transaction {}", lstSymbols.getSelectedListItem());
        }
    }

    /**
     * Add a new symbol transaction to the list
     */
    private void addNewSymbolTransaction() {
        SymbolTransaction newItem = lstSymbols.addItem(symbolsManager.createNewSymbolTransaction());
        lstSymbols.setSelectedValue(newItem, true);
        btnOk.setEnabled(true);
        txtSymbol.requestFocus();
    }

    /**
     * Set-up the display with data from storage
     */
    private void loadFromStorage() {
        lstSymbols.clear();
        for (SymbolTransaction symbolTransaction : symbolsManager.getSymbolTransactions()) {
            lstSymbols.addItem(symbolTransaction);
        }
        if (lstSymbols.getModel().getSize() > 0) {
            lstSymbols.setSelectedIndex(0);
        }
    }

    /**
     * Handles the OK button click event.
     */
    private void onOK() {
        saveSymbolChanges();
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
        symbolsManager.clearChanges();
        dispose();
    }

    @Override
    public void changed(Component c) {
        if (ignoreChanges) {
            return;
        }
        btnOk.setEnabled(true);
        if (lstSymbols.getSelectedListItem() != null) {
            SymbolTransaction symbol = lstSymbols.getSelectedListItem();
            symbol.setCode(txtSymbol.getText());
            symbol.setDisabled(chkDisabled.isSelected());
            symbol.setAlias(txtDisplayName.getText());
            symbol.setPricePaid(txtPricePaid.getValue());
            symbol.setSharesBought(txtSharesBought.getValue());
            symbol.setCurrencyCode(txtCurrencyCode.getText());
            symbol.setCurrencySymbol(txtCurrencySymbol.getText());

            symbol.setShowPrice(chkShowPrice.isSelected());
            symbol.setExcludeFromSummary(chkExcludeFromSummary.isSelected());
            symbol.setShowChange(chkShowChange.isSelected());
            symbol.setShowChangePercent(chkShowChangePercent.isSelected());
            symbol.setShowChangeUpDown(chlShowUpDown.isSelected());
            symbol.setShowProfitLoss(chkShowProfitLoss.isSelected());

            symbol.setShowDayChangePercent(chkShowDayChangePercent.isSelected());
            symbol.setShowDayChange(chkShowDayChange.isSelected());
            symbol.setShowDayChangeUpDown(chkShowDayUpDown.isSelected());

            symbol.setLowAlarmEnabled(pnlAlarmLow.isSelected());
            symbol.setLowAlarmIsPercent(chkAlarmLowPercent.isSelected());
            symbol.setLowAlarmSoundEnabled(chkAlarmLowPlaySound.isSelected());
            symbol.setLowAlarmValue(txtAlarmLow.getValue());

            symbol.setHighAlarmEnabled(pnlAlarmHigh.isSelected());
            symbol.setHighAlarmIsPercent(chkAlarmHighPercent.isSelected());
            symbol.setHighAlarmSoundEnabled(chkAlarmHighPlaySound.isSelected());
            symbol.setHighAlarmValue(txtAlarmHigh.getValue());

            symbolsManager.markSymbolTransactionAsModified(symbol);
            lstSymbols.repaint();
        }
    }

    /**
     * Saves the current settings from the form to the Settings object.
     */
    private void saveSymbolChanges() {
        symbolsManager.persistChanges();
    }

    /**
     * This method is called from within the constructor to initialize the form.
     */
    private void initComponents() {
        int vGap = 7;
        int hGap = 10;
        int lblGap = 5;
        int stdWidth = 100;
        int stdHeight = 20;
        int width = StartupManager.isWindows() ? 610 :  600;

        // Set the dialog size and use null layout
        setResizable(false);
        getContentPane().setLayout(null);
        LineBorder lineBorder = new LineBorder(UIManager.getColor("Component.borderColor"), 1);
        chkHideDisabled = SettingsCheckbox.create("Hide Disabled Symbols").atLeft(hGap).atTop(vGap).withWidth(150).to(getContentPane());
        chkHideDisabled.setFont(new Font(chkHideDisabled.getFont().getName(), Font.PLAIN, 10));
        chkHideDisabled.setSelected(settings.isHideDisabledSymbols());

        // List of Symbols
        lstSymbols = SymbolsList.create().hideDisabled(settings.isHideDisabledSymbols()).withBorder(null);
        SettingsScrollPane jScrollPane1 = SettingsScrollPane.create().withBorder(lineBorder).withViewport(lstSymbols)
                .at(hGap, chkHideDisabled.getY() + chkHideDisabled.getHeight() + vGap / 2).withDimensions(150, 400)
                .to(getContentPane());

        int left = jScrollPane1.getX() + jScrollPane1.getWidth() + (int)(hGap * 1.5);

        // Symbol & Details Labels/Fields
        JLabel jLabel1 = SettingsLabel.create("Symbol").at(left, hGap).withDimensions(stdWidth, stdHeight).to(getContentPane());
        txtSymbol = CapableTextField.create().setConversionType(CapableTextField.CONVERSION_TYPE.UPPER).tail(jLabel1, lblGap).to(getContentPane());
        lblTransactionTimestamp = SettingsLabel.create().setAlignment(SwingConstants.CENTER).tail(txtSymbol, 0).withWidth(120).to(getContentPane());
        chkDisabled = SettingsCheckbox.create("Disabled", "Do not use this symbol").tail(lblTransactionTimestamp, 0).to(getContentPane());

        // Display Name
        JLabel jLabel2 = SettingsLabel.create("Display Name").below(jLabel1, vGap).to(getContentPane());
        txtDisplayName = SettingsTextField.create("", "Name you want to appear instead of symbol").tail(jLabel2, lblGap).withWidth(300).to(getContentPane());

        // Price Paid & Shares Bought
        JLabel jLabel3 = SettingsLabel.create("Price Paid").below(jLabel2, vGap).to(getContentPane());
        txtPricePaid = CapableTextField.create().setConversionType(CapableTextField.CONVERSION_TYPE.NUMERIC).tail(jLabel3, lblGap).withWidth(70).to(getContentPane());

        txtSharesBought = CapableTextField.create().setConversionType(CapableTextField.CONVERSION_TYPE.NUMERIC)
                .atTop(jLabel3).withWidth(70).atRight(txtDisplayName.getRight()).to(getContentPane());
        JLabel jLabel4 = SettingsLabel.create("No. of Shares Bought").atTop(jLabel3).atRight(txtSharesBought.getX() - lblGap).to(getContentPane());

        // Currency
        JLabel jLabel5 = SettingsLabel.create("Currency Code").below(jLabel3, vGap).to(getContentPane());
        txtCurrencyCode = CapableTextField.create("", "e.g. GBP, USD").setConversionType(CapableTextField.CONVERSION_TYPE.UPPER)
                .below(txtPricePaid, vGap).withWidth(70).to(getContentPane());

        JLabel jLabel6 = SettingsLabel.create("Currency Symbol").below(jLabel4, vGap).to(getContentPane());
        txtCurrencySymbol = SettingsTextField.create("", "e.g. $, £, p, c").below(txtSharesBought, vGap).withWidth(50).to(getContentPane());

        // "Show" Panel (manual border and grouping)
        SettingsPanel jPanel1 = SettingsPanel.create(null).withBorder(BorderFactory.createTitledBorder(lineBorder, "Show"))
                .below(jLabel5, vGap).withDimensions(txtDisplayName.getRight() - jLabel5.getX(), 300).to(getContentPane());

        // Show-panel checkboxes
        // Left column
        chkShowPrice = SettingsCheckbox.create("Price").at(30, 20).withWidth(140).to(jPanel1);
        chkShowChange = SettingsCheckbox.create("Change").below(chkShowPrice, vGap / 3).to(jPanel1);
        chlShowUpDown = SettingsCheckbox.create("Up/Down").below(chkShowChange, vGap / 3).to(jPanel1);

        // Right column
        chkExcludeFromSummary = SettingsCheckbox.create("Hide from Summary").tail(chkShowPrice, 30).to(jPanel1);
        chkShowChangePercent = SettingsCheckbox.create("Change %").below(chkExcludeFromSummary, vGap / 3).to(jPanel1);
        chkShowProfitLoss = SettingsCheckbox.create("Profit & Loss").below(chkShowChangePercent, vGap / 3).to(jPanel1);

        // Day Change
        // Left column
        chkShowDayChange = SettingsCheckbox.create("Day Change").below(chlShowUpDown, vGap).to(jPanel1).to(jPanel1);
        chkShowDayUpDown = SettingsCheckbox.create("Day Up/Down").below(chkShowDayChange, vGap / 3).to(jPanel1);
        chkShowDayChangePercent = SettingsCheckbox.create("Day Change %").below(chkShowProfitLoss, vGap).to(jPanel1);

        // Resize the Show panel to fit the checkboxes
        jPanel1.withHeight(chkShowDayUpDown.getBottom() + vGap);

        // Low Alarm Panel
        pnlAlarmLow = CheckBoxFrame.create("Enable Low Alarm").withLayout(null).below(jPanel1, vGap).to(getContentPane());
        JLabel jLabel7 = SettingsLabel.create("Prices Drops to").at(20, 15).to(pnlAlarmLow.getContentPanel());
        txtAlarmLow = CapableTextField.create("", "Threshold for low alarm").setConversionType(CapableTextField.CONVERSION_TYPE.NUMERIC)
                .tail(jLabel7, lblGap).withWidth(70).to(pnlAlarmLow.getContentPanel());
        chkAlarmLowPercent = SettingsCheckbox.create("Percent").tail(txtAlarmLow, hGap).to(pnlAlarmLow.getContentPanel());
        chkAlarmLowPlaySound = SettingsCheckbox.create("Sound Alarm").below(chkAlarmLowPercent, vGap / 3).to(pnlAlarmLow.getContentPanel());
        pnlAlarmLow.withHeight(chkAlarmLowPlaySound.getBottom() + vGap * 2);

        // High Alarm Panel
        pnlAlarmHigh = CheckBoxFrame.create("Enable High Alarm").withLayout(null).below(pnlAlarmLow, vGap).to(getContentPane());
        JLabel jLabel8 = SettingsLabel.create("Prices Rises to").at(20, 15).to(pnlAlarmHigh.getContentPanel());
        txtAlarmHigh = CapableTextField.create("", "Threshold for high alarm").setConversionType(CapableTextField.CONVERSION_TYPE.NUMERIC)
                .tail(jLabel8, lblGap).withWidth(70).to(pnlAlarmHigh.getContentPanel());
        chkAlarmHighPercent = SettingsCheckbox.create("Percent").tail(txtAlarmHigh, hGap).to(pnlAlarmHigh.getContentPanel());
        chkAlarmHighPlaySound = SettingsCheckbox.create("Sound Alarm").below(chkAlarmHighPercent, vGap / 3).to(pnlAlarmHigh.getContentPanel());
        pnlAlarmHigh.withHeight(chkAlarmHighPlaySound.getBottom() + vGap * 2);

        // Buttons
        jScrollPane1.withHeight(pnlAlarmHigh.getBottom() - jScrollPane1.getY());
        btnAdd = SettingsButton.create("Add").below(jScrollPane1, vGap).withDimensions(70, 15).to(getContentPane());
        btnDelete = SettingsButton.create("Delete").tail(btnAdd, 0).withDimensions(btnAdd).atRight(jScrollPane1.getRight()).to(getContentPane());
        btnCancel = SettingsButton.create("Cancel").atTop(btnAdd).withWidth(80).atRight(txtDisplayName.getRight()).to(getContentPane());
        btnOk = SettingsButton.create("Ok").atTop(btnCancel).atLeft(btnCancel.getX() - btnCancel.getWidth() - hGap).withDimensions(btnCancel).to(getContentPane());

        setSize(width, btnOk.getBottom() + 40);
        setPreferredSize(getSize());
        setMaximumSize(getSize());
        setMinimumSize(getSize());

        btnOk.setEnabled(false);
    }

}
