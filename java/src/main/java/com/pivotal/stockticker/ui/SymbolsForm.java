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
import com.pivotal.stockticker.ui.components.CapableTextField;
import com.pivotal.stockticker.ui.components.CheckBoxFrame;
import com.pivotal.stockticker.ui.components.SymbolsList;
import com.pivotal.stockticker.utils.CallbackInterface;
import com.pivotal.stockticker.utils.StartupManager;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import javax.swing.border.LineBorder;
import javax.swing.border.TitledBorder;
import javax.swing.event.ListSelectionEvent;
import java.awt.*;
import java.awt.event.*;

/**
 * Form to manage stock symbols
 */
@Slf4j
public class SymbolsForm extends JDialog implements CallbackInterface {

    private JButton btnAdd, btnCancel, btnDelete, btnOk;
    private JCheckBox chkHideDisabled, chkAlarmHighPercent, chkAlarmHighPlaySound, chkAlarmLowPercent, chkAlarmLowPlaySound, chkDisabled, chkExcludeFromSummary, chkShowChange, chkShowChangePercent, chkShowDayChange, chkShowDayChangePercent, chkShowDayUpDown, chkShowPrice, chkShowProfitLoss, chlShowUpDown;
    private CheckBoxFrame pnlAlarmLow, pnlAlarmHigh;
    private SymbolsList lstSymbols;
    private CapableTextField txtAlarmHigh, txtAlarmLow, txtCurrencyCode, txtPricePaid, txtSharesBought, txtSymbol;
    private JTextField txtCurrencySymbol, txtDisplayName;
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
        int vGap = 10;
        int hGap = 10;
        int lblGap = 5;
        int stdWidth = 100;
        int stdHeight = 20;
        int width = StartupManager.isWindows() ? 610 :  600;

        // Set the dialog size and use null layout
        setResizable(false);
        getContentPane().setLayout(null);

        LineBorder lineBorder = new LineBorder(UIManager.getColor("Component.borderColor"), 1);

        chkHideDisabled = new JCheckBox("Hide Disabled Symbols");
        chkHideDisabled.setBounds(hGap, vGap, 200, stdHeight);
        chkHideDisabled.setFont(new Font(chkHideDisabled.getFont().getName(), Font.PLAIN, 10));
        chkHideDisabled.setSelected(settings.isHideDisabledSymbols());
        getContentPane().add(chkHideDisabled);

        // List of Symbols
        JScrollPane jScrollPane1 = new JScrollPane();
        lstSymbols = new SymbolsList();
        lstSymbols.setBorder(null);
        lstSymbols.hideDisabled(settings.isHideDisabledSymbols());
        jScrollPane1.setBorder(lineBorder);
        jScrollPane1.setViewportView(lstSymbols);
        jScrollPane1.setBounds(hGap, chkHideDisabled.getX() + chkHideDisabled.getHeight() + vGap / 2, 150, 400);
        getContentPane().add(jScrollPane1);

        int left = jScrollPane1.getX() + jScrollPane1.getWidth() + (int)(hGap * 1.5);

        // Symbol & Details Labels/Fields
        JLabel jLabel1 = new JLabel("Symbol");
        jLabel1.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel1.setBounds(left, hGap, stdWidth, stdHeight);
        getContentPane().add(jLabel1);

        txtSymbol = new CapableTextField(CapableTextField.CONVERSION_TYPE.UPPER);
        txtSymbol.setBounds(jLabel1.getX() + jLabel1.getWidth() + lblGap, jLabel1.getY(), 100, stdHeight);
        getContentPane().add(txtSymbol);

        chkDisabled = new JCheckBox("Disabled");
        chkDisabled.setToolTipText("Do not use this symbol");
        chkDisabled.setBounds(width - stdWidth - hGap, txtSymbol.getY(), stdWidth, stdHeight);
        getContentPane().add(chkDisabled);

        lblTransactionTimestamp = new JLabel();
        lblTransactionTimestamp.setHorizontalAlignment(SwingConstants.CENTER);
        lblTransactionTimestamp.setBounds(txtSymbol.getX() + txtSymbol.getWidth(), txtSymbol.getY(), chkDisabled.getX() - txtSymbol.getX() - txtSymbol.getWidth(), stdHeight);
        getContentPane().add(lblTransactionTimestamp);

        // Display Name
        JLabel jLabel2 = new JLabel("Display Name");
        jLabel2.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel2.setBounds(left, jLabel1.getY() + jLabel1.getHeight() + vGap, stdWidth, stdHeight);
        getContentPane().add(jLabel2);

        txtDisplayName = new JTextField();
        txtDisplayName.setToolTipText("Name you want to appear instead of symbol");
        txtDisplayName.setBounds(jLabel2.getX() + jLabel2.getWidth() + lblGap, jLabel2.getY(), 300, stdHeight);
        getContentPane().add(txtDisplayName);

        // Price Paid & Shares Bought
        JLabel jLabel3 = new JLabel("Price Paid");
        jLabel3.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel3.setBounds(left, jLabel2.getY() + jLabel2.getHeight() + vGap, stdWidth, stdHeight);
        getContentPane().add(jLabel3);

        txtPricePaid = new CapableTextField(CapableTextField.CONVERSION_TYPE.NUMERIC);
        txtPricePaid.setBounds(jLabel3.getX() + jLabel3.getWidth() + lblGap, jLabel3.getY(), 70, stdHeight);
        getContentPane().add(txtPricePaid);

        txtSharesBought = new CapableTextField(CapableTextField.CONVERSION_TYPE.NUMERIC);
        txtSharesBought.setBounds(txtDisplayName.getX() + txtDisplayName.getWidth() - 70, txtPricePaid.getY(), 70, stdHeight);
        getContentPane().add(txtSharesBought);

        JLabel jLabel4 = new JLabel("No. of Shares Bought");
        jLabel4.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel4.setBounds(txtSharesBought.getX() - 146 - lblGap, txtSharesBought.getY(), 146, stdHeight);
        getContentPane().add(jLabel4);

        // Currency
        JLabel jLabel5 = new JLabel("Currency Code");
        jLabel5.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel5.setBounds(left, jLabel4.getY() + jLabel4.getHeight() + vGap, stdWidth, stdHeight);
        getContentPane().add(jLabel5);

        txtCurrencyCode = new CapableTextField(CapableTextField.CONVERSION_TYPE.UPPER);
        txtCurrencyCode.setToolTipText("e.g. GBP, USD");
        txtCurrencyCode.setBounds(jLabel5.getX() + jLabel5.getWidth() + lblGap, jLabel5.getY(), 70, stdHeight);
        getContentPane().add(txtCurrencyCode);

        JLabel jLabel6 = new JLabel("Currency Symbol");
        jLabel6.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel6.setBounds(jLabel4.getX(), jLabel4.getY() + jLabel4.getHeight() + vGap, jLabel4.getWidth(), stdHeight);
        getContentPane().add(jLabel6);

        txtCurrencySymbol = new JTextField();
        txtCurrencySymbol.setToolTipText("e.g. $, £, p, c");
        txtCurrencySymbol.setBounds(txtSharesBought.getX(), jLabel6.getY(), txtSharesBought.getWidth(), stdHeight);
        getContentPane().add(txtCurrencySymbol);

        // "Show" Panel (manual border and grouping)
        JPanel jPanel1 = new JPanel(null);
        TitledBorder border = BorderFactory.createTitledBorder(lineBorder, "Show");
        border.setTitleColor(lineBorder.getLineColor());
        jPanel1.setBorder(border);
        jPanel1.setBounds(left, txtCurrencyCode.getY() + txtCurrencyCode.getHeight() + vGap, txtCurrencySymbol.getX() + txtCurrencySymbol.getWidth() - left, 300);
        getContentPane().add(jPanel1);

        // Show-panel checkboxes
        // Left column
        chkShowPrice = new JCheckBox("Price");
        chkShowPrice.setBounds(30, 20, stdWidth, stdHeight);
        jPanel1.add(chkShowPrice);

        chkShowChange = new JCheckBox("Change");
        chkShowChange.setBounds(chkShowPrice.getX(), chkShowPrice.getY() + chkShowPrice.getHeight() + vGap / 2, stdWidth, stdHeight);
        jPanel1.add(chkShowChange);

        chlShowUpDown = new JCheckBox("Up/Down");
        chlShowUpDown.setBounds(chkShowChange.getX(), chkShowChange.getY() + chkShowChange.getHeight() + vGap / 2, stdWidth, stdHeight);
        jPanel1.add(chlShowUpDown);

        // Right column
        chkExcludeFromSummary = new JCheckBox("Hide from Summary");
        chkExcludeFromSummary.setBounds(200, chkShowPrice.getY(), stdWidth * 2, stdHeight);
        jPanel1.add(chkExcludeFromSummary);

        chkShowChangePercent = new JCheckBox("Change %");
        chkShowChangePercent.setBounds(chkExcludeFromSummary.getX(), chkShowChange.getY(), stdWidth, stdHeight);
        jPanel1.add(chkShowChangePercent);

        chkShowProfitLoss = new JCheckBox("Profit & Loss");
        chkShowProfitLoss.setBounds(chkShowChangePercent.getX(), chlShowUpDown.getY(), stdWidth, stdHeight);
        jPanel1.add(chkShowProfitLoss);

        // Day Change
        // Left column
        chkShowDayChange = new JCheckBox("Day Change");
        chkShowDayChange.setBounds(chkShowPrice.getX(), chlShowUpDown.getY() + chlShowUpDown.getHeight() + vGap, stdWidth, stdHeight);
        jPanel1.add(chkShowDayChange);

        chkShowDayChangePercent = new JCheckBox("Day Change %");
        chkShowDayChangePercent.setBounds(chkShowDayChange.getX(), chkShowDayChange.getY() + chkShowDayChange.getHeight() + vGap / 2, stdWidth * 2, stdHeight);
        jPanel1.add(chkShowDayChangePercent);

        // Right column
        chkShowDayUpDown = new JCheckBox("Day Up/Down");
        chkShowDayUpDown.setBounds(chkExcludeFromSummary.getX(), chkShowDayChange.getY(), stdWidth * 2, stdHeight);
        jPanel1.add(chkShowDayUpDown);

        jPanel1.setBounds(left, jPanel1.getY(), jPanel1.getWidth(), chkShowDayChangePercent.getY() + chkShowDayChangePercent.getHeight() + vGap);

        // Low Alarm Panel
        pnlAlarmLow = new CheckBoxFrame("Enable Low Alarm");
        pnlAlarmLow.setBounds(left, jPanel1.getY() + jPanel1.getHeight() + vGap, jPanel1.getWidth(), 100);
        getContentPane().add(pnlAlarmLow);

        JLabel jLabel7 = new JLabel("Prices Drops to");
        jLabel7.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel7.setBounds(20, 20, stdWidth, stdHeight);
        pnlAlarmLow.getContentPanel().setLayout(null);
        pnlAlarmLow.getContentPanel().add(jLabel7);

        txtAlarmLow = new CapableTextField(CapableTextField.CONVERSION_TYPE.NUMERIC);
        txtAlarmLow.setToolTipText("Threshold for low alarm");
        txtAlarmLow.setBounds(jLabel7.getX() + jLabel7.getWidth() + lblGap, jLabel7.getY(), txtPricePaid.getWidth(), stdHeight);
        pnlAlarmLow.getContentPanel().add(txtAlarmLow);

        chkAlarmLowPercent = new JCheckBox("Percent");
        chkAlarmLowPercent.setBounds(250, txtAlarmLow.getY(), stdWidth, stdHeight);
        pnlAlarmLow.getContentPanel().add(chkAlarmLowPercent);

        chkAlarmLowPlaySound = new JCheckBox("Sound Alarm");
        chkAlarmLowPlaySound.setBounds(chkAlarmLowPercent.getX(), chkAlarmLowPercent.getY() + chkAlarmLowPercent.getHeight() + vGap / 2, stdWidth, stdHeight);
        pnlAlarmLow.getContentPanel().add(chkAlarmLowPlaySound);

        pnlAlarmLow.setBounds(left, pnlAlarmLow.getY(), pnlAlarmLow.getWidth(), chkAlarmLowPlaySound.getY() + chkAlarmLowPlaySound.getHeight() + vGap * 2);

        // High Alarm Panel
        pnlAlarmHigh = new CheckBoxFrame("Enable High Alarm");
        pnlAlarmHigh.setBounds(left, pnlAlarmLow.getY() + pnlAlarmLow.getHeight() + vGap, pnlAlarmLow.getWidth(), 100);
        getContentPane().add(pnlAlarmHigh);

        JLabel jLabel8 = new JLabel("Prices Drops to");
        jLabel8.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel8.setBounds(20, 20, stdWidth, stdHeight);
        pnlAlarmHigh.getContentPanel().setLayout(null);
        pnlAlarmHigh.getContentPanel().add(jLabel8);

        txtAlarmHigh = new CapableTextField(CapableTextField.CONVERSION_TYPE.NUMERIC);
        txtAlarmHigh.setToolTipText("Threshold for high alarm");
        txtAlarmHigh.setBounds(jLabel8.getX() + jLabel8.getWidth() + lblGap, jLabel8.getY(), txtPricePaid.getWidth(), stdHeight);
        pnlAlarmHigh.getContentPanel().add(txtAlarmHigh);

        chkAlarmHighPercent = new JCheckBox("Percent");
        chkAlarmHighPercent.setBounds(250, txtAlarmHigh.getY(), stdWidth, stdHeight);
        pnlAlarmHigh.getContentPanel().add(chkAlarmHighPercent);

        chkAlarmHighPlaySound = new JCheckBox("Sound Alarm");
        chkAlarmHighPlaySound.setBounds(chkAlarmHighPercent.getX(), chkAlarmHighPercent.getY() + chkAlarmHighPercent.getHeight() + vGap / 2, stdWidth, stdHeight);
        pnlAlarmHigh.getContentPanel().add(chkAlarmHighPlaySound);

        pnlAlarmHigh.setBounds(left, pnlAlarmHigh.getY(), pnlAlarmHigh.getWidth(), chkAlarmHighPlaySound.getY() + chkAlarmHighPlaySound.getHeight() + vGap * 2);

        // OK and Cancel Buttons
        btnCancel = new JButton("Cancel");
        btnCancel.setBounds(pnlAlarmHigh.getX() + pnlAlarmHigh.getWidth() - 75, pnlAlarmHigh.getY() + pnlAlarmHigh.getHeight() + vGap, 75, 30);
        getContentPane().add(btnCancel);

        btnOk = new JButton("OK");
        btnOk.setEnabled(false);
        btnOk.setBounds(btnCancel.getX() - btnCancel.getWidth() - vGap, btnCancel.getY(), btnCancel.getWidth(), btnCancel.getHeight());
        getContentPane().add(btnOk);

        // Add & Delete Buttons
        jScrollPane1.setBounds(jScrollPane1.getX(), jScrollPane1.getY(), jScrollPane1.getWidth(), pnlAlarmHigh.getY() + pnlAlarmHigh.getHeight() - jScrollPane1.getY());
        btnAdd = new JButton("Add");
        btnAdd.setBounds(jScrollPane1.getX(), jScrollPane1.getY() + jScrollPane1.getHeight() + vGap, 70, 20);
        getContentPane().add(btnAdd);

        btnDelete = new JButton("Delete");
        btnDelete.setBounds(jScrollPane1.getX() + jScrollPane1.getWidth() - 70, btnAdd.getY(), btnAdd.getWidth(), btnAdd.getHeight());
        getContentPane().add(btnDelete);

        // Size the dialog
        setSize(width, btnCancel.getY() + btnCancel.getHeight() + (StartupManager.isWindows() ? 50 :  40));
        setPreferredSize(getSize());
        setMinimumSize(getSize());
        setMaximumSize(getSize());
    }

}
