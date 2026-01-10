/*
 *
 * Copyright (c) 2026, 4NG and/or its affiliates. All rights reserved.
 * 4NG PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.Utils;
import com.pivotal.stockticker.model.Settings;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import java.awt.*;
import java.awt.event.*;

/**
 * Form to manage stock symbols
 */
@Slf4j
public class SymbolsForm extends JDialog implements CallbackInterface {

    private final Settings settings;
    private final CallbackInterface caller;

    /**
     * Creates new form Symbols
     */
    public SymbolsForm(CallbackInterface caller, Settings settings) {
        this.caller = caller;
        this.settings = settings;
        initComponents();
        setTitle("Symbols");
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

        // Listen for changes
        Utils.attachChangeListeners(getContentPane(), this);
    }



    /**
     * Set-up the display with data from storage
     *
     * @param settings Settings object to load data from
     */
    private void loadFromSettings(Settings settings) {

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
     * Saves the current settings from the form to the Settings object.
     */
    private void saveSettings() {

    }


    /**
     * This method is called from within the constructor to initialize the form.
     */
    private void initComponents() {

        jScrollPane1 = new JScrollPane();
        lstSymbols = new JList<>();
        btnAdd = new JButton();
        btnDelete = new JButton();
        btnCancel = new JButton();
        btnOk = new JButton();
        jLabel1 = new JLabel();
        txtSymbol = new JTextField();
        chkDisabled = new JCheckBox();
        jLabel2 = new JLabel();
        txtDisplayName = new JTextField();
        jLabel3 = new JLabel();
        txtPricePaid = new JTextField();
        jLabel4 = new JLabel();
        txtSharesBought = new JTextField();
        jLabel5 = new JLabel();
        txtCurrencyCode = new JTextField();
        jLabel6 = new JLabel();
        txtCurrencySymbol = new JTextField();
        jPanel1 = new JPanel();
        chkShowPrice = new JCheckBox();
        chkShowChange = new JCheckBox();
        chkHideFromSummary = new JCheckBox();
        chkShowChangePercent = new JCheckBox();
        chkShowProfitLoss = new JCheckBox();
        chlShowUpDown = new JCheckBox();
        chkShowDayChangePercent = new JCheckBox();
        chkShowDayChange = new JCheckBox();
        chkShowDayUpDown = new JCheckBox();
        jSeparator1 = new JSeparator();
        pnlAlarmLow = new CheckBoxFrame("Enable Low Alarm");
        chkAlarmLowPercent = new JCheckBox();
        chkAlarmLowPlaySound = new JCheckBox();
        txtAlarmLow = new JTextField();
        jLabel7 = new JLabel();
        pnlAlarmHigh = new CheckBoxFrame("Enable High Alarm");;
        chkAlarmHighPercent = new JCheckBox();
        chkAlarmHighPlaySound = new JCheckBox();
        txtAlarmHigh = new JTextField();
        jLabel9 = new JLabel();

        setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);

        lstSymbols.setBorder(BorderFactory.createLineBorder(null));
        jScrollPane1.setViewportView(lstSymbols);

        btnAdd.setText("Add");

        btnDelete.setText("Delete");

        btnCancel.setText("Cancel");

        btnOk.setText("OK");
        btnOk.setEnabled(false);

        jLabel1.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel1.setLabelFor(txtSymbol);
        jLabel1.setText("Symbol");

        chkDisabled.setText("Disabled");
        chkDisabled.setToolTipText("Do not use this symbol");

        jLabel2.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel2.setLabelFor(txtDisplayName);
        jLabel2.setText("Display Name");

        txtDisplayName.setToolTipText("The name you want to appear on the ticker for this stock instead of the symbol");

        jLabel3.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel3.setLabelFor(txtPricePaid);
        jLabel3.setText("Price Paid");

        jLabel4.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel4.setLabelFor(txtSharesBought);
        jLabel4.setText("No. of Shares Bought");

        jLabel5.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel5.setLabelFor(txtCurrencyCode);
        jLabel5.setText("Currency Code");

        txtCurrencyCode.setToolTipText("e.g. GBP, USD etc.");

        jLabel6.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel6.setLabelFor(txtCurrencySymbol);
        jLabel6.setText("Currency Symbol");

        txtCurrencySymbol.setToolTipText("e.g. $, £, p, c");

        jPanel1.setBorder(BorderFactory.createLineBorder(null));

        chkShowPrice.setText("Price");
        chkShowPrice.setToolTipText("Current price that this stock is being tradded at");

        chkShowChange.setText("Change");
        chkShowChange.setToolTipText("The change in value between the current price and the price you paid for this stock");

        chkHideFromSummary.setText("Hide from Summary");
        chkHideFromSummary.setToolTipText("Get values and display on the ticker bnut exclude from the Summary and Day Summary");

        chkShowChangePercent.setText("Change %");
        chkShowChangePercent.setToolTipText("The change in percent between the current price and the price you paid for this stock");

        chkShowProfitLoss.setText("Profit & Loss");
        chkShowProfitLoss.setToolTipText("Show the amount of money you are up or down on the stock");

        chlShowUpDown.setText("Up/Down");
        chlShowUpDown.setToolTipText("Show a symbol to indicate if the current price is higher or lower the price you bought at");

        chkShowDayChangePercent.setText("Day Change %");
        chkShowDayChangePercent.setToolTipText("Show the change in percent between the current price and days starting price for this stock");

        chkShowDayChange.setText("Day Change");
        chkShowDayChange.setToolTipText("Show the change in value between the current price and days starting price for this stock");

        chkShowDayUpDown.setText("Day Up/Down");
        chkShowDayUpDown.setToolTipText("Show show a symbol to indicate if the current price is higher or lower the days starting price");
        chkShowDayUpDown.setAlignmentY(0.0F);

        GroupLayout jPanel1Layout = new GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(10, 10, 10)
                .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.TRAILING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                            .addComponent(chkShowChange)
                            .addComponent(chkShowPrice)
                            .addComponent(chlShowUpDown))
                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                            .addComponent(chkShowChangePercent)
                            .addComponent(chkHideFromSummary)
                            .addComponent(chkShowProfitLoss))
                        .addGap(56, 56, 56))
                    .addGroup(GroupLayout.Alignment.LEADING, jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.TRAILING)
                            .addGroup(GroupLayout.Alignment.LEADING, jPanel1Layout.createSequentialGroup()
                                .addComponent(chkShowDayChange)
                                .addGap(66, 66, 66)
                                .addComponent(chkShowDayChangePercent))
                            .addComponent(chkShowDayUpDown, GroupLayout.Alignment.LEADING)
                            .addComponent(jSeparator1, GroupLayout.Alignment.LEADING, GroupLayout.PREFERRED_SIZE, 296, GroupLayout.PREFERRED_SIZE))
                        .addGap(0, 0, Short.MAX_VALUE))))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(15, 15, 15)
                .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                    .addComponent(chkShowPrice)
                    .addComponent(chkHideFromSummary))
                .addGap(3, 3, 3)
                .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                    .addComponent(chkShowChangePercent)
                    .addComponent(chkShowChange))
                .addGap(3, 3, 3)
                .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                    .addComponent(chkShowProfitLoss)
                    .addComponent(chlShowUpDown))
                .addGap(6, 6, 6)
                .addComponent(jSeparator1, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                .addGap(6, 6, 6)
                .addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                    .addComponent(chkShowDayChangePercent)
                    .addComponent(chkShowDayChange))
                .addGap(3, 3, 3)
                .addComponent(chkShowDayUpDown)
                .addContainerGap(GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        pnlAlarmLow.setEnabled(false);
        chkAlarmLowPercent.setText("Percent");
        chkAlarmLowPercent.setToolTipText("Treat the value as a percentage of the base ");

        chkAlarmLowPlaySound.setText("Sound Alarm");
        chkAlarmLowPlaySound.setToolTipText("Sound an audible alert when the alarm is triggered");

        txtAlarmLow.setToolTipText("The threshold at which a dropping price will trigger the low alarm");

        jLabel7.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel7.setLabelFor(txtAlarmLow);
        jLabel7.setText("Prices Drops to");

        GroupLayout pnlAlarmLowLayout = new GroupLayout(pnlAlarmLow.getContentPanel());
        pnlAlarmLow.getContentPanel().setLayout(pnlAlarmLowLayout);
        pnlAlarmLowLayout.setHorizontalGroup(
            pnlAlarmLowLayout.createParallelGroup(GroupLayout.Alignment.LEADING)
            .addGroup(GroupLayout.Alignment.TRAILING, pnlAlarmLowLayout.createSequentialGroup()
                .addContainerGap(20, Short.MAX_VALUE)
                .addComponent(jLabel7)
                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(txtAlarmLow, GroupLayout.PREFERRED_SIZE, 93, GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED, 21, Short.MAX_VALUE)
                .addGroup(pnlAlarmLowLayout.createParallelGroup(GroupLayout.Alignment.LEADING)
                    .addComponent(chkAlarmLowPercent)
                    .addComponent(chkAlarmLowPlaySound))
                .addGap(40, 40, 40))
        );
        pnlAlarmLowLayout.setVerticalGroup(
            pnlAlarmLowLayout.createParallelGroup(GroupLayout.Alignment.LEADING)
            .addGroup(pnlAlarmLowLayout.createSequentialGroup()
                .addGap(19, 19, 19)
                .addGroup(pnlAlarmLowLayout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                    .addComponent(chkAlarmLowPercent)
                    .addComponent(jLabel7)
                    .addComponent(txtAlarmLow, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
                .addGap(3, 3, 3)
                .addComponent(chkAlarmLowPlaySound)
                .addContainerGap(GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        pnlAlarmHigh.setEnabled(false);
        chkAlarmHighPercent.setText("Percent");
        chkAlarmHighPercent.setToolTipText("Treat the value as a percentage of the base ");

        chkAlarmHighPlaySound.setText("Sound Alarm");
        chkAlarmHighPlaySound.setToolTipText("Sound an audible alert when the alarm is triggered");

        txtAlarmHigh.setToolTipText("The threshold at which a rising price will trigger the low alarm");

        jLabel9.setHorizontalAlignment(SwingConstants.RIGHT);
        jLabel9.setLabelFor(txtAlarmHigh);
        jLabel9.setText("Prices Rises to");

        GroupLayout pnlAlarmHighLayout = new GroupLayout(pnlAlarmHigh.getContentPanel());
        pnlAlarmHigh.getContentPanel().setLayout(pnlAlarmHighLayout);
        pnlAlarmHighLayout.setHorizontalGroup(
            pnlAlarmHighLayout.createParallelGroup(GroupLayout.Alignment.LEADING)
            .addGroup(GroupLayout.Alignment.TRAILING, pnlAlarmHighLayout.createSequentialGroup()
                .addContainerGap(23, Short.MAX_VALUE)
                .addComponent(jLabel9)
                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(txtAlarmHigh, GroupLayout.PREFERRED_SIZE, 93, GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED, 23, Short.MAX_VALUE)
                .addGroup(pnlAlarmHighLayout.createParallelGroup(GroupLayout.Alignment.LEADING)
                    .addComponent(chkAlarmHighPercent)
                    .addComponent(chkAlarmHighPlaySound))
                .addGap(40, 40, 40))
        );
        pnlAlarmHighLayout.setVerticalGroup(
            pnlAlarmHighLayout.createParallelGroup(GroupLayout.Alignment.LEADING)
            .addGroup(pnlAlarmHighLayout.createSequentialGroup()
                .addGap(19, 19, 19)
                .addGroup(pnlAlarmHighLayout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                    .addComponent(chkAlarmHighPercent)
                    .addComponent(jLabel9)
                    .addComponent(txtAlarmHigh, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
                .addGap(3, 3, 3)
                .addComponent(chkAlarmHighPlaySound)
                .addContainerGap(GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        GroupLayout layout = new GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(btnAdd)
                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(btnDelete, GroupLayout.PREFERRED_SIZE, 64, GroupLayout.PREFERRED_SIZE)
                        .addGap(0, 0, Short.MAX_VALUE))
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jScrollPane1, GroupLayout.PREFERRED_SIZE, 145, GroupLayout.PREFERRED_SIZE)
                        .addGroup(layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                                    .addGroup(layout.createSequentialGroup()
                                        .addGroup(layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                                            .addGroup(layout.createSequentialGroup()
                                                .addGap(26, 26, 26)
                                                .addGroup(layout.createParallelGroup(GroupLayout.Alignment.LEADING, false)
                                                    .addComponent(jLabel2, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                                    .addComponent(jLabel1, GroupLayout.PREFERRED_SIZE, 73, GroupLayout.PREFERRED_SIZE)))
                                            .addComponent(jLabel3, GroupLayout.Alignment.TRAILING, GroupLayout.PREFERRED_SIZE, 63, GroupLayout.PREFERRED_SIZE)
                                            .addComponent(jLabel5, GroupLayout.Alignment.TRAILING))
                                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                        .addGroup(layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                                            .addGroup(layout.createSequentialGroup()
                                                .addGroup(layout.createParallelGroup(GroupLayout.Alignment.LEADING, false)
                                                    .addComponent(txtPricePaid)
                                                    .addComponent(txtCurrencyCode, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED, 33, Short.MAX_VALUE)
                                                .addGroup(layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                                                    .addGroup(GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                                                        .addComponent(jLabel4, GroupLayout.PREFERRED_SIZE, 115, GroupLayout.PREFERRED_SIZE)
                                                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                        .addComponent(txtSharesBought, GroupLayout.PREFERRED_SIZE, 49, GroupLayout.PREFERRED_SIZE))
                                                    .addGroup(GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                                                        .addComponent(jLabel6)
                                                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                        .addComponent(txtCurrencySymbol, GroupLayout.PREFERRED_SIZE, 27, GroupLayout.PREFERRED_SIZE)
                                                        .addGap(22, 22, 22))))
                                            .addGroup(layout.createSequentialGroup()
                                                .addComponent(txtSymbol, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                                .addComponent(chkDisabled))
                                            .addComponent(txtDisplayName)))
                                    .addGroup(layout.createSequentialGroup()
                                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addGroup(layout.createParallelGroup(GroupLayout.Alignment.TRAILING)
                                            .addGroup(layout.createSequentialGroup()
                                                .addComponent(btnOk)
                                                .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                                                .addComponent(btnCancel))
                                            .addGroup(layout.createParallelGroup(GroupLayout.Alignment.TRAILING, false)
                                                .addComponent(pnlAlarmLow, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                                .addComponent(jPanel1, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))))
                                .addGap(25, 25, 25))
                            .addGroup(layout.createSequentialGroup()
                                .addGap(18, 18, 18)
                                .addComponent(pnlAlarmHigh, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                                .addContainerGap(GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))))))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(GroupLayout.Alignment.TRAILING, false)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel1)
                            .addComponent(txtSymbol, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                            .addComponent(chkDisabled))
                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel2)
                            .addComponent(txtDisplayName, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel3)
                            .addComponent(txtPricePaid, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel4)
                            .addComponent(txtSharesBought, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel5)
                            .addComponent(txtCurrencyCode, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel6)
                            .addComponent(txtCurrencySymbol, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(jPanel1, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(pnlAlarmLow, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(pnlAlarmHigh, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
                    .addComponent(jScrollPane1))
                .addGroup(layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                            .addComponent(btnAdd)
                            .addComponent(btnDelete))
                        .addGap(0, 0, Short.MAX_VALUE))
                    .addGroup(GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                        .addPreferredGap(LayoutStyle.ComponentPlacement.RELATED, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addGroup(layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                            .addComponent(btnOk)
                            .addComponent(btnCancel))
                        .addGap(10, 10, 10))))
        );

        pack();
    }// </editor-fold>                        

    // Variables declaration - do not modify
    private JButton btnAdd;
    private JButton btnCancel;
    private JButton btnDelete;
    private JButton btnOk;
    private JCheckBox chkAlarmHighPercent;
    private JCheckBox chkAlarmHighPlaySound;
    private JCheckBox chkAlarmLowPercent;
    private JCheckBox chkAlarmLowPlaySound;
    private JCheckBox chkDisabled;
    private JCheckBox chkHideFromSummary;
    private JCheckBox chkShowChange;
    private JCheckBox chkShowChangePercent;
    private JCheckBox chkShowDayChange;
    private JCheckBox chkShowDayChangePercent;
    private JCheckBox chkShowDayUpDown;
    private JCheckBox chkShowPrice;
    private JCheckBox chkShowProfitLoss;
    private JCheckBox chlShowUpDown;
    private JLabel jLabel1;
    private JLabel jLabel2;
    private JLabel jLabel3;
    private JLabel jLabel4;
    private JLabel jLabel5;
    private JLabel jLabel6;
    private JLabel jLabel7;
    private JLabel jLabel9;
    private JPanel jPanel1;
    private CheckBoxFrame pnlAlarmLow;
    private CheckBoxFrame pnlAlarmHigh;
    private JScrollPane jScrollPane1;
    private JSeparator jSeparator1;
    private JList<String> lstSymbols;
    private JTextField txtAlarmHigh;
    private JTextField txtAlarmLow;
    private JTextField txtCurrencyCode;
    private JTextField txtCurrencySymbol;
    private JTextField txtDisplayName;
    private JTextField txtPricePaid;
    private JTextField txtSharesBought;
    private JTextField txtSymbol;

}
