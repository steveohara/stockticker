package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.model.Symbol;
import com.pivotal.stockticker.service.PreferencesService;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.util.List;

/**
 * Dialog for managing stock symbols and their settings.
 */
@Slf4j
public class SymbolsDialogX extends JDialog {
    @Getter
    private final List<Symbol> symbols;
    private final PreferencesService prefsService;
    private DefaultListModel<Symbol> listModel;
    private JList<Symbol> symbolList;
    private JTextField codeField, aliasField, priceField, sharesField, currencyCodeField, currencySymbolField;
    private JCheckBox showPriceCheck, showChangeCheck, showChangePercentCheck, showUpDownCheck, showProfitLossCheck;
    private JCheckBox showDayChangeCheck, showDayChangePercentCheck, showDayUpDownCheck, excludeFromSummaryCheck, disabledCheck;
    private JCheckBox highAlarmEnabledCheck, highAlarmPercentCheck, highAlarmSoundCheck;
    private JTextField highAlarmValueField;
    private JCheckBox lowAlarmEnabledCheck, lowAlarmPercentCheck, lowAlarmSoundCheck;
    private JTextField lowAlarmValueField;
    private Symbol currentSymbol;

    /**
     * Constructs the SymbolsDialog.
     *
     * @param parent      Parent frame
     * @param symbols     List of symbols to manage
     * @param prefsService Preferences service for saving/loading symbols
     */
    public SymbolsDialogX(Frame parent, List<Symbol> symbols, PreferencesService prefsService) {
        super(parent, "Symbols", true);
        this.symbols = symbols;
        this.prefsService = prefsService;
        initializeUI();
        loadSymbols();
    }

    /**
     * Initializes the user interface components.
     */
    private void initializeUI() {
        setLayout(new BorderLayout(10, 10));
        setSize(850, 650);
        setLocationRelativeTo(getParent());
        add(createSymbolListPanel(), BorderLayout.WEST);
        add(createDetailsPanel(), BorderLayout.CENTER);
        add(createButtonPanel(), BorderLayout.SOUTH);
    }

    /**
     * Creates the panel containing the list of symbols.
     *
     * @return JPanel containing the symbol list
     */
    private JPanel createSymbolListPanel() {
        JPanel panel = new JPanel(new BorderLayout(5, 5));
        panel.setBorder(new EmptyBorder(10, 10, 10, 5));
        panel.setPreferredSize(new Dimension(250, 0));
        listModel = new DefaultListModel<>();
        symbolList = new JList<>(listModel);
        symbolList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        symbolList.setCellRenderer(new DefaultListCellRenderer() {
            @Override
            public Component getListCellRendererComponent(JList<?> list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
                super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
                if (value instanceof Symbol) {
                    Symbol symbol = (Symbol) value;
                    setText(symbol.getDisplayName());
                    if (symbol.isDisabled()) {
                        setForeground(Color.GRAY);
                    }
                }
                return this;
            }
        });
        symbolList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                loadSymbolDetails();
            }
        });
        JScrollPane scrollPane = new JScrollPane(symbolList);
        panel.add(scrollPane, BorderLayout.CENTER);
        JPanel listButtonPanel = new JPanel(new GridLayout(2, 1, 5, 5));
        JButton addButton = new JButton("Add");
        addButton.addActionListener(e -> addSymbol());
        listButtonPanel.add(addButton);
        JButton deleteButton = new JButton("Delete");
        deleteButton.addActionListener(e -> deleteSymbol());
        listButtonPanel.add(deleteButton);
        panel.add(listButtonPanel, BorderLayout.SOUTH);
        return panel;
    }

    /**
     * Creates the panel containing the symbol details with tabs.
     *
     * @return JPanel containing symbol details
     */
    private JPanel createDetailsPanel() {
        JPanel panel = new JPanel(new BorderLayout(5, 5));
        panel.setBorder(new EmptyBorder(10, 5, 10, 10));
        JTabbedPane tabbedPane = new JTabbedPane();
        tabbedPane.addTab("Basic", createBasicPanel());
        tabbedPane.addTab("Display", createDisplayPanel());
        tabbedPane.addTab("Alarms", createAlarmsPanel());
        panel.add(tabbedPane, BorderLayout.CENTER);
        return panel;
    }

    /**
     * Creates the "Basic" tab panel.
     *
     * @return JPanel for the Basic tab
     */
    private JPanel createBasicPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.anchor = GridBagConstraints.WEST;
        int row = 0;
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel("Symbol:"), gbc);
        gbc.gridx = 1;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 1.0;
        codeField = new JTextField(20);
        panel.add(codeField, gbc);
        row++;
        gbc.weightx = 0;
        gbc.fill = GridBagConstraints.NONE;
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel("Display Name:"), gbc);
        gbc.gridx = 1;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 1.0;
        aliasField = new JTextField(20);
        panel.add(aliasField, gbc);
        row++;
        gbc.weightx = 0;
        gbc.fill = GridBagConstraints.NONE;
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel("Price Paid:"), gbc);
        gbc.gridx = 1;
        priceField = new JTextField(10);
        panel.add(priceField, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel("Number of Shares:"), gbc);
        gbc.gridx = 1;
        sharesField = new JTextField(10);
        panel.add(sharesField, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel("Currency Code:"), gbc);
        gbc.gridx = 1;
        currencyCodeField = new JTextField(10);
        panel.add(currencyCodeField, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel("Currency Symbol:"), gbc);
        gbc.gridx = 1;
        currencySymbolField = new JTextField(10);
        panel.add(currencySymbolField, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        gbc.gridwidth = 2;
        disabledCheck = new JCheckBox("Disabled");
        panel.add(disabledCheck, gbc);
        return panel;
    }

    /**
     * Creates the "Display" tab panel.
     *
     * @return JPanel for the Display tab
     */
    private JPanel createDisplayPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.anchor = GridBagConstraints.WEST;
        gbc.gridx = 0;
        int row = 0;
        JPanel showPanel = new JPanel(new GridLayout(0, 2, 10, 5));
        showPanel.setBorder(BorderFactory.createTitledBorder("Show"));
        showPriceCheck = new JCheckBox("Price");
        showPanel.add(showPriceCheck);
        showChangeCheck = new JCheckBox("Change");
        showPanel.add(showChangeCheck);
        showChangePercentCheck = new JCheckBox("Change %");
        showPanel.add(showChangePercentCheck);
        showUpDownCheck = new JCheckBox("Up/Down");
        showPanel.add(showUpDownCheck);
        showProfitLossCheck = new JCheckBox("Profit & Loss");
        showPanel.add(showProfitLossCheck);
        showDayChangeCheck = new JCheckBox("Day Change");
        showPanel.add(showDayChangeCheck);
        showDayChangePercentCheck = new JCheckBox("Day Change %");
        showPanel.add(showDayChangePercentCheck);
        showDayUpDownCheck = new JCheckBox("Day Up/Down");
        showPanel.add(showDayUpDownCheck);
        gbc.gridy = row++;
        gbc.gridwidth = 2;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        panel.add(showPanel, gbc);
        gbc.gridy = row++;
        excludeFromSummaryCheck = new JCheckBox("Exclude from Summary");
        panel.add(excludeFromSummaryCheck, gbc);
        return panel;
    }

    /**
     * Creates the "Alarms" tab panel.
     *
     * @return JPanel for the Alarms tab
     */
    private JPanel createAlarmsPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.anchor = GridBagConstraints.WEST;
        JPanel highAlarmPanel = new JPanel(new GridBagLayout());
        highAlarmPanel.setBorder(BorderFactory.createTitledBorder("High Alarm"));
        GridBagConstraints hgbc = new GridBagConstraints();
        hgbc.insets = new Insets(5, 5, 5, 5);
        hgbc.anchor = GridBagConstraints.WEST;
        hgbc.gridx = 0;
        hgbc.gridy = 0;
        hgbc.gridwidth = 2;
        highAlarmEnabledCheck = new JCheckBox("Enable High Alarm");
        highAlarmPanel.add(highAlarmEnabledCheck, hgbc);
        hgbc.gridy = 1;
        hgbc.gridwidth = 1;
        highAlarmPanel.add(new JLabel("Value:"), hgbc);
        hgbc.gridx = 1;
        highAlarmValueField = new JTextField(10);
        highAlarmPanel.add(highAlarmValueField, hgbc);
        hgbc.gridx = 0;
        hgbc.gridy = 2;
        highAlarmPercentCheck = new JCheckBox("Percent");
        highAlarmPanel.add(highAlarmPercentCheck, hgbc);
        hgbc.gridx = 1;
        highAlarmSoundCheck = new JCheckBox("Sound Alarm");
        highAlarmPanel.add(highAlarmSoundCheck, hgbc);
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        panel.add(highAlarmPanel, gbc);
        JPanel lowAlarmPanel = new JPanel(new GridBagLayout());
        lowAlarmPanel.setBorder(BorderFactory.createTitledBorder("Low Alarm"));
        GridBagConstraints lgbc = new GridBagConstraints();
        lgbc.insets = new Insets(5, 5, 5, 5);
        lgbc.anchor = GridBagConstraints.WEST;
        lgbc.gridx = 0;
        lgbc.gridy = 0;
        lgbc.gridwidth = 2;
        lowAlarmEnabledCheck = new JCheckBox("Enable Low Alarm");
        lowAlarmPanel.add(lowAlarmEnabledCheck, lgbc);
        lgbc.gridy = 1;
        lgbc.gridwidth = 1;
        lowAlarmPanel.add(new JLabel("Value:"), lgbc);
        lgbc.gridx = 1;
        lowAlarmValueField = new JTextField(10);
        lowAlarmPanel.add(lowAlarmValueField, lgbc);
        lgbc.gridx = 0;
        lgbc.gridy = 2;
        lowAlarmPercentCheck = new JCheckBox("Percent");
        lowAlarmPanel.add(lowAlarmPercentCheck, lgbc);
        lgbc.gridx = 1;
        lowAlarmSoundCheck = new JCheckBox("Sound Alarm");
        lowAlarmPanel.add(lowAlarmSoundCheck, lgbc);
        gbc.gridy = 1;
        panel.add(lowAlarmPanel, gbc);
        return panel;
    }

    /**
     * Creates the panel containing action buttons.
     *
     * @return JPanel containing action buttons
     */
    private JPanel createButtonPanel() {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        JButton saveButton = new JButton("Save");
        saveButton.addActionListener(e -> saveSymbol());
        panel.add(saveButton);
        JButton okButton = new JButton("OK");
        okButton.addActionListener(e -> {
            saveSymbol();
            dispose();
        });
        panel.add(okButton);
        JButton cancelButton = new JButton("Cancel");
        cancelButton.addActionListener(e -> dispose());
        panel.add(cancelButton);
        return panel;
    }

    /**
     * Loads symbols into the list model.
     */
    private void loadSymbols() {
        listModel.clear();
        for (Symbol symbol : symbols) {
            listModel.addElement(symbol);
        }
        if (!symbols.isEmpty()) {
            symbolList.setSelectedIndex(0);
        }
    }

    /**
     * Loads the details of the selected symbol into the form fields.
     */
    private void loadSymbolDetails() {
        currentSymbol = symbolList.getSelectedValue();
        if (currentSymbol == null) {
            return;
        }
        codeField.setText(currentSymbol.getCode());
        aliasField.setText(currentSymbol.getAlias());
        priceField.setText(String.valueOf(currentSymbol.getPrice()));
        sharesField.setText(String.valueOf(currentSymbol.getShares()));
        currencyCodeField.setText(currentSymbol.getCurrencyName());
        currencySymbolField.setText(currentSymbol.getCurrencySymbol());
        showPriceCheck.setSelected(currentSymbol.isShowPrice());
        showChangeCheck.setSelected(currentSymbol.isShowChange());
        showChangePercentCheck.setSelected(currentSymbol.isShowChangePercent());
        showUpDownCheck.setSelected(currentSymbol.isShowChangeUpDown());
        showProfitLossCheck.setSelected(currentSymbol.isShowProfitLoss());
        showDayChangeCheck.setSelected(currentSymbol.isShowDayChange());
        showDayChangePercentCheck.setSelected(currentSymbol.isShowDayChangePercent());
        showDayUpDownCheck.setSelected(currentSymbol.isShowDayChangeUpDown());
        excludeFromSummaryCheck.setSelected(currentSymbol.isExcludeFromSummary());
        disabledCheck.setSelected(currentSymbol.isDisabled());
        highAlarmEnabledCheck.setSelected(currentSymbol.isHighAlarmEnabled());
        highAlarmValueField.setText(String.valueOf(currentSymbol.getHighAlarmValue()));
        highAlarmPercentCheck.setSelected(currentSymbol.isHighAlarmIsPercent());
        highAlarmSoundCheck.setSelected(currentSymbol.isHighAlarmSoundEnabled());
        lowAlarmEnabledCheck.setSelected(currentSymbol.isLowAlarmEnabled());
        lowAlarmValueField.setText(String.valueOf(currentSymbol.getLowAlarmValue()));
        lowAlarmPercentCheck.setSelected(currentSymbol.isLowAlarmIsPercent());
        lowAlarmSoundCheck.setSelected(currentSymbol.isLowAlarmSoundEnabled());
    }

    /**
     * Saves the current symbol's details from the form fields.
     */
    private void saveSymbol() {
        if (currentSymbol == null) {
            return;
        }
        try {
            currentSymbol.setCode(codeField.getText());
            currentSymbol.setAlias(aliasField.getText());
            currentSymbol.setPrice(Double.parseDouble(priceField.getText()));
            currentSymbol.setShares(Double.parseDouble(sharesField.getText()));
            currentSymbol.setCurrencyName(currencyCodeField.getText());
            currentSymbol.setCurrencySymbol(currencySymbolField.getText());
            currentSymbol.setShowPrice(showPriceCheck.isSelected());
            currentSymbol.setShowChange(showChangeCheck.isSelected());
            currentSymbol.setShowChangePercent(showChangePercentCheck.isSelected());
            currentSymbol.setShowChangeUpDown(showUpDownCheck.isSelected());
            currentSymbol.setShowProfitLoss(showProfitLossCheck.isSelected());
            currentSymbol.setShowDayChange(showDayChangeCheck.isSelected());
            currentSymbol.setShowDayChangePercent(showDayChangePercentCheck.isSelected());
            currentSymbol.setShowDayChangeUpDown(showDayUpDownCheck.isSelected());
            currentSymbol.setExcludeFromSummary(excludeFromSummaryCheck.isSelected());
            currentSymbol.setDisabled(disabledCheck.isSelected());
            currentSymbol.setHighAlarmEnabled(highAlarmEnabledCheck.isSelected());
            currentSymbol.setHighAlarmValue(Double.parseDouble(highAlarmValueField.getText()));
            currentSymbol.setHighAlarmIsPercent(highAlarmPercentCheck.isSelected());
            currentSymbol.setHighAlarmSoundEnabled(highAlarmSoundCheck.isSelected());
            currentSymbol.setLowAlarmEnabled(lowAlarmEnabledCheck.isSelected());
            currentSymbol.setLowAlarmValue(Double.parseDouble(lowAlarmValueField.getText()));
            currentSymbol.setLowAlarmIsPercent(lowAlarmPercentCheck.isSelected());
            currentSymbol.setLowAlarmSoundEnabled(lowAlarmSoundCheck.isSelected());
            prefsService.saveSymbol(currentSymbol);
            symbolList.repaint();
        }
        catch (NumberFormatException ex) {
            JOptionPane.showMessageDialog(this, "Please enter valid numeric values", "Input Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * Adds a new symbol to the list.
     */
    private void addSymbol() {
        Symbol newSymbol = new Symbol();
        newSymbol.setCode("NEW");
        symbols.add(newSymbol);
        listModel.addElement(newSymbol);
        symbolList.setSelectedValue(newSymbol, true);
    }

    /**
     * Deletes the selected symbol from the list.
     */
    private void deleteSymbol() {
        Symbol selected = symbolList.getSelectedValue();
        if (selected != null) {
            int confirm = JOptionPane.showConfirmDialog(this, "Delete symbol " + selected.getCode() + "?", "Confirm Delete", JOptionPane.YES_NO_OPTION);
            if (confirm == JOptionPane.YES_OPTION) {
                prefsService.deleteSymbol(selected.getRegKey());
                symbols.remove(selected);
                listModel.removeElement(selected);
            }
        }
    }

}
