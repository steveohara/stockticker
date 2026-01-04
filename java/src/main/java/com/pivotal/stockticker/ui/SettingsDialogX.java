package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.model.Settings;
import com.pivotal.stockticker.service.PreferencesService;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;

@Slf4j
public class SettingsDialogX extends JDialog {
    @Getter
    private final Settings settings;
    private final PreferencesService prefsService;
    private JTextField proxyField, currencyCodeField, currencySymbolField;
    private JSpinner frequencySpinner;
    private JCheckBox alwaysOnTopCheck, showTotalCheck, showTotalPercentCheck;
    private JCheckBox showTotalCostCheck, showTotalValueCheck, showDailyChangeCheck, showPriceCheck, showCostBaseCheck;
    private JButton upColorButton, downColorButton, normalTextButton, upArrowButton, downArrowButton, fontButton;
    private JLabel fontPreviewLabel;

    public SettingsDialogX(Frame parent, Settings settings, PreferencesService prefsService) {
        super(parent, "Settings", true);
        this.settings = settings;
        this.prefsService = prefsService;
        initializeUI();
        loadSettings();
    }

    private void initializeUI() {
        setLayout(new BorderLayout(10, 10));
        setSize(700, 600);
        setLocationRelativeTo(getParent());
        JTabbedPane tabbedPane = new JTabbedPane();
        tabbedPane.addTab("General", createGeneralPanel());
        tabbedPane.addTab("Display", createDisplayPanel());
        tabbedPane.addTab("Colors & Fonts", createColorsPanel());
        add(tabbedPane, BorderLayout.CENTER);
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        JButton okButton = new JButton("OK");
        okButton.addActionListener(e -> {
            saveSettings();
            dispose();
        });
        buttonPanel.add(okButton);
        JButton cancelButton = new JButton("Cancel");
        cancelButton.addActionListener(e -> dispose());
        buttonPanel.add(cancelButton);
        add(buttonPanel, BorderLayout.SOUTH);
    }

    private JPanel createGeneralPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(new EmptyBorder(15, 15, 15, 15));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.anchor = GridBagConstraints.WEST;
        int row = 0;
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel("Proxy:"), gbc);
        gbc.gridx = 1;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 1.0;
        proxyField = new JTextField(30);
        panel.add(proxyField, gbc);
        row++;
        gbc.weightx = 0;
        gbc.fill = GridBagConstraints.NONE;
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel("Update Frequency (seconds):"), gbc);
        gbc.gridx = 1;
        frequencySpinner = new JSpinner(new SpinnerNumberModel(60, 1, 600, 1));
        panel.add(frequencySpinner, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel("Summary Currency Code:"), gbc);
        gbc.gridx = 1;
        currencyCodeField = new JTextField(10);
        panel.add(currencyCodeField, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel("Summary Currency Symbol:"), gbc);
        gbc.gridx = 1;
        currencySymbolField = new JTextField(10);
        panel.add(currencySymbolField, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        gbc.gridwidth = 2;
        alwaysOnTopCheck = new JCheckBox("Always on Top");
        panel.add(alwaysOnTopCheck, gbc);
        return panel;
    }

    private JPanel createDisplayPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(new EmptyBorder(15, 15, 15, 15));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.anchor = GridBagConstraints.WEST;
        gbc.gridx = 0;
        int row = 0;
        JLabel summaryLabel = new JLabel("Summary Display Options:");
        summaryLabel.setFont(summaryLabel.getFont().deriveFont(Font.BOLD));
        gbc.gridy = row++;
        panel.add(summaryLabel, gbc);
        gbc.insets = new Insets(2, 20, 2, 5);
        gbc.gridy = row++;
        showTotalCheck = new JCheckBox("Show Total Profit & Loss");
        panel.add(showTotalCheck, gbc);
        gbc.gridy = row++;
        showTotalPercentCheck = new JCheckBox("Show Total Profit & Loss as Percentage");
        panel.add(showTotalPercentCheck, gbc);
        gbc.gridy = row++;
        showTotalCostCheck = new JCheckBox("Show Total Cost");
        panel.add(showTotalCostCheck, gbc);
        gbc.gridy = row++;
        showTotalValueCheck = new JCheckBox("Show Total Value");
        panel.add(showTotalValueCheck, gbc);
        gbc.gridy = row++;
        showDailyChangeCheck = new JCheckBox("Show Daily Change");
        panel.add(showDailyChangeCheck, gbc);
        row++;
        gbc.insets = new Insets(5, 5, 5, 5);
        JLabel tickerLabel = new JLabel("Ticker Display Options:");
        tickerLabel.setFont(tickerLabel.getFont().deriveFont(Font.BOLD));
        gbc.gridy = row++;
        panel.add(tickerLabel, gbc);
        gbc.insets = new Insets(2, 20, 2, 5);
        gbc.gridy = row++;
        showPriceCheck = new JCheckBox("Show Current Price of each Stock");
        panel.add(showPriceCheck, gbc);
        gbc.gridy = row++;
        showCostBaseCheck = new JCheckBox("Show Average Cost of each Stock (Cost Base)");
        panel.add(showCostBaseCheck, gbc);
        return panel;
    }

    private JPanel createColorsPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(new EmptyBorder(15, 15, 15, 15));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.anchor = GridBagConstraints.WEST;
        int row = 0;
        JLabel colorsLabel = new JLabel("Colors:");
        colorsLabel.setFont(colorsLabel.getFont().deriveFont(Font.BOLD));
        gbc.gridx = 0;
        gbc.gridy = row++;
        gbc.gridwidth = 2;
        panel.add(colorsLabel, gbc);
        gbc.gridwidth = 1;
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel("Up Color:"), gbc);
        gbc.gridx = 1;
        upColorButton = createColorButton(settings.getUpColor());
        panel.add(upColorButton, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel("Down Color:"), gbc);
        gbc.gridx = 1;
        downColorButton = createColorButton(settings.getDownColor());
        panel.add(downColorButton, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel("Normal Text:"), gbc);
        gbc.gridx = 1;
        normalTextButton = createColorButton(settings.getNormalTextColor());
        panel.add(normalTextButton, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel("Up Arrow Color:"), gbc);
        gbc.gridx = 1;
        upArrowButton = createColorButton(settings.getUpArrowColor());
        panel.add(upArrowButton, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel("Down Arrow Color:"), gbc);
        gbc.gridx = 1;
        downArrowButton = createColorButton(settings.getDownArrowColor());
        panel.add(downArrowButton, gbc);
        row += 2;
        JLabel fontLabel = new JLabel("Font:");
        fontLabel.setFont(fontLabel.getFont().deriveFont(Font.BOLD));
        gbc.gridx = 0;
        gbc.gridy = row++;
        gbc.gridwidth = 2;
        panel.add(fontLabel, gbc);
        gbc.gridwidth = 1;
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel("Ticker Font:"), gbc);
        gbc.gridx = 1;
        fontButton = new JButton("Choose Font...");
        fontButton.addActionListener(e -> chooseFont());
        panel.add(fontButton, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        gbc.gridwidth = 2;
        fontPreviewLabel = new JLabel("Preview: AaBbCc 123");
        fontPreviewLabel.setFont(settings.getTickerFont());
        panel.add(fontPreviewLabel, gbc);
        return panel;
    }

    private JButton createColorButton(Color color) {
        JButton button = new JButton("   ");
        button.setBackground(color);
        button.setOpaque(true);
        button.addActionListener(e -> {
            Color newColor = JColorChooser.showDialog(this, "Choose Color", color);
            if (newColor != null) {
                button.setBackground(newColor);
            }
        });
        return button;
    }

    private void chooseFont() {
        String[] fonts = GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames();
        String fontName = (String) JOptionPane.showInputDialog(this, "Choose Font:", "Font Selection", JOptionPane.QUESTION_MESSAGE, null, fonts, settings.getTickerFont().getName());
        if (fontName != null) {
            Integer[] sizes = {8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32};
            Integer size = (Integer) JOptionPane.showInputDialog(this, "Choose Size:", "Font Size", JOptionPane.QUESTION_MESSAGE, null, sizes, settings.getTickerFont().getSize());
            if (size != null) {
                Font newFont = new Font(fontName, Font.PLAIN, size);
                fontPreviewLabel.setFont(newFont);
            }
        }
    }

    private void loadSettings() {
        proxyField.setText(settings.getProxy());
        frequencySpinner.setValue(settings.getFrequency());
        currencyCodeField.setText(settings.getSummaryCurrency());
        currencySymbolField.setText(settings.getSummaryCurrencySymbol());
        alwaysOnTopCheck.setSelected(settings.isAlwaysOnTop());
        showTotalCheck.setSelected(settings.isShowTotal());
        showTotalPercentCheck.setSelected(settings.isShowTotalPercent());
        showTotalCostCheck.setSelected(settings.isShowTotalCost());
        showTotalValueCheck.setSelected(settings.isShowTotalValue());
        showDailyChangeCheck.setSelected(settings.isShowDailyChange());
        showPriceCheck.setSelected(settings.isShowPrice());
        showCostBaseCheck.setSelected(settings.isShowCostBase());
    }

    private void saveSettings() {
        settings.setProxy(proxyField.getText());
        settings.setUpdateFrequency((Integer) frequencySpinner.getValue());
        settings.setSummaryCurrency(currencyCodeField.getText());
        settings.setSummaryCurrencySymbol(currencySymbolField.getText());
        settings.setAlwaysOnTop(alwaysOnTopCheck.isSelected());
        settings.setShowTotal(showTotalCheck.isSelected());
        settings.setShowTotalPercent(showTotalPercentCheck.isSelected());
        settings.setShowTotalCost(showTotalCostCheck.isSelected());
        settings.setShowTotalValue(showTotalValueCheck.isSelected());
        settings.setShowDailyChange(showDailyChangeCheck.isSelected());
        settings.setShowPrice(showPriceCheck.isSelected());
        settings.setShowCostBase(showCostBaseCheck.isSelected());
        settings.setUpColor(upColorButton.getBackground());
        settings.setDownColor(downColorButton.getBackground());
        settings.setNormalTextColor(normalTextButton.getBackground());
        settings.setUpArrowColor(upArrowButton.getBackground());
        settings.setDownArrowColor(downArrowButton.getBackground());
        settings.setTickerFont(fontPreviewLabel.getFont());
        prefsService.saveSettings(settings);
    }

}
