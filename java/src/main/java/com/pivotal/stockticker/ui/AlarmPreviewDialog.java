/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.Utils;
import com.pivotal.stockticker.model.SettingsManager;
import com.pivotal.stockticker.model.SymbolTransaction;
import com.pivotal.stockticker.service.AlarmManager.AlarmType;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;

/**
 * Small popup that previews what a low/high alarm notification will look and sound like, using
 * the currently entered (possibly unsaved) settings from the Symbols dialog. Purely for
 * demonstration - nothing is saved, triggered or persisted.
 */
public class AlarmPreviewDialog extends JDialog {

    /**
     * Constructs and immediately shows the preview, playing the configured sound if enabled.
     *
     * @param owner          Owning dialog (the Symbols dialog), so this stays above it
     * @param symbol         Symbol the alarm belongs to
     * @param type           Whether this is previewing the low or high alarm
     * @param threshold      Currently entered threshold value
     * @param percent        Whether the threshold is a percentage rather than an absolute price
     * @param soundEnabled   Whether the "Sound Alarm" checkbox is currently checked
     * @param currentPrice   The symbol's current live price, or null if not yet known
     * @param currentPercent The symbol's current percentage change, or null if not yet known
     */
    public AlarmPreviewDialog(Dialog owner, SymbolTransaction symbol, AlarmType type, double threshold, boolean percent,
                               boolean soundEnabled, Double currentPrice, Double currentPercent) {
        super(owner, type == AlarmType.HIGH ? "High Alarm Preview" : "Low Alarm Preview", false);
        setResizable(false);
        setLayout(new BorderLayout());

        Color accent = type == AlarmType.HIGH ? new Color(190, 40, 40) : new Color(30, 130, 60);

        JPanel header = new JPanel(new BorderLayout());
        header.setBackground(accent);
        header.setBorder(new EmptyBorder(10, 14, 10, 14));
        JLabel title = new JLabel(type + " Alarm - " + symbol.getDisplayName());
        title.setFont(new Font("Arial", Font.BOLD, 15));
        title.setForeground(Color.WHITE);
        header.add(title, BorderLayout.WEST);
        add(header, BorderLayout.NORTH);

        JPanel body = new JPanel(new GridLayout(0, 2, 10, 6));
        body.setBorder(new EmptyBorder(14, 14, 6, 14));
        body.add(boldLabel("Current:"));
        body.add(new JLabel(currentPrice == null ? "N/A - no price data yet" : formatValue(currentPrice, currentPercent, percent, symbol)));
        body.add(boldLabel("Threshold:"));
        body.add(new JLabel(percent ? String.format("%.1f%%", threshold) : Utils.formatCurrencyValue(threshold, symbol.getCurrencySymbol())));
        body.add(boldLabel("Sound:"));
        body.add(new JLabel(soundEnabled ? "Enabled" : "Disabled (silent)"));
        add(body, BorderLayout.CENTER);

        JPanel footer = new JPanel(new BorderLayout());
        JLabel note = new JLabel("Preview only - nothing has been saved or triggered.");
        note.setFont(note.getFont().deriveFont(Font.ITALIC, 11f));
        note.setForeground(Color.GRAY);
        note.setBorder(new EmptyBorder(0, 14, 10, 14));
        footer.add(note, BorderLayout.WEST);

        JButton closeButton = new JButton("Close");
        closeButton.addActionListener(e -> dispose());
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 8));
        buttonPanel.add(closeButton);
        footer.add(buttonPanel, BorderLayout.SOUTH);
        add(footer, BorderLayout.SOUTH);

        setMinimumSize(new Dimension(320, 0));
        pack();
        setLocationRelativeTo(owner);

        if (soundEnabled) {
            SettingsManager settings = SettingsManager.getInstance();
            Utils.playAlarmSound(type == AlarmType.HIGH ? settings.getHighAlarmWaveFile() : settings.getLowAlarmWaveFile(),
                    type == AlarmType.HIGH ? Utils.DEFAULT_HIGH_ALARM_SOUND : Utils.DEFAULT_LOW_ALARM_SOUND);
        }

        setVisible(true);
    }

    private static JLabel boldLabel(String text) {
        JLabel label = new JLabel(text);
        label.setFont(label.getFont().deriveFont(Font.BOLD));
        return label;
    }

    private static String formatValue(double price, Double percent, boolean isPercent, SymbolTransaction symbol) {
        if (isPercent && percent != null) {
            return String.format("%.2f%%", percent);
        }
        return Utils.formatCurrencyValue(price, symbol.getCurrencySymbol());
    }
}
