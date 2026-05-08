/*
 *
 * Copyright (c) 2026, 4NG and/or its affiliates. All rights reserved.
 * 4NG PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.model.SymbolTransaction;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableCellRenderer;
import java.awt.*;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

/**
 * Compact alarm panel dialog for displaying and managing active stock alarms.
 */
@Slf4j
public class AlarmPanelDialog extends JDialog {
    private final List<AlarmEntry> activeAlarms;
    private final AlarmTableModel tableModel;
    private JTable alarmTable;
    private JLabel statusLabel;

    /**
     * Represents an active alarm entry.
     */
    private static class AlarmEntry {
        SymbolTransaction symbol;
        String type; // "HIGH" or "LOW"
        double threshold;
        double currentPrice;
        LocalDateTime triggeredTime;
        boolean isMuted;

        AlarmEntry(SymbolTransaction symbol, String type, double threshold) {
            this.symbol = symbol;
            this.type = type;
            this.threshold = threshold;
//            this.currentPrice = symbol.getCurrentPrice();
            this.triggeredTime = LocalDateTime.now();
            this.isMuted = false;
        }
    }

    /**
     * Table model for displaying alarms.
     */
    private class AlarmTableModel extends AbstractTableModel {
        private final String[] columnNames = {"Status", "Symbol", "Type", "Current", "Threshold", "Time"};

        @Override
        public int getRowCount() {
            return activeAlarms.size();
        }

        @Override
        public int getColumnCount() {
            return columnNames.length;
        }

        @Override
        public String getColumnName(int column) {
            return columnNames[column];
        }

        @Override
        public Object getValueAt(int rowIndex, int columnIndex) {
            AlarmEntry alarm = activeAlarms.get(rowIndex);
            return switch (columnIndex) {
                case 0 -> alarm.isMuted ? "🔇" : "🔔";
                case 1 -> alarm.symbol.getDisplayName();
                case 2 -> alarm.type;
                case 3 -> String.format("$%.2f", alarm.currentPrice);
                case 4 -> alarm.symbol.isHighAlarmIsPercent() || alarm.symbol.isLowAlarmIsPercent()
                        ? String.format("%.1f%%", alarm.threshold)
                        : String.format("$%.2f", alarm.threshold);
                case 5 -> formatTimeSince(alarm.triggeredTime);
                default -> "";
            };
        }

        private String formatTimeSince(LocalDateTime time) {
            long seconds = java.time.Duration.between(time, LocalDateTime.now()).getSeconds();
            if (seconds < 60) {
                return seconds + "s";
            }
            long minutes = seconds / 60;
            if (minutes < 60) {
                return minutes + "m";
            }
            long hours = minutes / 60;
            return hours + "h";
        }
    }

    /**
     * Custom cell renderer for alarm table.
     */
    private class AlarmCellRenderer extends DefaultTableCellRenderer {
        @Override
        public Component getTableCellRendererComponent(JTable table, Object value,
                                                       boolean isSelected, boolean hasFocus,
                                                       int row, int column) {
            Component c = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

            if (!isSelected) {
                AlarmEntry alarm = activeAlarms.get(row);
                if (alarm.isMuted) {
                    c.setForeground(Color.GRAY);
                    c.setBackground(new Color(245, 245, 245));
                }
                else {
                    Color bgColor = alarm.type.equals("HIGH")
                            ? new Color(255, 240, 240)
                            : new Color(240, 255, 240);
                    c.setBackground(bgColor);
                    c.setForeground(Color.BLACK);
                }
            }

            // Center align status and type columns
            if (column == 0 || column == 2) {
                setHorizontalAlignment(SwingConstants.CENTER);
            }
            else {
                setHorizontalAlignment(SwingConstants.LEFT);
            }

            return c;
        }
    }

    /**
     * Constructs the alarm panel dialog.
     *
     * @param parent  Parent frame
     * @param symbols List of symbols to check for active alarms
     */
    public AlarmPanelDialog(Frame parent, List<SymbolTransaction> symbols) {
        super(parent, "Active Alarms", false); // Non-modal
        this.activeAlarms = new ArrayList<>();
        loadActiveAlarms(symbols);
        this.tableModel = new AlarmTableModel();
        initializeUI();
    }

    /**
     * Loads active alarms from symbols.
     */
    private void loadActiveAlarms(List<SymbolTransaction> symbols) {
        for (SymbolTransaction symbol : symbols) {
            if (symbol.isAlarmShowing()) {
                // Check which alarm was triggered
                if (symbol.isHighAlarmEnabled() && isHighAlarmTriggered(symbol)) {
                    activeAlarms.add(new AlarmEntry(symbol, "HIGH", symbol.getHighAlarmValue()));
                }
                if (symbol.isLowAlarmEnabled() && isLowAlarmTriggered(symbol)) {
                    activeAlarms.add(new AlarmEntry(symbol, "LOW", symbol.getLowAlarmValue()));
                }
            }
        }
    }

    private boolean isHighAlarmTriggered(SymbolTransaction symbol) {
//        double threshold = symbol.isHighAlarmIsPercent() ?
//                symbol.getPrice() * (1 + symbol.getHighAlarmValue() / 100) : symbol.getHighAlarmValue();
//        return symbol.getCurrentPrice() >= threshold;
        return false;
    }

    private boolean isLowAlarmTriggered(SymbolTransaction symbol) {
//        double threshold = symbol.isLowAlarmIsPercent() ?
//                symbol.getPrice() * (1 - symbol.getLowAlarmValue() / 100) : symbol.getLowAlarmValue();
//        return symbol.getCurrentPrice() <= threshold;
        return false;
    }

    /**
     * Initializes the user interface.
     */
    private void initializeUI() {
        setLayout(new BorderLayout(10, 10));
        setSize(700, 400);
        setLocationRelativeTo(getParent());

        // Header panel
        JPanel headerPanel = new JPanel(new BorderLayout());
        headerPanel.setBorder(new EmptyBorder(10, 10, 5, 10));
        headerPanel.setBackground(new Color(240, 240, 240));

        JLabel titleLabel = new JLabel("⚠ Active Alarms");
        titleLabel.setFont(new Font("Arial", Font.BOLD, 16));
        headerPanel.add(titleLabel, BorderLayout.WEST);

        statusLabel = new JLabel(activeAlarms.size() + " active alarm(s)");
        statusLabel.setFont(new Font("Arial", Font.PLAIN, 12));
        statusLabel.setForeground(Color.DARK_GRAY);
        headerPanel.add(statusLabel, BorderLayout.EAST);

        add(headerPanel, BorderLayout.NORTH);

        // Table panel
        alarmTable = new JTable(tableModel);
        alarmTable.setRowHeight(30);
        alarmTable.setShowGrid(true);
        alarmTable.setGridColor(new Color(220, 220, 220));
        alarmTable.setFont(new Font("Arial", Font.PLAIN, 12));
        alarmTable.getTableHeader().setFont(new Font("Arial", Font.BOLD, 12));
        alarmTable.getTableHeader().setBackground(new Color(230, 230, 230));
        alarmTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

        // Set column widths
        alarmTable.getColumnModel().getColumn(0).setPreferredWidth(50);  // Status
        alarmTable.getColumnModel().getColumn(1).setPreferredWidth(120); // SymbolTransaction
        alarmTable.getColumnModel().getColumn(2).setPreferredWidth(60);  // Type
        alarmTable.getColumnModel().getColumn(3).setPreferredWidth(80);  // Current
        alarmTable.getColumnModel().getColumn(4).setPreferredWidth(80);  // Threshold
        alarmTable.getColumnModel().getColumn(5).setPreferredWidth(60);  // Time

        // Apply custom renderer
        AlarmCellRenderer renderer = new AlarmCellRenderer();
        for (int i = 0; i < alarmTable.getColumnCount(); i++) {
            alarmTable.getColumnModel().getColumn(i).setCellRenderer(renderer);
        }

        JScrollPane scrollPane = new JScrollPane(alarmTable);
        scrollPane.setBorder(BorderFactory.createEmptyBorder(0, 10, 10, 10));
        add(scrollPane, BorderLayout.CENTER);

        // Button panel
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 10));

        JButton muteButton = new JButton("Mute");
        muteButton.setToolTipText("Mute selected alarm");
        muteButton.addActionListener(e -> muteSelectedAlarm());
        buttonPanel.add(muteButton);

        JButton unmuteButton = new JButton("Unmute");
        unmuteButton.setToolTipText("Unmute selected alarm");
        unmuteButton.addActionListener(e -> unmuteSelectedAlarm());
        buttonPanel.add(unmuteButton);

        JButton cancelButton = new JButton("Cancel");
        cancelButton.setToolTipText("Cancel selected alarm");
        cancelButton.addActionListener(e -> cancelSelectedAlarm());
        buttonPanel.add(cancelButton);

        buttonPanel.add(new JSeparator(SwingConstants.VERTICAL));

        JButton muteAllButton = new JButton("Mute All");
        muteAllButton.addActionListener(e -> muteAllAlarms());
        buttonPanel.add(muteAllButton);

        JButton cancelAllButton = new JButton("Cancel All");
        cancelAllButton.addActionListener(e -> cancelAllAlarms());
        buttonPanel.add(cancelAllButton);

        JButton closeButton = new JButton("Close");
        closeButton.addActionListener(e -> dispose());
        buttonPanel.add(closeButton);

        add(buttonPanel, BorderLayout.SOUTH);

        // Show empty state if no alarms
        if (activeAlarms.isEmpty()) {
            showEmptyState();
        }
    }

    private void showEmptyState() {
        JPanel emptyPanel = new JPanel(new GridBagLayout());
        emptyPanel.setBorder(new EmptyBorder(20, 20, 20, 20));

        JLabel emptyLabel = new JLabel("No active alarms");
        emptyLabel.setFont(new Font("Arial", Font.PLAIN, 14));
        emptyLabel.setForeground(Color.GRAY);

        emptyPanel.add(emptyLabel);
        add(emptyPanel, BorderLayout.CENTER);
    }

    private void muteSelectedAlarm() {
        int selectedRow = alarmTable.getSelectedRow();
        if (selectedRow >= 0) {
            activeAlarms.get(selectedRow).isMuted = true;
            tableModel.fireTableRowsUpdated(selectedRow, selectedRow);
        }
        else {
            JOptionPane.showMessageDialog(this, "Please select an alarm to mute",
                    "No Selection", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    private void unmuteSelectedAlarm() {
        int selectedRow = alarmTable.getSelectedRow();
        if (selectedRow >= 0) {
            activeAlarms.get(selectedRow).isMuted = false;
            tableModel.fireTableRowsUpdated(selectedRow, selectedRow);
        }
        else {
            JOptionPane.showMessageDialog(this, "Please select an alarm to unmute",
                    "No Selection", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    private void cancelSelectedAlarm() {
        int selectedRow = alarmTable.getSelectedRow();
        if (selectedRow >= 0) {
            AlarmEntry alarm = activeAlarms.get(selectedRow);
            int confirm = JOptionPane.showConfirmDialog(this,
                    "Cancel alarm for " + alarm.symbol.getDisplayName() + "?",
                    "Confirm Cancel", JOptionPane.YES_NO_OPTION);
            if (confirm == JOptionPane.YES_OPTION) {
                alarm.symbol.setAlarmShowing(false);
                activeAlarms.remove(selectedRow);
                tableModel.fireTableRowsDeleted(selectedRow, selectedRow);
                updateStatusLabel();

                if (activeAlarms.isEmpty()) {
                    dispose();
                }
            }
        }
        else {
            JOptionPane.showMessageDialog(this, "Please select an alarm to cancel",
                    "No Selection", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    private void muteAllAlarms() {
        for (AlarmEntry alarm : activeAlarms) {
            alarm.isMuted = true;
        }
        tableModel.fireTableDataChanged();
    }

    private void cancelAllAlarms() {
        int confirm = JOptionPane.showConfirmDialog(this,
                "Cancel all " + activeAlarms.size() + " alarm(s)?",
                "Confirm Cancel All", JOptionPane.YES_NO_OPTION);
        if (confirm == JOptionPane.YES_OPTION) {
            for (AlarmEntry alarm : activeAlarms) {
                alarm.symbol.setAlarmShowing(false);
            }
            activeAlarms.clear();
            tableModel.fireTableDataChanged();
            dispose();
        }
    }

    private void updateStatusLabel() {
        statusLabel.setText(activeAlarms.size() + " active alarm(s)");
    }

    /**
     * Updates the alarm list with new symbol data.
     */
    public void refreshAlarms(List<SymbolTransaction> symbols) {
        activeAlarms.clear();
        loadActiveAlarms(symbols);
        tableModel.fireTableDataChanged();
        updateStatusLabel();

        if (activeAlarms.isEmpty()) {
            dispose();
        }
    }
}
