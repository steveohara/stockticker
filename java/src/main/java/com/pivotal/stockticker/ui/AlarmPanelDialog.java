/*
 *
 * Copyright (c) 2026, 4NG and/or its affiliates. All rights reserved.
 * 4NG PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.Utils;
import com.pivotal.stockticker.model.SymbolTransaction;
import com.pivotal.stockticker.service.AlarmManager;
import com.pivotal.stockticker.service.AlarmManager.AlarmEvent;
import com.pivotal.stockticker.service.PricesManager;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableCellRenderer;
import java.awt.*;
import java.time.LocalDateTime;
import java.util.List;

/**
 * Compact alarm panel dialog for displaying and managing active stock alarms raised by the
 * {@link AlarmManager}.
 */
@Slf4j
public class AlarmPanelDialog extends JDialog {

    private final AlarmManager alarmManager;
    private final PricesManager prices;
    private final AlarmTableModel tableModel;
    private JTable alarmTable;
    private JLabel statusLabel;

    /**
     * Table model for displaying alarms.
     */
    private class AlarmTableModel extends AbstractTableModel {
        private final String[] columnNames = {"Status", "Symbol", "Type", "Current", "Threshold", "Time"};

        @Override
        public int getRowCount() {
            return alarmManager.getActiveAlarms().size();
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
            List<AlarmEvent> alarms = alarmManager.getActiveAlarms();
            if (rowIndex >= alarms.size()) {
                return "";
            }
            AlarmEvent alarm = alarms.get(rowIndex);
            SymbolTransaction symbol = alarm.getSymbolTransaction();
            return switch (columnIndex) {
                case 0 -> alarm.isMuted() ? "🔇" : "🔔";
                case 1 -> symbol.getDisplayName();
                case 2 -> alarm.getType().toString();
                case 3 -> alarm.isPercent()
                        ? String.format("%.2f%%", currentPercentFor(alarm))
                        : Utils.formatCurrencyValue(currentPriceFor(alarm), symbol.getCurrencySymbol());
                case 4 -> alarm.isPercent()
                        ? String.format("%.1f%%", alarm.getThreshold())
                        : Utils.formatCurrencyValue(alarm.getThreshold(), symbol.getCurrencySymbol());
                case 5 -> formatTimeSince(alarm.getTriggeredTime());
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
     * Returns the live current price for the symbol associated with the given alarm.
     */
    private double currentPriceFor(AlarmEvent alarm) {
        var price = prices.getPrice(alarm.getSymbolTransaction().getCode());
        return price == null ? alarm.getPriceAtTrigger() : price.getCurrentPrice();
    }

    /**
     * Returns the live current percentage change for the symbol associated with the given alarm.
     */
    private double currentPercentFor(AlarmEvent alarm) {
        double pricePaid = alarm.getSymbolTransaction().getPricePaid();
        if (pricePaid == 0) {
            return alarm.getPercentAtTrigger();
        }
        return ((currentPriceFor(alarm) - pricePaid) * 100) / pricePaid;
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

            List<AlarmEvent> alarms = alarmManager.getActiveAlarms();
            if (!isSelected && row < alarms.size()) {
                AlarmEvent alarm = alarms.get(row);
                if (alarm.isMuted()) {
                    c.setForeground(Color.GRAY);
                    c.setBackground(new Color(245, 245, 245));
                }
                else {
                    Color bgColor = alarm.getType() == AlarmManager.AlarmType.HIGH
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
     * @param parent       Parent frame
     * @param alarmManager Alarm manager tracking the active alarms
     * @param prices       Prices manager, used to show the live current price/percentage
     */
    public AlarmPanelDialog(Frame parent, AlarmManager alarmManager, PricesManager prices) {
        super(parent, "Active Alarms", false); // Non-modal
        this.alarmManager = alarmManager;
        this.prices = prices;
        this.tableModel = new AlarmTableModel();
        initializeUI();
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

        statusLabel = new JLabel();
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
        alarmTable.getColumnModel().getColumn(1).setPreferredWidth(120); // Symbol
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
        muteButton.addActionListener(e -> withSelectedAlarm(alarm -> {
            alarm.setMuted(true);
            tableModel.fireTableDataChanged();
        }));
        buttonPanel.add(muteButton);

        JButton unmuteButton = new JButton("Unmute");
        unmuteButton.setToolTipText("Unmute selected alarm");
        unmuteButton.addActionListener(e -> withSelectedAlarm(alarm -> {
            alarm.setMuted(false);
            tableModel.fireTableDataChanged();
        }));
        buttonPanel.add(unmuteButton);

        JButton dismissButton = new JButton("Dismiss");
        dismissButton.setToolTipText("Dismiss the selected alarm - it will fire again if the price crosses the threshold again");
        dismissButton.addActionListener(e -> withSelectedAlarm(alarmManager::dismissAlarm));
        buttonPanel.add(dismissButton);

        JButton disableButton = new JButton("Disable");
        disableButton.setToolTipText("Dismiss and disable the selected alarm");
        disableButton.addActionListener(e -> withSelectedAlarm(alarm -> {
            int confirm = JOptionPane.showConfirmDialog(this,
                    "Disable the " + alarm.getType() + " alarm for " + alarm.getSymbolTransaction().getDisplayName() + "?\nYou can re-enable it from the Symbols dialog.",
                    "Confirm Disable", JOptionPane.YES_NO_OPTION);
            if (confirm == JOptionPane.YES_OPTION) {
                alarmManager.disableAlarm(alarm);
            }
        }));
        buttonPanel.add(disableButton);

        buttonPanel.add(new JSeparator(SwingConstants.VERTICAL));

        JButton dismissAllButton = new JButton("Dismiss All");
        dismissAllButton.addActionListener(e -> {
            if (!alarmManager.getActiveAlarms().isEmpty()
                    && JOptionPane.showConfirmDialog(this, "Dismiss all " + alarmManager.getActiveAlarms().size() + " alarm(s)?",
                    "Confirm Dismiss All", JOptionPane.YES_NO_OPTION) == JOptionPane.YES_OPTION) {
                alarmManager.dismissAllAlarms();
            }
        });
        buttonPanel.add(dismissAllButton);

        JButton closeButton = new JButton("Close");
        closeButton.addActionListener(e -> setVisible(false));
        buttonPanel.add(closeButton);

        add(buttonPanel, BorderLayout.SOUTH);

        refresh();
    }

    /**
     * Runs the given action against the currently selected alarm, if any.
     */
    private void withSelectedAlarm(java.util.function.Consumer<AlarmEvent> action) {
        int selectedRow = alarmTable.getSelectedRow();
        List<AlarmEvent> alarms = alarmManager.getActiveAlarms();
        if (selectedRow >= 0 && selectedRow < alarms.size()) {
            action.accept(alarms.get(selectedRow));
        }
        else {
            JOptionPane.showMessageDialog(this, "Please select an alarm first",
                    "No Selection", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    private void updateStatusLabel() {
        int count = alarmManager.getActiveAlarms().size();
        statusLabel.setText(count + " active alarm" + (count == 1 ? "" : "s"));
    }

    /**
     * Refreshes the table to reflect the current set of active alarms.
     */
    public void refresh() {
        tableModel.fireTableDataChanged();
        updateStatusLabel();
    }
}
