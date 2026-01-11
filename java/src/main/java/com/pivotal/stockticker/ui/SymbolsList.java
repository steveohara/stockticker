/*
 *
 * Copyright (c) 2026, 4NG and/or its affiliates. All rights reserved.
 * 4NG PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.model.SymbolTransaction;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import java.awt.*;

/**
 * Custom JList to display SymbolTransaction objects with specific rendering based on their state.
 */
@Getter
@Slf4j
public class SymbolsList extends JList<SymbolTransaction> {

    private final DefaultListModel<SymbolTransaction> model = new DefaultListModel<>();

    /**
     * Constructor
     */
    public SymbolsList() {
        super();
        setModel(model);
        initialize();
    }

    /**
     * Initialize the custom list
     */
    private void initialize() {
        // Set custom cell renderer
        setCellRenderer(new ListCellRenderer<>() {
            private final SymbolsListCellRenderer defaultRenderer = new SymbolsListCellRenderer();

            @Override
            public Component getListCellRendererComponent(JList<? extends SymbolTransaction> list, SymbolTransaction value, int index, boolean isSelected, boolean cellHasFocus) {
                JLabel renderer = (JLabel) defaultRenderer.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);

                String text = "<html>" + escapeHtml(value.getCode());
                if (value.isAdded()) {
                    text += " <b>*</b>";
                }
                if (value.isEdited()) {
                    text += " <b>*</b>";
                }
                text += "</html>";
                renderer.setText(text);

                text = "<html>";
                if (value.isAdded()) {
                    if (value.isEdited()) {
                        text += "<b>New Symbol Added (edited)</b><br/>";
                    }
                    else {
                        text += "<b>New Symbol Added</b><br/>";
                    }
                }
                else if (value.isEdited()) {
                    text += "<b>Symbol Edited</b><br/>";
                }
                if (value.getAlias() != null && !value.getAlias().isEmpty()) {
                    text += String.format("<b>Alias:</b> %s<br/>", escapeHtml(value.getAlias()));
                }
                text += String.format("<b>Date of Transaction:</b> %s<br/>", escapeHtml(value.getDisplayTimestamp()));
                text += String.format("<b>Bought:</b> %.0f <b>@</b> %s<br/>", value.getSharesBought(), value.getFormattedCost());
                text += String.format("<font style='font-size:0.8em;color:lightgrey'>ID: %s</font><br/>", value.getKey());
                text += "</html>";
                renderer.setToolTipText(text);

                // Set colors based on the disabled state
                if (isSelected) {
                    renderer.setBackground(Color.BLUE);
                    renderer.setForeground(Color.WHITE);
                }
                else if (value.isDisabled()) {
                    renderer.setBackground(Color.WHITE);
                    renderer.setForeground(Color.LIGHT_GRAY);
                }
                else {
                    renderer.setBackground(Color.WHITE);
                    renderer.setForeground(Color.BLACK);
                }

                // Change to italic if it has changed
                renderer.setFont(renderer.getFont().deriveFont(value.isEdited() || value.isAdded() ? Font.ITALIC | Font.BOLD : Font.PLAIN));

                renderer.setOpaque(true);
                return renderer;
            }
        });

    }

    // Helper to prevent HTML injection (if data could contain <, >)
    private String escapeHtml(String s) {
        return s == null ? "" : s.replace("&", "&amp;")
                                 .replace("<", "&lt;")
                                 .replace(">", "&gt;");
    }
    /**
     * Add a SymbolTransaction item to the list
     *
     * @param item SymbolTransaction to add
     * @return Added SymbolTransaction
     */
    public SymbolTransaction addItem(SymbolTransaction item) {
        model.addElement(item);
        return item;
    }

    /**
     * Remove a SymbolTransaction item from the list
     *
     * @param item SymbolTransaction to remove
     */
    public void removeItem(SymbolTransaction item) {
        model.removeElement(item);
    }

    /**
     * Get the key of the selected SymbolTransaction item
     *
     * @return Key of the selected SymbolTransaction, or null if none is selected
     */
    public String getSelectedKey() {
        SymbolTransaction item = getSelectedValue();
        return item != null ? item.getKey() : null;
    }

    /**
     * Get the selected SymbolTransaction item
     *
     * @return Selected SymbolTransaction
     */
    public SymbolTransaction getSelectedListItem() {
        return getSelectedValue();
    }

    /**
     * Clear all items from the list
     */
    public void clear() {
        model.clear();
    }

    /**
     * Custom cell renderer to enable anti-aliased text rendering
     */
    private static class SymbolsListCellRenderer extends DefaultListCellRenderer {
        @Override
        protected void paintComponent(Graphics g) {
            Graphics2D g2 = (Graphics2D) g.create();
            g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
            super.paintComponent(g2);
            g2.dispose();
        }
    }
}
