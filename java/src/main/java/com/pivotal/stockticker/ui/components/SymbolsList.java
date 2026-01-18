/*
 *
 * Copyright (c) 2026, 4NG and/or its affiliates. All rights reserved.
 * 4NG PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.ui.components;

import com.pivotal.stockticker.model.SymbolTransaction;
import lombok.Getter;
import lombok.experimental.Delegate;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import java.awt.*;
import java.util.ArrayList;
import java.util.stream.Collectors;

/**
 * Custom JList to display SymbolTransaction objects with specific rendering based on their state.
 */
@Getter
@Slf4j
public class SymbolsList extends JList<SymbolTransaction> {

    public static final int DEFAULT_WIDTH = 50;
    public static final int DEFAULT_HEIGHT = 200;

    @Delegate
    private final SettingsComponent<SymbolsList> helper = new SettingsComponent<>(this);
    private final FilterableStockListModel model = new FilterableStockListModel(this);

    /**
     * Constructor
     */
    private SymbolsList() {
        super();
        setModel(model);
        initialize();
    }

    /**
     * Creates a SymbolsList with specified settings.
     */
    public static SymbolsList create() {
        SymbolsList component = new SymbolsList();
        component.setBounds(0, 0, DEFAULT_WIDTH, DEFAULT_HEIGHT);
        return component;
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
     * Set whether to hide disabled SymbolTransaction items
     *
     * @param hideDisabled true to hide disabled items, false to show all
     * @return SymbolsList
     */
    public SymbolsList hideDisabled(boolean hideDisabled) {
        model.hideDisabled(hideDisabled);
        return this;
    }

    /**
     * Check if disabled SymbolTransaction items are hidden
     *
     * @return true if disabled items are hidden, false otherwise
     */
    public boolean isHideDisabled() {
        return model.isHideDisabled();
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

    /**
     * Custom ListModel that supports filtering of disabled SymbolTransaction items.
     */
    @Getter
    public static class FilterableStockListModel extends AbstractListModel<SymbolTransaction> {
        private final ArrayList<SymbolTransaction> backingList;
        private ArrayList<SymbolTransaction> filteredList;
        private boolean hideDisabled = false;
        private final JList<SymbolTransaction> jList;

        /**
         * Constructor
         */
        public FilterableStockListModel(JList<SymbolTransaction> jList) {
            this.jList = jList;
            backingList = new ArrayList<>();
            filteredList = new ArrayList<>();
        }

        /**
         * Set whether to hide disabled SymbolTransaction items
         *
         * @param hideDisabled true to hide disabled items, false to show all
         */
        public void hideDisabled(boolean hideDisabled) {
            this.hideDisabled = hideDisabled;
            updateFilter();
        }

        /**
         * Update the filtered list based on the hideDisabled flag
         */
        private void updateFilter() {
            SymbolTransaction selectedItem = jList.getSelectedValue();
            int selectedIndex = jList.getSelectedIndex();

            int oldSize = filteredList.size();
            if (hideDisabled) {
                filteredList = backingList.stream()
                        .filter(stock -> !stock.isDisabled())
                        .collect(Collectors.toCollection(ArrayList::new));
            }
            else {
                filteredList = new ArrayList<>(backingList);
            }

            // Notify listeners of the change
            fireContentsChanged(this, 0, Math.max(oldSize, filteredList.size()) - 1);

            // Restore selection if item still visible
            if (selectedItem != null) {
                int newIndex = filteredList.indexOf(selectedItem);
                if (newIndex < 0) {
                    jList.clearSelection();
                    if (selectedIndex >= 0 && !filteredList.isEmpty()) {
                        jList.setSelectedIndex(selectedIndex >= filteredList.size() ? filteredList.size() - 1 : selectedIndex);
                    }
                }
            }
        }

        /**
         * Add a SymbolTransaction item to the backing list
         *
         * @param item SymbolTransaction to add
         */
        private void addElement(SymbolTransaction item) {
            backingList.add(item);
            updateFilter();
        }

        /**
         * Remove a SymbolTransaction item from the backing list
         *
         * @param item SymbolTransaction to remove
         */
        private void removeElement(SymbolTransaction item) {
            backingList.remove(item);
            updateFilter();
        }

        /**
         * Clear all items from the backing list
         */
        private void clear() {
            backingList.clear();
            updateFilter();
        }

        /**
         * Check if the filtered list is empty
         *
         * @return true if the filtered list is empty, false otherwise
         */
        public boolean isEmpty() {
            return filteredList.isEmpty();
        }

        @Override
        public int getSize() {
            return filteredList.size();
        }

        @Override
        public SymbolTransaction getElementAt(int index) {
            return filteredList.get(index);
        }
    }
}
