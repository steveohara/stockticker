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
            private final DefaultListCellRenderer defaultRenderer = new DefaultListCellRenderer();

            @Override
            public Component getListCellRendererComponent(JList<? extends SymbolTransaction> list, SymbolTransaction value, int index, boolean isSelected, boolean cellHasFocus) {
                JLabel renderer = (JLabel) defaultRenderer.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);

                renderer.setText(value.getCode() + " (ID: " + value.getKey() + ")");

                if (!value.isDisabled()) {
                    renderer.setBackground(isSelected ? Color.MAGENTA : Color.RED);
                    renderer.setForeground(Color.WHITE);
                }
                else {
                    renderer.setBackground(Color.WHITE);
                    renderer.setForeground(Color.BLACK);
                }
                renderer.setOpaque(true);
                return renderer;
            }
        });
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
}
