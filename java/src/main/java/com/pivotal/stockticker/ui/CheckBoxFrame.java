/*
 *
 * Copyright (c) 2026, 4NG and/or its affiliates. All rights reserved.
 * 4NG PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.ui;

import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import java.awt.*;

/**
 * A custom JPanel with a titled border that includes a checkbox in the title area.
 */
@Slf4j
@Getter
public class CheckBoxFrame extends JPanel {

    private static final int TITLE_X = 10;

    private final JCheckBox titleCheckbox;
    private final JPanel contentPanel;

    /**
     * Constructs a CheckboxFrame with the specified title.
     *
     * @param title The title to display in the checkbox.
     */
    public CheckBoxFrame(String title) {
        setLayout(null);
        setOpaque(false);

        titleCheckbox = new JCheckBox(title);
        titleCheckbox.setOpaque(true);
        titleCheckbox.setBounds(TITLE_X, 0, titleCheckbox.getPreferredSize().width, titleCheckbox.getPreferredSize().height);

        contentPanel = new JPanel();
        contentPanel.setOpaque(false);

        add(titleCheckbox);
        add(contentPanel);

        // Listen to the checkbox to enable/disable contents
        titleCheckbox.addActionListener(e -> {
            setEnabled(titleCheckbox.isSelected());
        });
    }

    @Override
    public Insets getInsets() {
        return new Insets(titleCheckbox.getPreferredSize().height / 2, 0, 0, 0);
    }

    @Override
    public Dimension getPreferredSize() {
        Insets insets = getInsets();
        Dimension titleSize = titleCheckbox.getPreferredSize();
        Dimension contentSize = contentPanel.getPreferredSize();
        int width = Math.max(titleSize.width, contentSize.width) + insets.left + insets.right;
        int height = contentSize.height + insets.top + insets.bottom;
        return new Dimension(width, height);
    }

    @Override
    protected void paintComponent(Graphics g) {
        super.paintComponent(g);
        Graphics2D g2 = (Graphics2D) g.create();
        Color shadow = UIManager.getColor("controlShadow");
        Color highlight = UIManager.getColor("controlHighlight");

        int w = getWidth() - 1;
        int h = getHeight() - 1;

        int y = titleCheckbox.getPreferredSize().height / 2;
        int titleWidth = titleCheckbox.getPreferredSize().width;

        g2.setColor(shadow);
        g2.drawLine(0, y, TITLE_X - 2, y);
        g2.drawLine(TITLE_X + titleWidth + 2, y, w, y);

        g2.setColor(highlight);
        g2.drawLine(1, y + 1, TITLE_X - 1, y + 1);
        g2.drawLine(TITLE_X + titleWidth + 3, y + 1, w - 1, y + 1);

        // Left / right / bottom borders
        g2.setColor(shadow);
        g2.drawRect(0, y, w, h - y);

        g2.setColor(highlight);
        g2.drawRect(1, y + 1, w - 2, h - y - 2);

        g2.dispose();
    }

    @Override
    public void doLayout() {
        Insets in = getInsets();
        contentPanel.setBounds(in.left, in.top, getWidth() - in.left - in.right, getHeight() - in.top - in.bottom);
    }

    @Override
    public void setEnabled(boolean enabled) {
        super.setEnabled(enabled);
        titleCheckbox.setSelected(enabled);
        setEnabledRecursive(contentPanel, enabled);
    }

    /** Recursively sets the enabled state of all components within a container.
     *
     * @param c       The container whose components' enabled state is to be set.
     * @param enabled The enabled state to set.
     */
    private void setEnabledRecursive(Container c, boolean enabled) {
        for (Component comp : c.getComponents()) {
            comp.setEnabled(enabled);
            if (comp instanceof Container) {
                setEnabledRecursive((Container) comp, enabled);
            }
        }
    }
}

