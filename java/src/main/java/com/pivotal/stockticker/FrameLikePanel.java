/*
 *
 * Copyright (c) 2026, 4NG and/or its affiliates. All rights reserved.
 * 4NG PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker;

import javax.swing.*;
import java.awt.*;

public class FrameLikePanel extends JPanel {

    private final JCheckBox titleCheckBox;
    private final JPanel contentPanel;

    private static final int TITLE_HEIGHT = 18;
    private static final int TITLE_X = 10;
    private static final int BORDER_GAP = 6;

    public FrameLikePanel(String title) {
        setLayout(null);
        setOpaque(false);

        titleCheckBox = new JCheckBox(title);
        titleCheckBox.setOpaque(true);
        titleCheckBox.setBounds(TITLE_X, 0, 200, TITLE_HEIGHT);

        contentPanel = new JPanel();
        contentPanel.setOpaque(false);

        add(titleCheckBox);
        add(contentPanel);

        // Default behavior: enable/disable content
        titleCheckBox.addActionListener(e ->
                setEnabledRecursive(contentPanel, titleCheckBox.isSelected())
        );
    }

    /** Panel where clients add their components */
    public JPanel getContentPanel() {
        return contentPanel;
    }

    /** Access to the title checkbox */
    public JCheckBox getTitleCheckBox() {
        return titleCheckBox;
    }

    @Override
    public Insets getInsets() {
        return new Insets(TITLE_HEIGHT + BORDER_GAP, 8, 8, 8);
    }

    @Override
    public void doLayout() {
        Insets in = getInsets();

        contentPanel.setBounds(
                in.left,
                in.top,
                getWidth() - in.left - in.right,
                getHeight() - in.top - in.bottom
        );
    }

    @Override
    protected void paintComponent(Graphics g) {
        super.paintComponent(g);

        Graphics2D g2 = (Graphics2D) g.create();

        Color shadow = UIManager.getColor("controlShadow");
        Color highlight = UIManager.getColor("controlHighlight");

        int w = getWidth() - 1;
        int h = getHeight() - 1;
        int y = TITLE_HEIGHT / 2;

        // Top border (split around checkbox)
        int titleWidth = titleCheckBox.getPreferredSize().width;

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

    private void setEnabledRecursive(Container c, boolean enabled) {
        for (Component comp : c.getComponents()) {
            comp.setEnabled(enabled);
            if (comp instanceof Container) {
                setEnabledRecursive((Container) comp, enabled);
            }
        }
    }

    // Demo
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            JFrame frame = new JFrame("GroupBox Demo");
            frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

            FrameLikePanel group = new FrameLikePanel("Advanced Options");
            group.setPreferredSize(new Dimension(300, 160));

            group.getContentPanel().setLayout(new GridLayout(2, 2, 5, 5));
            group.getContentPanel().add(new JLabel("Width:"));
            group.getContentPanel().add(new JTextField());
            group.getContentPanel().add(new JLabel("Height:"));
            group.getContentPanel().add(new JTextField());

            frame.add(group);
            frame.pack();
            frame.setLocationRelativeTo(null);
            frame.setVisible(true);
        });
    }
}
