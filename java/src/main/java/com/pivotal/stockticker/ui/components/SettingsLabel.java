/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.ui.components;

import lombok.experimental.Delegate;

import javax.swing.*;

/**
 * A JLabel subclass that incorporates settings-related functionalities.
 */
public class SettingsLabel extends JLabel {

    @Delegate
    private final SettingsComponent<SettingsLabel> helper = new SettingsComponent<>(this);

    public static final int DEFAULT_LABEL_WIDTH = 100;
    public static final int DEFAULT_LABEL_HEIGHT = 20;

    /**
     * Creates a SettingsLabel with specified text.
     *
     * @param text The text of the label.
     */
    public SettingsLabel(String text) {
        super(text);
    }

    /**
     * Creates a SettingsLabel aligned to the right and of
     * default size.
     *
     * @return A configured SettingsLabel instance.
     */
    public static SettingsLabel create() {
        return create(null, null);
    }

    /**
     * Creates a SettingsLabel with specified text, aligned to the right and of
     * default size.
     *
     * @param text        The text of the label.
     * @return A configured SettingsLabel instance.
     */
    public static SettingsLabel create(String text) {
        return create(text, null);
    }

    /**
     * Creates a SettingsLabel with specified text and tooltip, aligned to the right.
     *
     * @param text        The text of the label.
     * @param toolTipText The tooltip text for the label.
     * @return A configured SettingsLabel instance.
     */
    public static SettingsLabel create(String text, String toolTipText) {
        SettingsLabel label = new SettingsLabel(text);
        label.setToolTipText(toolTipText);
        label.setHorizontalAlignment(SwingConstants.RIGHT);
        label.setBounds(0, 0, DEFAULT_LABEL_WIDTH, DEFAULT_LABEL_HEIGHT);
        return label;
    }

    /**
     * Sets the horizontal alignment of the label.
     *
     * @param alignment The horizontal alignment value.
     * @return The SettingsLabel instance for method chaining.
     */
    public SettingsLabel setAlignment(int alignment) {
        super.setHorizontalAlignment(alignment);
        return this;
    }

    /**
     * Sets the characteristics of the SettingsLabel to same as another component.
     *
     * @param hostComponent The SettingsLabel to match with.
     * @return The component itself for method chaining.
     */
    public SettingsLabel sameAs(SettingsLabel hostComponent) {
        helper.sameAs(hostComponent);
        setAlignment(hostComponent.getHorizontalAlignment());
        return this;
    }
}
