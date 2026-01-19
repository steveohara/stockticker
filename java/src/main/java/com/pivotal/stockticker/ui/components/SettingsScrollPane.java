/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.ui.components;

import lombok.NoArgsConstructor;
import lombok.experimental.Delegate;

import javax.swing.*;
import java.awt.*;

/**
 * A JScrolPane subclass that incorporates settings-related functionalities.
 */
@NoArgsConstructor
public class SettingsScrollPane extends JScrollPane {

    @Delegate
    private final SettingsComponent<SettingsScrollPane> helper = new SettingsComponent<>(this);

    public static final int DEFAULT_WIDTH = 50;
    public static final int DEFAULT_HEIGHT = 100;

    /**
     * Creates a SettingsSeparator with specified settings.
     */
    public static SettingsScrollPane create() {
        SettingsScrollPane component = new SettingsScrollPane();
        component.setBounds(0, 0, DEFAULT_WIDTH, DEFAULT_HEIGHT);
        return component;
    }

    /**
     * Sets the viewport view component and returns this instance for chaining.
     *
     * @param view the component to set as the viewport view
     * @return this SettingsScrollPane instance
     */
    public SettingsScrollPane withViewport(Component view) {
        setViewportView(view);
        return this;
    }
}
