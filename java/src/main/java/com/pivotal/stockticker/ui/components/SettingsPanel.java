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
 * A JPanel subclass that incorporates settings-related functionalities.
 */
@NoArgsConstructor
public class SettingsPanel extends JPanel {

    @Delegate
    private final SettingsComponent<SettingsPanel> helper = new SettingsComponent<>(this);

    public static final int DEFAULT_WIDTH = 100;
    public static final int DEFAULT_HEIGHT = 100;

    /**
     * Private constructor to enforce the use of the static create method.
     */
    private SettingsPanel(LayoutManager layout) {
        super(layout);
    }

    /**
     * Creates a SettingsPanel with specified settings.
     */
    public static SettingsPanel create() {
        SettingsPanel component = new SettingsPanel();
        component.setBounds(0, 0, DEFAULT_WIDTH, DEFAULT_HEIGHT);
        return component;
    }

    /**
     * Creates a SettingsPanel with specified settings.
     */
    public static SettingsPanel create(LayoutManager layout) {
        SettingsPanel component = new SettingsPanel(layout);
        component.setBounds(0, 0, DEFAULT_WIDTH, DEFAULT_HEIGHT);
        return component;
    }
}
