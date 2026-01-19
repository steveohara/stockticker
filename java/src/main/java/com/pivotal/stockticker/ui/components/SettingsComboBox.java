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
 * A JComboBox subclass that incorporates settings-related functionalities.
 */
public class SettingsComboBox<E> extends JComboBox<E> {

    @Delegate
    private final SettingsComponent<SettingsComboBox<E>> helper = new SettingsComponent<>(this);

    public static final int DEFAULT_WIDTH = 100;
    public static final int DEFAULT_HEIGHT = 20;

    /**
     * Private constructor to enforce the use of the static factory method.
     */
    private SettingsComboBox() {
        super();
    }

    /**
     * Creates a SettingsComboBox with specified settings.
     */
    public static <E> SettingsComboBox<E> create() {
        SettingsComboBox<E> combo = new SettingsComboBox<>();
        combo.setBounds(0, 0, DEFAULT_WIDTH, DEFAULT_HEIGHT);
        return combo;
    }
}
