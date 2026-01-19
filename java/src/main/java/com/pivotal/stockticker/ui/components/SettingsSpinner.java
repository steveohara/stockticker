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

/**
 * A JSpinner subclass that incorporates settings-related functionalities.
 */
@NoArgsConstructor
public class SettingsSpinner extends JSpinner {

    @Delegate
    private final SettingsComponent<SettingsSpinner> helper = new SettingsComponent<>(this);

    public static final int DEFAULT_WIDTH = 70;
    public static final int DEFAULT_HEIGHT = 20;

    /**
     * Private constructor to enforce the use of the static factory method.
     *
     * @param model The SpinnerModel to use.
     */
    private SettingsSpinner(SpinnerModel model) {
        super(model);
    }

    /**
     * Creates a SettingsSpinner with specified settings.
     *
     * @param value    The initial value.
     * @param minimum  The minimum value.
     * @param maximum  The maximum value.
     * @param stepSize The step size.
     */
    public static SettingsSpinner create(int value, int minimum, int maximum, int stepSize) {
        SettingsSpinner spinner = new SettingsSpinner(new SpinnerNumberModel(value, minimum, maximum, stepSize));
        spinner.setBounds(0, 0, DEFAULT_WIDTH, DEFAULT_HEIGHT);
        return spinner;
    }
}
