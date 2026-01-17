package com.pivotal.stockticker.ui.components;

import lombok.NoArgsConstructor;
import lombok.experimental.Delegate;

import javax.swing.*;

/**
 * A JSeparator subclass that incorporates settings-related functionalities.
 */
@NoArgsConstructor
public class SettingsSeparator extends JSeparator {

    @Delegate
    private final SettingsComponent<SettingsSeparator> helper = new SettingsComponent<>(this);

    public static final int DEFAULT_WIDTH = 70;
    public static final int DEFAULT_HEIGHT = 10;

    /**
     * Creates a SettingsSeparator with specified settings.
     */
    public static SettingsSeparator create() {
        SettingsSeparator component = new SettingsSeparator();
        component.setBounds(0, 0, DEFAULT_WIDTH, DEFAULT_HEIGHT);
        return component;
    }
}
