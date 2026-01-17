package com.pivotal.stockticker.ui.components;

import lombok.NoArgsConstructor;
import lombok.experimental.Delegate;

import javax.swing.*;

/**
 * A JSpinner subclass that incorporates settings-related functionalities.
 */
@NoArgsConstructor
public class SettingsCheckbox extends JCheckBox {

    @Delegate
    private final SettingsComponent<SettingsCheckbox> helper = new SettingsComponent<>(this);

    public static final int DEFAULT_WIDTH = 100;
    public static final int DEFAULT_HEIGHT = 20;

    /**
     * Private constructor to enforce the use of the static factory method.
     *
     * @param text The text of the checkbox.
     */
    private SettingsCheckbox(String text) {
        super(text);
    }

    /**
     * Creates a SettingsCheckbox with specified text.
     *
     * @param text The text of the checkbox.
     * @return A configured SettingsCheckbox instance.
     */
    public static SettingsCheckbox create(String text) {
        SettingsCheckbox checkbox = new SettingsCheckbox(text);
        checkbox.setBounds(0, 0, DEFAULT_WIDTH, DEFAULT_HEIGHT);
        return checkbox;
    }
}
