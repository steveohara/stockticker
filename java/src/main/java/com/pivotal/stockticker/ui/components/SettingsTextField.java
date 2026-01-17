package com.pivotal.stockticker.ui.components;

import lombok.NoArgsConstructor;
import lombok.experimental.Delegate;

import javax.swing.*;

/**
 * A JLabel subclass that incorporates settings-related functionalities.
 */
@NoArgsConstructor
public class SettingsTextField extends JTextField {

    @Delegate
    private final SettingsComponent<SettingsTextField> helper = new SettingsComponent<>(this);

    public static final int DEFAULT_WIDTH = 100;
    public static final int DEFAULT_HEIGHT = 20;

    /**
     * Creates a SettingsTextField with specified text.
     *
     * @param text The text of the label.
     */
    public SettingsTextField(String text) {
        super(text);
    }

    /**
     * Creates a SettingsTextField aligned to the right and of
     * default size.
     *
     * @return A configured SettingsTextField instance.
     */
    public static SettingsTextField create() {
        return create(null, null);
    }

    /**
     * Creates a SettingsTextField with specified text of
     * default size.
     *
     * @param text        The text.
     * @return A configured SettingsTextField instance.
     */
    public static SettingsTextField create(String text) {
        return create(text, null);
    }

    /**
     * Creates a SettingsTextField with specified text and tooltip.
     *
     * @param text        The text.
     * @param toolTipText The tooltip text.
     * @return A configured SettingsTextField instance.
     */
    public static SettingsTextField create(String text, String toolTipText) {
        SettingsTextField label = new SettingsTextField(text);
        label.setToolTipText(toolTipText);
        label.setBounds(0, 0, DEFAULT_WIDTH, DEFAULT_HEIGHT);
        return label;
    }
}
