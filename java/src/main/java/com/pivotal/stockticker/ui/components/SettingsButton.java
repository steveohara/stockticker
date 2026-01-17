package com.pivotal.stockticker.ui.components;

import lombok.experimental.Delegate;

import javax.swing.*;

/**
 * A SettingsButton subclass that incorporates settings-related functionalities.
 */
public class SettingsButton extends JButton {

    @Delegate
    private final SettingsComponent<SettingsButton> helper = new SettingsComponent<>(this);

    public static final int DEFAULT_LABEL_WIDTH = 60;
    public static final int DEFAULT_LABEL_HEIGHT = 20;

    /**
     * Creates a SettingsButton with specified text.
     *
     * @param text The text of the label.
     */
    public SettingsButton(String text) {
        super(text);
    }

    /**
     * Creates a SettingsButton aligned to the right and of
     * default size.
     *
     * @return A configured SettingsButton instance.
     */
    public static SettingsButton create() {
        return create(null, null);
    }

    /**
     * Creates a SettingsButton with specified text, aligned to the right and of
     * default size.
     *
     * @param text        The text of the label.
     * @return A configured SettingsLabel instance.
     */
    public static SettingsButton create(String text) {
        return create(text, null);
    }

    /**
     * Creates a SettingsButton with specified text and tooltip, aligned to the right.
     *
     * @param text        The text of the label.
     * @param toolTipText The tooltip text for the label.
     * @return A configured SettingsButton instance.
     */
    public static SettingsButton create(String text, String toolTipText) {
        SettingsButton label = new SettingsButton(text);
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
    public SettingsButton setAlignment(int alignment) {
        super.setHorizontalAlignment(alignment);
        return this;
    }
}
