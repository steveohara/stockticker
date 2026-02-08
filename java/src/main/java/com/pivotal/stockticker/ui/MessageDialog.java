package com.pivotal.stockticker.ui;

import javax.swing.*;
import java.awt.*;

/**
 * A custom message dialog that stays on top of all other windows and allows updating the message dynamically.
 */
public class MessageDialog extends JDialog {
    private final JOptionPane pane;
    private final JDialog dialog;
    private final Component parent;

    /**
     * Creates and displays a message dialog that stays on top of all other windows.
     *
     * @param parent      The parent component for the dialog.
     * @param message     The message to display.
     * @param title       The title of the dialog.
     * @param messageType The type of message (e.g., JOptionPane.INFORMATION_MESSAGE).
     */
    public MessageDialog(Component parent, String message, String title, int messageType) {
        this.parent = parent;
        if (messageType == JOptionPane.PLAIN_MESSAGE) {
            pane = new JOptionPane(
                    message,
                    messageType,
                    JOptionPane.DEFAULT_OPTION,
                    null,
                    new Object[]{},  // empty array = no buttons
                    null
            );
        }
        else {
            pane = new JOptionPane(message, messageType);
        }
        dialog = pane.createDialog(parent, title);
        dialog.setAlwaysOnTop(true);
        dialog.setModal(messageType != JOptionPane.PLAIN_MESSAGE);
        recenter();
        dialog.setVisible(true);

    }

    /**
     * Adjusts the position of the dialog to be centered on the screen.
     * For the ticker bar, it centers relative to the ticker bar but
     * aligns to the middle of the ticker bar.
     */
    private void recenter() {
        if (parent instanceof TickerBar) {
            dialog.setLocationRelativeTo(parent);
            Point locationRelativeToTicker = dialog.getLocation();
            dialog.setLocationRelativeTo(parent);
            Point point = parent.getLocationOnScreen();
            dialog.setLocation((int) locationRelativeToTicker.getX(), (int) point.getY());
        }
        else {
            dialog.setLocationRelativeTo(parent);
        }
    }

    /**
     * Updates the message displayed in the dialog and repacks it to adjust size.
     *
     * @param message The new message to display.
     */
    public void showMessage(String message) {
        pane.setMessage(message);
        dialog.pack();
    }

    /**
     * Closes the dialog.
     */
    public void close() {
        dialog.dispose();
    }

    /**
     * Displays a message dialog that stays on top of all other windows.
     *
     * @param message     The message to display.
     * @param title       The title of the dialog.
     * @param messageType The type of message (e.g., JOptionPane.INFORMATION_MESSAGE).
     * @return The dialog instance that was displayed.
     */
    public static MessageDialog create(String message, String title, int messageType) {
        return create(null, message, title, messageType);
    }

    /**
     * Displays a message dialog that stays on top of all other windows.
     *
     * @param parent      The parent component for the dialog.
     * @param message     The message to display.
     * @param title       The title of the dialog.
     * @param messageType The type of message (e.g., JOptionPane.INFORMATION_MESSAGE).
     * @return The dialog instance that was displayed.
     */
    public static MessageDialog create(Component parent, String message, String title, int messageType) {
        return new MessageDialog(parent, message, title, messageType);
    }
}
