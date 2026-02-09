package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.Utils;

import javax.swing.*;
import java.awt.*;

/**
 * A custom message dialog that stays on top of all other windows.
 */
public class MessageDialog extends JDialog {

    /**
     * Creates and displays a message dialog that stays on top of all other windows.
     *
     * @param parent      The parent component for the dialog.
     * @param message     The message to display.
     * @param title       The title of the dialog.
     * @param messageType The type of message (e.g., JOptionPane.INFORMATION_MESSAGE).
     */
    private MessageDialog(Component parent, String message, String title, int messageType) {
        JOptionPane pane = new JOptionPane(message, messageType);
        JDialog dialog = pane.createDialog(parent, title);
        dialog.setAlwaysOnTop(true);
        dialog.setModal(true);
        Utils.recenterDialog(dialog, parent);
        dialog.setVisible(true);
        dialog.dispose();
    }

    /**
     * Displays a message dialog that stays on top of all other windows.
     *
     * @param message     The message to display.
     * @param title       The title of the dialog.
     * @param messageType The type of message (e.g., JOptionPane.INFORMATION_MESSAGE).
     */
    public static void show(String message, String title, int messageType) {
        show(null, message, title, messageType);
    }

    /**
     * Displays a message dialog that stays on top of all other windows.
     *
     * @param parent      The parent component for the dialog.
     * @param message     The message to display.
     * @param title       The title of the dialog.
     * @param messageType The type of message (e.g., JOptionPane.INFORMATION_MESSAGE).
     */
    public static void show(Component parent, String message, String title, int messageType) {
        new MessageDialog(parent, message, title, messageType);
    }
}
