package com.pivotal.stockticker;

import javax.swing.*;

public class Utils {

    /**
     * Displays a message dialog that stays on top of all other windows.
     *
     * @param message     The message to display.
     * @param title       The title of the dialog.
     * @param messageType The type of message (e.g., JOptionPane.INFORMATION_MESSAGE).
     */
    public static void showTopmostMessage(String message, String title, int messageType) {
        JOptionPane pane = new JOptionPane(message, messageType);
        JDialog dialog = pane.createDialog(null, title);
        dialog.setAlwaysOnTop(true);
        dialog.setModal(true);
        dialog.setVisible(true);
    }
}
