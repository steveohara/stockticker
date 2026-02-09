package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.Utils;

import javax.swing.*;
import java.awt.*;

/**
 * A custom message dialog that stays on top of all other windows and allows updating the message dynamically.
 */
public class ProgressDialog extends JDialog {
    private final Component parent;
    private final JDialog dialog;
    private final JLabel label;

    /**
     * Creates and displays a message dialog that stays on top of all other windows.
     *
     * @param parent  The parent component for the dialog.
     * @param message The message to display.
     * @param title   The title of the dialog.
     */
    private ProgressDialog(Component parent, String message, String title) {
        this.parent = parent;
        ImageIcon icon = new ImageIcon(getClass().getResource("/busy.gif"));
        label = new JLabel("", icon, JLabel.LEFT);
        label.setPreferredSize(new Dimension(360, 80));
        label.setBorder(BorderFactory.createEmptyBorder(10, 20, 20, 20));

        dialog = new JDialog((Frame) null, title, false);
        dialog.setUndecorated(false);
        JPanel contentPanel = (JPanel) dialog.getContentPane();
        contentPanel.add(label);
        dialog.setAlwaysOnTop(true);
        dialog.setDefaultCloseOperation(JDialog.DO_NOTHING_ON_CLOSE);
        dialog.pack();

        recenter();
        showMessage(message);
        dialog.setVisible(true);
    }

    /**
     * Adjusts the position of the dialog to be centered on the screen.
     * For the ticker bar, it centers relative to the ticker bar but
     * aligns to the middle of the ticker bar.
     */
    private void recenter() {
        Utils.recenterDialog(dialog, parent);
    }

    /**
     * Updates the message displayed in the dialog and repacks it to adjust size.
     *
     * @param message The new message to display.
     */
    public void showMessage(String message) {
        if (message == null) {
            message = "";
        }
        if (message.startsWith("<html>")) {
            label.setText(message);
        }
        else {
            label.setText(String.format("<html><div style='width:240px;margin-left:7px;padding-right:10px'>%s</div></html>", message));
        }
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
     * @param message The message to display.
     * @param title   The title of the dialog.
     * @return The dialog instance that was displayed.
     */
    public static ProgressDialog create(String message, String title) {
        return create(null, message, title);
    }

    /**
     * Displays a message dialog that stays on top of all other windows.
     *
     * @param parent  The parent component for the dialog.
     * @param message The message to display.
     * @param title   The title of the dialog.
     * @return The dialog instance that was displayed.
     */
    public static ProgressDialog create(Component parent, String message, String title) {
        return new ProgressDialog(parent, message, title);
    }
}
