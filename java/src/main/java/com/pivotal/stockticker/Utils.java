package com.pivotal.stockticker;

import com.pivotal.stockticker.ui.CallbackInterface;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.text.*;
import java.awt.*;

@Slf4j
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

    /**
     * Attaches change listeners to various Swing components within a container.
     *
     * @param container The container holding the components.
     * @param callback  The callback interface to invoke on changes.
     */
    public static void attachChangeListeners(Container container, CallbackInterface callback) {
        for (Component c : container.getComponents()) {

            if (c instanceof JTextComponent text) {
                text.getDocument().addDocumentListener(new DocumentListener() {
                    public void insertUpdate(DocumentEvent e) {
                        callback.changed(c);
                    }

                    public void removeUpdate(DocumentEvent e) {
                        callback.changed(c);
                    }

                    public void changedUpdate(DocumentEvent e) {
                        callback.changed(c);
                    }
                });
            }

            else if (c instanceof AbstractButton btn) {
                btn.addItemListener(e -> callback.changed(c));
            }

            else if (c instanceof JComboBox<?> combo) {
                combo.addActionListener(e -> callback.changed(c));
            }

            else if (c instanceof JSpinner spinner) {
                spinner.addChangeListener(e -> callback.changed(c));
            }

            else if (c instanceof JSlider slider) {
                slider.addChangeListener(e -> callback.changed(c));
            }

            if (c instanceof Container child) {
                attachChangeListeners(child, callback);
            }
        }
    }

    /**
     * Configures a JTextComponent to accept only numeric input.
     *
     * @param textComponent The text component to configure.
     */
    public static void makeTextFieldNumeric(JTextComponent textComponent) {
        ((AbstractDocument) textComponent.getDocument()).setDocumentFilter(new DocumentFilter() {
            @Override
            public void insertString(FilterBypass fb, int offset, String text, AttributeSet attr) throws BadLocationException {
                String content = fb.getDocument().getText(0, fb.getDocument().getLength());
                if (text.matches("[\\d*.]|\\.") && (!text.contains(".") || !content.contains("."))) { // allow only digits
                    super.insertString(fb, offset, text, attr);
                }
            }

            @Override
            public void replace(FilterBypass fb, int offset, int length, String text, AttributeSet attrs) throws BadLocationException {
                String content = fb.getDocument().getText(0, fb.getDocument().getLength());
                if (text.matches("[\\d.]*|\\.") && (!text.contains(".") || !content.contains("."))) { // allow only digits
                    super.replace(fb, offset, length, text, attrs);
                }
            }
        });
    }

    /**
     * Parses a string to a double, returning a default value if parsing fails.
     *
     * @param text The string to parse.
     * @param i    The default value to return on failure.
     * @return The parsed double or the default value.
     */
    public static double parseDouble(String text, int i) {
        try {
            if (text == null || text.isBlank()) {
                return i;
            }
            return Double.parseDouble(text.trim());
        }
        catch (NumberFormatException e) {
            log.debug("Failed to parse double from text '{}', returning default value {}", text, i);
        }
        return i;
    }

    /**
     * Parses a string to an integer, returning a default value if parsing fails.
     *
     * @param text The string to parse.
     * @param i    The default value to return on failure.
     * @return The parsed integer or the default value.
     */
    public static double parseInt(String text, int i) {
        try {
            if (text == null || text.isBlank()) {
                return i;
            }
            return Integer.parseInt(text.trim());
        }
        catch (NumberFormatException e) {
            log.debug("Failed to parse int from text '{}', returning default value {}", text, i);
        }
        return i;
    }
}
