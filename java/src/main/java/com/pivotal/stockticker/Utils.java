/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker;

import com.pivotal.stockticker.utils.CallbackInterface;
import lombok.extern.slf4j.Slf4j;

import javax.sound.sampled.AudioFileFormat;
import javax.sound.sampled.AudioSystem;
import javax.swing.*;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.JTextComponent;
import java.awt.*;
import java.io.File;
import java.util.Arrays;
import java.util.Comparator;
import java.util.HashSet;
import java.util.Set;

@Slf4j
public class Utils {

    private static JFileChooser fileChooser = null;

    /**
     * Displays a message dialog that stays on top of all other windows.
     *
     * @param message     The message to display.
     * @param title       The title of the dialog.
     * @param messageType The type of message (e.g., JOptionPane.INFORMATION_MESSAGE).
     */
    public static void showTopmostMessage(String message, String title, int messageType) {
        showTopmostMessage(null, message, title, messageType);
    }

    /**
     * Displays a message dialog that stays on top of all other windows.
     *
     * @param parent      The parent component for the dialog.
     * @param message     The message to display.
     * @param title       The title of the dialog.
     * @param messageType The type of message (e.g., JOptionPane.INFORMATION_MESSAGE).
     */
    public static void showTopmostMessage(Component parent, String message, String title, int messageType) {
        JOptionPane pane = new JOptionPane(message, messageType);
        JDialog dialog = pane.createDialog(parent, title);
        dialog.setAlwaysOnTop(true);
        dialog.setModal(true);
        dialog.setVisible(true);
    }

    /**
     * Attaches change listeners to various Swing components within a container.
     *
     * @param container        The container holding the components.
     * @param callback         The callback interface to invoke on changes.
     * @param ignoreComponents Components to ignore when attaching listeners.
     */
    public static void attachChangeListeners(Container container, CallbackInterface callback, Component... ignoreComponents) {
        Set<Component> ignoreComponentSet = new HashSet<>(Arrays.asList(ignoreComponents));
        for (Component c : container.getComponents()) {
            if (ignoreComponentSet.contains(c)) {
                continue;
            }
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

    /**
     * Formats a given value as a currency string with the appropriate currency symbol.
     *
     * @param value          Value to format.
     * @param currencySymbol Currency symbol to prepend.
     * @return Formatted currency string.
     */
    public static String formatCurrencyValue(double value, String currencySymbol) {

        // If this is a lowercase letter, then it should trail
        if (currencySymbol != null && currencySymbol.matches("^[a-z]$")) {
            return  String.format("%,.2f%s", value, currencySymbol);
        }
        if (value < 0) {
            return String.format("-%s%,.2f", currencySymbol, Math.abs(value));
        }
        else {
            return String.format("%s%,.2f", currencySymbol, value);
        }
    }

    /**
     * Recursively dumps the bounds of components in the hierarchy.
     *
     * @param c      The component to dump.
     * @param indent Indentation for formatting.
     */
    private static void dumpBoundsTree(Component c, String indent) {
        Rectangle b = c.getBounds();
        Dimension pref = c.getPreferredSize();

        String extra = "";
        if (c instanceof JLabel l) {
            extra = " text=\"" + l.getText() + "\"";
        }
        else if (c instanceof AbstractButton bttn) {
            extra = " text=\"" + bttn.getText() + "\"";
        }
        else if (c instanceof JTextField tf) {
            extra = " textField";
        }
        else if (c instanceof JComboBox<?> cb) {
            extra = " combo";
        }
        else if (c instanceof JSpinner sp) {
            extra = " spinner";
        }

        System.out.printf(
                "%s%s%s bounds=[x=%d,y=%d,w=%d,h=%d] pref=[w=%d,h=%d]%n",
                indent,
                c.getClass().getSimpleName(),
                extra,
                b.x, b.y, b.width, b.height,
                pref.width, pref.height
        );

        if (c instanceof Container container) {
            for (Component child : container.getComponents()) {
                dumpBoundsTree(child, indent + "  ");
            }
        }
    }

    /**
     * Dumps the bounds of all components in the hierarchy starting from the given root component.
     *
     * @param root The root component to start dumping from.
     */
    public static void dumpAllBounds(Component root) {
        System.out.println("===== BOUNDS DUMP START =====");
        dumpBoundsTree(root, "");
        System.out.println("===== BOUNDS DUMP END =====");
    }


    /**
     * Opens a file chooser dialog to select an audio file.
     *
     * @param base        The parent component for the dialog.
     * @param dialogTitle The title of the dialog.
     * @param filePath    Existing file path to pre-select (can be null).
     * @return The selected audio file path, or null if no file was selected.
     */
    public static String selectAudioFile(Component base, String dialogTitle, String filePath) {

        // Initialise the file chooser if not already done
        if (fileChooser == null) {
            fileChooser = new JFileChooser();
            Arrays.stream(AudioSystem.getAudioFileTypes())
                    .sorted(Comparator.comparing(AudioFileFormat.Type::toString))
                    .forEach(type -> {
                        log.info("Supported audio file type: {}", type);
                        fileChooser.addChoosableFileFilter(new FileNameExtensionFilter(type + " (*." + type.getExtension() + ")", type.getExtension()));
                    });
        }
        fileChooser.setDialogTitle(dialogTitle);
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        fileChooser.setAcceptAllFileFilterUsed(false);
        fileChooser.setMultiSelectionEnabled(false);
        fileChooser.setDialogType(JFileChooser.OPEN_DIALOG);
        if (filePath != null) {
            File file = new File(filePath);
            fileChooser.setCurrentDirectory(file.getParentFile());
            fileChooser.setSelectedFile(file);

            // Find and set matching filter
            String extension = getFileExtension(file);
            FileFilter matchingFilter = findMatchingFilter(fileChooser, extension);
            if (matchingFilter != null) {
                fileChooser.setFileFilter(matchingFilter);
            }
        }
        fileChooser.setApproveButtonText("Select");
        int userSelection = fileChooser.showSaveDialog(base);

        // If user approved, return the selected file path
        if (userSelection == JFileChooser.APPROVE_OPTION) {
            return fileChooser.getSelectedFile().getAbsolutePath();
        }
        return null;
    }

    /**
     * Retrieves the file extension from a given file.
     *
     * @param file The file to extract the extension from.
     * @return The file extension in lowercase, or an empty string if none exists.
     */
    public static String getFileExtension(File file) {
        String name = file.getName();
        int lastDot = name.lastIndexOf('.');
        return (lastDot > 0) ? name.substring(lastDot + 1).toLowerCase() : "";
    }

    /**
     * Finds a matching file filter in the JFileChooser for a given extension.
     *
     * @param chooser   The JFileChooser instance.
     * @param extension The file extension to match.
     * @return The matching FileFilter, or null if none found.
     */
    private static FileFilter findMatchingFilter(JFileChooser chooser, String extension) {
        for (FileFilter filter : chooser.getChoosableFileFilters()) {
            if (filter instanceof FileNameExtensionFilter) {
                FileNameExtensionFilter extFilter = (FileNameExtensionFilter) filter;
                for (String ext : extFilter.getExtensions()) {
                    if (ext.equalsIgnoreCase(extension)) {
                        return filter;
                    }
                }
            }
        }
        return null;
    }

    /**
     * Gets the combined bounds of all screens.
     *
     * @return Rectangle representing the bounds of all screens.
     */
    public static Rectangle getAllScreensBounds() {
        Rectangle allScreensBounds = new Rectangle();

        GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();
        GraphicsDevice[] screens = ge.getScreenDevices();

        for (GraphicsDevice screen : screens) {
            GraphicsConfiguration gc = screen.getDefaultConfiguration();
            Rectangle screenBounds = gc.getBounds();
            allScreensBounds = allScreensBounds.union(screenBounds);
        }

        return allScreensBounds;
    }
}
