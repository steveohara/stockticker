/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker;

import com.formdev.flatlaf.FlatLightLaf;
import com.pivotal.stockticker.ui.TickerBar;
import com.pivotal.stockticker.utils.LogCapture;
import lombok.extern.slf4j.Slf4j;

import javax.imageio.ImageIO;
import javax.swing.*;
import java.awt.*;
import java.io.InputStream;

@Slf4j
public class App {

    /**
     * Main method to launch the application.
     *
     * @param args Command-line arguments.
     */
    public static void main(String[] args) {
        // Install the log capture tee as early as possible so nothing is missed
        LogCapture.install();

        try {
            // Add the application icon to the taskbar - simply fail through if not found
            try (InputStream stream = App.class.getResourceAsStream("/icon.png")) {
                if (stream == null) {
                    throw new Exception("Resource not found: /icon.png");
                }
                Image image = ImageIO.read(stream);
                setApplicationIcon(image);
            }

            // Set the L&F to something modern
            UIManager.setLookAndFeel(new FlatLightLaf());
        }
        catch (Exception e) {
            log.warn("Cannot set look and feel", e);
        }

        // Launch the application main screen
        SwingUtilities.invokeLater(() -> {
            try {
                new TickerBar();
            }
            catch (Exception e) {
                log.error("Cannot create main application window", e);
            }
        });
    }

    /**
     * Sets the application icon for both the taskbar and all existing frames.
     *
     * @param image The image to set as the application icon.
     */
    public static void setApplicationIcon(Image image) {

        // Try setting taskbar icon (Java 9+, platform-dependent)
        if (Taskbar.isTaskbarSupported()) {
            Taskbar taskbar = Taskbar.getTaskbar();
            if (taskbar.isSupported(Taskbar.Feature.ICON_IMAGE)) {
                try {
                    taskbar.setIconImage(image);
                }
                catch (UnsupportedOperationException e) {
                    System.err.println("Taskbar icon not supported on this platform");
                }
            }
        }

        // Always set frame icon as fallback
        for (Frame frame : Frame.getFrames()) {
            frame.setIconImage(image);
        }
    }
}
