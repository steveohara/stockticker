package com.pivotal.stockticker;

import com.formdev.flatlaf.FlatLightLaf;
import com.pivotal.stockticker.ui.TickerBar;
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
        try {
            // Set the L&F to something modern
            UIManager.setLookAndFeel(new FlatLightLaf());

            // Check if taskbar is supported
            if (Taskbar.isTaskbarSupported()) {
                Taskbar taskbar = Taskbar.getTaskbar();
                try (InputStream stream = App.class.getResourceAsStream("/icon.png")) {
                    if (stream == null) {
                        throw new Exception("Resource not found: /icon.png");
                    }
                    Image image = ImageIO.read(stream);
                    taskbar.setIconImage(image);
                }
            }
        }
        catch (Exception e) {
            log.error("Cannot set look and feel", e);
        }
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                try {
                    new TickerBar();
                }
                catch (Exception e) {
                    log.error("Cannot create main application window", e);
                }
            }
        });
    }
}
