package com.pivotal.stockticker;

import com.formdev.flatlaf.FlatLightLaf;
import com.pivotal.stockticker.ui.TickerBar;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;

@Slf4j
public class App {

    /**
     * Main method to launch the application.
     *
     * @param args Command-line arguments.
     */
    public static void main(String[] args) {
        try {
            UIManager.setLookAndFeel(new FlatLightLaf());
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
