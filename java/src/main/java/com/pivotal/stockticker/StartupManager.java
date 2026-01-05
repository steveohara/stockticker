/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker;

import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;

/**
 * Manages application startup settings for different operating systems.
 */
@Slf4j
public class StartupManager {

    // Determine the operating system
    private static final String OS = System.getProperty("os.name").toLowerCase();

    /**
     * Enables startup for the application on the current OS.
     */
    public static void enableStartup(boolean enable) {
        if (enable) {
            if (OS.contains("win")) {
                enableWindowsStartup();
            }
            else if (OS.contains("mac")) {
                enableMacStartup();
            }
        }
        else {
            if (OS.contains("win")) {
                disableWindowsStartup();
            }
            else if (OS.contains("mac")) {
                disableMacStartup();
            }
        }
    }

    /**
     * Checks if startup is enabled for the application on the current OS.
     *
     * @return true if startup is enabled, false otherwise.
     */
    public static boolean isStartupEnabled() {
        String os = System.getProperty("os.name").toLowerCase();

        if (os.contains("win")) {
            // Check Windows registry
            try {
                Process process = Runtime.getRuntime().exec(
                    "reg query HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run /v YourStockTicker"
                );
                return process.waitFor() == 0;
            } catch (Exception e) {
                return false;
            }
        } else if (os.contains("mac")) {
            // Check for plist file
            String homeDir = System.getProperty("user.home");
            String plistPath = homeDir + "/Library/LaunchAgents/com.yourcompany.stockticker.plist";
            return Files.exists(Paths.get(plistPath));
        }

        return false;
    }

    /**
     * Enables Windows startup by adding a registry entry.
     */
    private static void enableWindowsStartup() {
        try {
            String appPath = new File(".").getCanonicalPath() + "\\YourApp.exe";
            String appName = "YourStockTicker";

            // Add to registry
            String command = String.format(
                    "reg add HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run /v \"%s\" /d \"%s\" /f",
                    appName, appPath
            );

            Runtime.getRuntime().exec(command);
        }
        catch (Exception e) {
            log.error("Cannot enable Windows startup", e);
        }
    }

    /**
     * Disables Windows startup by removing the registry entry.
     */
    private static void disableWindowsStartup() {
        try {
            String appName = "YourStockTicker";
            String command = String.format(
                    "reg delete HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run /v \"%s\" /f",
                    appName
            );

            Runtime.getRuntime().exec(command);
        }
        catch (Exception e) {
            log.error("Cannot disable Windows startup", e);
        }
    }

    /**
     * Enables Mac startup by creating a Launch Agent plist.
     */
    private static void enableMacStartup() {
        try {
            String homeDir = System.getProperty("user.home");
            String plistPath = homeDir + "/Library/LaunchAgents/com.yourcompany.stockticker.plist";

            String plistContent = """
                    <?xml version="1.0" encoding="UTF-8"?>
                    <!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" 
                        "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
                    <plist version="1.0">
                    <dict>
                        <key>Label</key>
                        <string>com.yourcompany.stockticker</string>
                        <key>ProgramArguments</key>
                        <array>
                            <string>%s</string>
                        </array>
                        <key>RunAtLoad</key>
                        <true/>
                        <key>KeepAlive</key>
                        <false/>
                    </dict>
                    </plist>
                    """.formatted(getApplicationPath());

            // Write plist file
            Files.write(Paths.get(plistPath), plistContent.getBytes());

            // Load the launch agent
            Runtime.getRuntime().exec(new String[]{"launchctl", "load", plistPath});
        }
        catch (Exception e) {
            log.error("Cannot enable Mac startup", e);
        }
    }

    /**
     * Disables Mac startup by removing the Launch Agent plist.
     */
    private static void disableMacStartup() {
        try {
            String homeDir = System.getProperty("user.home");
            String plistPath = homeDir + "/Library/LaunchAgents/com.yourcompany.stockticker.plist";

            // Unload and remove
            Runtime.getRuntime().exec(new String[]{"launchctl", "unload", plistPath});
            Files.deleteIfExists(Paths.get(plistPath));
        }
        catch (Exception e) {
            log.error("Cannot disable mac startup", e);
        }
    }

    /**
     * Gets the application path for the startup script.
     *
     * @return The application path.
     */
    private static String getApplicationPath() {
        try {
            // For JAR files
            String path = StartupManager.class.getProtectionDomain()
                    .getCodeSource().getLocation().toURI().getPath();

            // If running from IDE, you might want to return the wrapper script
            if (path.endsWith(".jar")) {
                return "java -jar " + path;
            }

            // For packaged applications (.exe, .app)
            return new File(".").getCanonicalPath() + File.separator + "YourApp";
        }
        catch (Exception e) {
            log.error("Cannot get application path", e);
            return "";
        }
    }
}
