/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.utils;

import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * Manages application startup settings for different operating systems.
 */
@Slf4j
public class StartupManager {

    // The identifiers used to register the app for startup - must match the values used
    // when packaging the app with jpackage (see the mac/windows profiles in pom.xml)
    private static final String WINDOWS_APP_NAME = "pivotalstockticker";
    private static final String MAC_BUNDLE_ID = "com.pivotal.stockticker";

    // Determine the operating system
    private static final String OS = System.getProperty("os.name").toLowerCase();

    /**
     * Checks if the current OS is Windows.
     *
     * @return true if Windows, false otherwise.
     */
    public static boolean isWindows() {
        return OS.contains("win");
    }

    /**
     * Checks if the current OS is Mac.
     *
     * @return true if Mac, false otherwise.
     */
    public static boolean isMac() {
        return OS.contains("mac");
    }

    /**
     * Enables startup for the application on the current OS.
     */
    public static void enableStartup(boolean enable) {
        if (enable) {
            if (isWindows()) {
                enableWindowsStartup();
            }
            else if (isMac()) {
                enableMacStartup();
            }
        }
        else {
            if (isWindows()) {
                disableWindowsStartup();
            }
            else if (isMac()) {
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
        if (isWindows()) {

            // Check Windows registry
            try {
                Process process = Runtime.getRuntime().exec(new String[]{
                        "reg", "query", "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run", "/v", WINDOWS_APP_NAME
                });
                return process.waitFor(5, TimeUnit.SECONDS) && process.exitValue() == 0;
            }
            catch (Exception e) {
                log.warn("Cannot query Windows startup registry entry", e);
                return false;
            }
        }
        else if (isMac()) {

            // Check for plist file
            return Files.exists(Paths.get(getMacPlistPath()));
        }

        return false;
    }

    /**
     * Enables Windows startup by adding a registry entry.
     */
    private static void enableWindowsStartup() {
        try {
            List<String> launchCommand = getLaunchCommand();
            if (launchCommand.isEmpty()) {
                log.error("Cannot enable Windows startup: unable to determine the application launch command");
                return;
            }

            // Registry values can only hold a single command string. Quote each argument so
            // paths containing spaces (e.g. "Program Files") are preserved correctly.
            String appPath = String.join(" ", launchCommand.stream().map(arg -> "\"" + arg + "\"").toArray(String[]::new));

            Process process = Runtime.getRuntime().exec(new String[]{
                    "reg", "add", "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run",
                    "/v", WINDOWS_APP_NAME, "/d", appPath, "/f"
            });
            if (!process.waitFor(5, TimeUnit.SECONDS) || process.exitValue() != 0) {
                log.error("Cannot enable Windows startup: 'reg add' exited with a non-zero status");
            }
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
            Process process = Runtime.getRuntime().exec(new String[]{
                    "reg", "delete", "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run", "/v", WINDOWS_APP_NAME, "/f"
            });
            if (!process.waitFor(5, TimeUnit.SECONDS) || process.exitValue() != 0) {
                log.error("Cannot disable Windows startup: 'reg delete' exited with a non-zero status");
            }
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
            List<String> launchCommand = getLaunchCommand();
            if (launchCommand.isEmpty()) {
                log.error("Cannot enable Mac startup: unable to determine the application launch command");
                return;
            }

            String plistPath = getMacPlistPath();

            StringBuilder programArguments = new StringBuilder();
            for (String arg : launchCommand) {
                programArguments.append("        <string>").append(escapeXml(arg)).append("</string>\n");
            }

            String plistContent = """
                    <?xml version="1.0" encoding="UTF-8"?>
                    <!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
                        "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
                    <plist version="1.0">
                    <dict>
                        <key>Label</key>
                        <string>%s</string>
                        <key>ProgramArguments</key>
                        <array>
                    %s        </array>
                        <key>RunAtLoad</key>
                        <true/>
                        <key>KeepAlive</key>
                        <false/>
                    </dict>
                    </plist>
                    """.formatted(MAC_BUNDLE_ID, programArguments);

            // Make sure the LaunchAgents directory exists, then write the plist file
            Files.createDirectories(Paths.get(plistPath).getParent());
            Files.write(Paths.get(plistPath), plistContent.getBytes());

            // Unload first in case a stale agent is already loaded, then load the new one
            runAndWait(new String[]{"launchctl", "unload", plistPath});
            Process process = Runtime.getRuntime().exec(new String[]{"launchctl", "load", plistPath});
            if (!process.waitFor(5, TimeUnit.SECONDS) || process.exitValue() != 0) {
                log.error("Cannot enable Mac startup: 'launchctl load' exited with a non-zero status");
            }
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
            String plistPath = getMacPlistPath();
            if (Files.exists(Paths.get(plistPath))) {
                runAndWait(new String[]{"launchctl", "unload", plistPath});
                Files.deleteIfExists(Paths.get(plistPath));
            }
        }
        catch (Exception e) {
            log.error("Cannot disable Mac startup", e);
        }
    }

    /**
     * Runs a command and waits for it to complete, logging a warning if it fails.
     * Used for commands where failure is non-fatal (e.g. unloading an agent that may not be loaded).
     *
     * @param command The command and its arguments.
     */
    private static void runAndWait(String[] command) {
        try {
            Process process = Runtime.getRuntime().exec(command);
            process.waitFor(5, TimeUnit.SECONDS);
        }
        catch (Exception e) {
            log.warn("Command failed: {}", String.join(" ", command), e);
        }
    }

    /**
     * Gets the path to the Mac Launch Agent plist file.
     *
     * @return The plist file path.
     */
    private static String getMacPlistPath() {
        String homeDir = System.getProperty("user.home");
        return homeDir + "/Library/LaunchAgents/" + MAC_BUNDLE_ID + ".plist";
    }

    /**
     * Escapes reserved XML characters so a value can be safely embedded in a plist string.
     *
     * @param value The value to escape.
     * @return The escaped value.
     */
    private static String escapeXml(String value) {
        return value.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;");
    }

    /**
     * Determines the command used to launch this application, as a list of arguments.
     * <p>
     * When running as a jpackage-installed application (Windows .exe or Mac .app), the
     * {@code jpackage.app-path} system property is set by the native launcher to the real
     * installed executable, so that is used directly.
     * <p>
     * Otherwise (e.g. running from an IDE or a plain jar during development), falls back to
     * re-invoking the current JVM against the running jar file.
     *
     * @return The launch command as a list of arguments, or an empty list if it could not be determined.
     */
    private static List<String> getLaunchCommand() {
        List<String> command = new ArrayList<>();
        try {
            String appPath = System.getProperty("jpackage.app-path");
            if (appPath != null && !appPath.isBlank()) {
                command.add(appPath);
                return command;
            }

            // Not running as a packaged app - fall back to "java -jar <path-to-jar>"
            String jarPath = StartupManager.class.getProtectionDomain()
                    .getCodeSource().getLocation().toURI().getPath();
            if (jarPath.endsWith(".jar")) {
                String javaBin = System.getProperty("java.home") + File.separator + "bin" + File.separator +
                        (isWindows() ? "java.exe" : "java");
                command.add(javaBin);
                command.add("-jar");
                command.add(jarPath);
                return command;
            }

            log.warn("Cannot determine launch command: not a packaged app and not running from a jar (path={})", jarPath);
        }
        catch (Exception e) {
            log.error("Cannot determine application launch command", e);
        }
        return command;
    }
}
