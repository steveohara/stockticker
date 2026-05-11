/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.utils;

import lombok.extern.slf4j.Slf4j;
import org.slf4j.LoggerFactory;
import org.slf4j.simple.SimpleLogger;

import java.lang.reflect.Field;
import java.util.Map;

/**
 * Utility for dynamically changing the SLF4J SimpleLogger log level at runtime.
 *
 * <p>SimpleLogger caches the log level as a {@code protected int currentLogLevel} field
 * per logger instance and offers no public API for runtime changes.  This class reaches
 * into {@code SimpleLoggerFactory}'s internal logger map via reflection and updates every
 * logger whose name belongs to the application package, then updates the corresponding
 * system property so any loggers created after this call also inherit the new level.
 *
 * <p>All reflection failures are caught and logged as warnings; the application continues
 * to run at whatever level was active before the call.
 */
@Slf4j
public class LogLevelManager {

    /** Root package whose loggers are controlled by this manager. */
    public static final String APP_PACKAGE = "com.pivotal.stockticker";

    /**
     * SimpleLogger system-property key used to override the level for {@link #APP_PACKAGE}.
     * SimpleLogger checks this key (and its parent segments) during logger initialisation,
     * so setting it before a logger is first requested is sufficient for new loggers.
     */
    private static final String LEVEL_PROPERTY = "org.slf4j.simpleLogger.log." + APP_PACKAGE;

    // SimpleLogger integer level constants — mirror LocationAwareLogger values.
    private static final int TRACE = 0;
    private static final int DEBUG = 10;
    private static final int INFO  = 20;
    private static final int WARN  = 30;
    private static final int ERROR = 40;

    private LogLevelManager() {}

    /**
     * Returns the current effective log level name for the application package.
     *
     * <p>Reads the system property that SimpleLogger uses for the application package,
     * falling back to the global default level, then to {@code "INFO"} if neither is set.
     *
     * @return Level name in upper case, e.g. {@code "INFO"}.
     */
    public static String getAppLevel() {
        String prop = System.getProperty(LEVEL_PROPERTY,
                System.getProperty("org.slf4j.simpleLogger.defaultLogLevel", "info"));
        return prop.toUpperCase();
    }

    /**
     * Sets the log level for all active loggers in the application package and for any
     * loggers that are created subsequently.
     *
     * <p>Two things happen:
     * <ol>
     *   <li>The system property {@code org.slf4j.simpleLogger.log.com.pivotal.stockticker}
     *       is updated so that loggers initialised after this call use the new level.</li>
     *   <li>Every live {@code SimpleLogger} instance whose name starts with
     *       {@link #APP_PACKAGE} has its {@code currentLogLevel} field updated via
     *       reflection so the change takes effect immediately.</li>
     * </ol>
     *
     * @param levelName One of {@code TRACE}, {@code DEBUG}, {@code INFO}, {@code WARN},
     *                  {@code ERROR} (case-insensitive).
     */
    public static void setAppLevel(String levelName) {
        int level = parseLevel(levelName);

        // Update system property so loggers created after this call use the new level.
        System.setProperty(LEVEL_PROPERTY, levelName.toLowerCase());

        // Update all live logger instances via reflection.
        try {
            Object factory = LoggerFactory.getILoggerFactory();

            Field mapField = factory.getClass().getDeclaredField("loggerMap");
            mapField.setAccessible(true);
            Map<?, ?> loggerMap = (Map<?, ?>) mapField.get(factory);

            Field levelField = SimpleLogger.class.getDeclaredField("currentLogLevel");
            levelField.setAccessible(true);

            int updated = 0;
            for (Map.Entry<?, ?> entry : loggerMap.entrySet()) {
                if (entry.getKey().toString().startsWith(APP_PACKAGE)) {
                    levelField.set(entry.getValue(), level);
                    updated++;
                }
            }
            log.info("Log level for {} set to {} ({} loggers updated)", APP_PACKAGE, levelName.toUpperCase(), updated);
        }
        catch (Exception e) {
            log.warn("Could not update log levels via reflection: {}", e.getMessage());
        }
    }

    /**
     * Converts a level name to the integer constant used by SimpleLogger.
     *
     * @param levelName The level name (case-insensitive).
     * @return The corresponding SimpleLogger integer constant, defaulting to {@link #INFO}.
     */
    private static int parseLevel(String levelName) {
        return switch (levelName.toUpperCase()) {
            case "TRACE" -> TRACE;
            case "DEBUG" -> DEBUG;
            case "WARN"  -> WARN;
            case "ERROR" -> ERROR;
            default      -> INFO;
        };
    }
}
