/*
 *
 * Copyright (c) 2026, 4NG and/or its affiliates. All rights reserved.
 * 4NG PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker;

import lombok.extern.slf4j.Slf4j;

/**
 * Provides version information about the application.
 */
@Slf4j
public class VersionInfo {

    /**
     * Gets the implementation version of the application.
     * @return The implementation version as a String.
     */
    public static String getVersion() {
        Package pkg = App.class.getPackage();
        String version = pkg.getImplementationVersion();
        return version != null ? version : "Unknown";
    }

    /**
     * Gets the implementation title of the application.
     * @return The implementation title as a String.
     */
    public static String getTitle() {
        Package pkg = App.class.getPackage();
        return pkg.getImplementationTitle();
    }

    /**
     * Gets the implementation vendor of the application.
     * @return The implementation vendor as a String.
     */
    public static String getVendor() {
        Package pkg = App.class.getPackage();
        return pkg.getImplementationVendor();
    }

    /**
     * Gets a formatted version string including title, version, and vendor.
     * @return A formatted version string.
     */
    public static String getVersionString() {
        return String.format("%s - Version %s by %s",
            getTitle(),
            getVersion(),
            getVendor()
        );
    }
}
