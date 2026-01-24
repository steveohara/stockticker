/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.utils;

/**
 * Callback interface for settings changes
 */
public interface CallbackInterface {

    /**
     * Method called when a change occurs
     *
     * @param source The source object that triggered the change
     */
    void changed(Object source);
}
