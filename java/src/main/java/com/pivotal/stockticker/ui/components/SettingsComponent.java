/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.ui.components;

import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import java.awt.*;

/**
 * A class that provides utility methods for setting positions and retrieving dimensions
 * of Swing components.
 */
@Slf4j
public class SettingsComponent<T extends JComponent> {

    private final T component;

    /**
     * Constructs a SettingsComponent for the specified Swing component.
     *
     * @param component The Swing component to manage.
     */
    public SettingsComponent(T component) {
        this.component = component;
    }

    /**
     * Sets the position of the component to the specified x and y coordinates.
     *
     * @param x The x-coordinate.
     * @param y The y-coordinate.
     * @return The component itself for method chaining.
     */
    public T atPosition(int x, int y) {
        component.setBounds(x, y, component.getWidth(), component.getHeight());
        return component;
    }

    /**
     * Sets the top position of the component to the specified y coordinate.
     *
     * @param y The y-coordinate.
     * @return The component itself for method chaining.
     */
    public T atTop(int y) {
        component.setBounds(component.getX(), y, component.getWidth(), component.getHeight());
        return component;
    }

    /**
     * Sets the top position of the component to align with another component.
     *
     * @param alignmentComp The component to align with.
     * @return The component itself for method chaining.
     */
    public T atTop(JComponent alignmentComp) {
        return atTop(alignmentComp, 0);
    }

    /**
     * Sets the top position of the component to align with another component, with an offset.
     *
     * @param alignmentComp The component to align with.
     * @param offset        The offset to apply.
     * @return The component itself for method chaining.
     */
    public T atTop(JComponent alignmentComp, int offset) {
        component.setBounds(component.getX(), alignmentComp.getY() + offset, component.getWidth(), component.getHeight());
        return component;
    }

    /**
     * Sets the left position of the component to the specified x coordinate.
     *
     * @param x The x-coordinate.
     * @return The component itself for method chaining.
     */
    public T atLeft(int x) {
        component.setBounds(x, component.getY(), component.getWidth(), component.getHeight());
        return component;
    }

    /**
     * Sets the left position of the component to align with another component.
     *
     * @param alignmentComp The component to align with.
     * @return The component itself for method chaining.
     */
    public T atLeft(JComponent alignmentComp) {
        return atLeft(alignmentComp, 0);
    }

    /**
     * Sets the left position of the component to align with another component, with an offset.
     *
     * @param alignmentComp The component to align with.
     * @param offset        The offset to apply.
     * @return The component itself for method chaining.
     */
    public T atLeft(JComponent alignmentComp, int offset) {
        component.setBounds(alignmentComp.getX() + offset, component.getY(), component.getWidth(), component.getHeight());
        return component;
    }

    /**
     * Gets the right edge position of the component.
     *
     * @return The right edge position.
     */
    public int getRight() {
        return component.getX() + component.getWidth();
    }

    /**
     * Gets the bottom edge position of the component.
     *
     * @return The bottom edge position.
     */
    public int getBottom() {
        return component.getY() + component.getHeight();
    }

    /**
     * Adds the component to the specified container.
     *
     * @param contentPane The container to add the component to.
     * @return The component itself for method chaining.
     */
    public T to(Container contentPane) {
        contentPane.add(component);
        return component;
    }

    /**
     * Sets the right position of the component to align with another component.
     *
     * @param alignmentComp The component to align with.
     * @return The component itself for method chaining.
     */
    public T atRight(JComponent alignmentComp) {
        return atRight(alignmentComp, 0);
    }

    /**
     * Sets the right position of the component to align with another component, with an offset.
     *
     * @param alignmentComp The component to align with.
     * @param offset        The offset to apply.
     * @return The component itself for method chaining.
     */
    public T atRight(JComponent alignmentComp, int offset) {
        return atLeft(alignmentComp, alignmentComp.getWidth() + offset);
    }

    /**
     * Sets the right position of the component so that the right most
     * coordinate position is x.
     *
     * @param x The x-coordinate.
     * @return The component itself for method chaining.
     */
    public T atRight(int x) {
        return atLeft(x - component.getWidth());
    }

    /**
     * Sets the bottom position of the component to align with another component.
     *
     * @param alignmentComp The component to align with.
     * @return The component itself for method chaining.
     */
    public T atBottom(JComponent alignmentComp) {
        return atBottom(alignmentComp, 0);
    }

    /**
     * Sets the bottom position of the component to align with another component, with an offset.
     *
     * @param alignmentComp The component to align with.
     * @param offset        The offset to apply.
     * @return The component itself for method chaining.
     */
    public T atBottom(JComponent alignmentComp, int offset) {
        return atTop(alignmentComp, alignmentComp.getHeight() + offset);
    }

    /**
     * Sets the bottom position of the component to the specified y coordinate.
     *
     * @param x The y-coordinate.
     * @return The component itself for method chaining.
     */
    public T atBottom(int x) {
        return atLeft(x + component.getHeight());
    }

    /**
     * Sets the width of the component to the specified value.
     *
     * @param width The width to set.
     * @return The component itself for method chaining.
     */
    public T withWidth(int width) {
        component.setBounds(component.getX(), component.getY(), width, component.getHeight());
        return component;
    }

    /**
     * Sets the width of the component to match another component.
     *
     * @param alignmentComp The component to match width with.
     * @return The component itself for method chaining.
     */
    public T withWidth(JComponent alignmentComp) {
        component.setBounds(component.getX(), component.getY(), alignmentComp.getWidth(), component.getHeight());
        return component;
    }

    /**
     * Sets the height of the component to the specified value.
     *
     * @param height The height to set.
     * @return The component itself for method chaining.
     */
    public T withHeight(int height) {
        component.setBounds(component.getX(), component.getY(), component.getWidth(), height);
        return component;
    }

    /**
     * Sets the height of the component to match another component.
     *
     * @param alignmentComp The component to match height with.
     * @return The component itself for method chaining.
     */
    public T withHeight(JComponent alignmentComp) {
        component.setBounds(component.getX(), component.getY(), component.getWidth(), alignmentComp.getHeight());
        return component;
    }

    /**
     * Sets the position of the component to align with the bottom-left corner of another component, with an offset.
     *
     * @param alignmentComp The component to align with.
     * @param offset        The offset to apply.
     * @return The component itself for method chaining.
     */
    public T below(JComponent alignmentComp, int offset) {
        atLeft(alignmentComp);
        atBottom(alignmentComp, offset);
        withDimensions(alignmentComp);
        return component;
    }

    /**
     * Sets the position of the component to align with the top-right corner of another component, with an offset.
     *
     * @param alignmentComp The component to align with.
     * @param offset        The offset to apply.
     * @return The component itself for method chaining.
     */
    public T tail(JComponent alignmentComp, int offset) {
        atTop(alignmentComp);
        atRight(alignmentComp, offset);
        withHeight(alignmentComp);
        return component;
    }

    /**
     * Sets the tooltip text of the component.
     *
     * @param tooltip The tooltip text to set.
     * @return The component itself for method chaining.
     */
    public T setTooltip(String tooltip) {
        component.setToolTipText(tooltip);
        return component;
    }

    /**
     * Sets the background color of the component.
     *
     * @param color The background color to set.
     * @return The component itself for method chaining.
     */
    public T setBackColor(Color color) {
        component.setBackground(color);
        return component;
    }

    /**
     * Sets the foreground color of the component.
     *
     * @param color The foreground color to set.
     * @return The component itself for method chaining.
     */
    public T setForeColor(Color color) {
        component.setForeground(color);
        return component;
    }

    /**
     * Sets the position of the component to the specified x and y coordinates.
     *
     * @param x The x-coordinate.
     * @param y The y-coordinate.
     * @return The component itself for method chaining.
     */
    public T at(int x, int y) {
        component.setBounds(x, y, component.getWidth(), component.getHeight());
        return component;
    }

    /**
     * Sets the dimensions of the component to the specified width and height.
     *
     * @param width  The width to set.
     * @param height The height to set.
     * @return The component itself for method chaining.
     */
    public T withDimensions(int width, int height) {
        component.setBounds(component.getX(), component.getY(), width, height);
        return component;
    }

    /**
     * Sets the dimensions of the component to match another component.
     *
     * @param alignmentComp The component to match dimensions with.
     * @return The component itself for method chaining.
     */
    public T withDimensions(JComponent alignmentComp) {
        component.setBounds(component.getX(), component.getY(), alignmentComp.getWidth(), alignmentComp.getHeight());
        return component;
    }

    /**
     * Sets the border of the component.
     *
     * @param border The border to set.
     * @return The component itself for method chaining.
     */
    public T withBorder(javax.swing.border.Border border) {
        component.setBorder(border);
        return component;
    }

}
