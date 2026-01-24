/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.ui.components;

import lombok.AccessLevel;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.util.concurrent.CopyOnWriteArrayList;

/**
 * A custom JPanel that allows printing coloured and styled text
 * at specific positions.
 */
@Slf4j
@Getter
@Setter
public class ColouredTextPanel extends JPanel {

    private static final int SCROLL_SPEED = 2;
    private static final int SCROLL_DELAY = 30;

    @Getter(AccessLevel.NONE)
    @Setter(AccessLevel.NONE)
    private BufferedImage cachedContent;

    @Getter(AccessLevel.NONE)
    @Setter(AccessLevel.NONE)
    private boolean contentDirty = true;

    /**
     * Display styles for text rendering.
     */
    public enum DISPLAY_STYLE {
        CLIP,    // Overflow will be clipped and invisible
        FIT,    // Panel will grow to fit all text
        SCROLL   // Panel will rotate the text if text overflows
    }

    private int currentX = 0;
    private int currentY = 0;
    private Color fontColor = Color.WHITE;
    private boolean fontBold = false;
    private boolean fontItalic = false;
    private float fontSize = 13.3f;
    private String fontFamily = "Arial";
    private int totalTextWidth = 0;
    private int totalTextHeight = 0;
    private int scrollPosition = 0;

    @Getter(AccessLevel.NONE)
    @Setter(AccessLevel.NONE)
    private Timer scrollTimer;

    @Setter(AccessLevel.NONE)
    private DISPLAY_STYLE displayStyle = DISPLAY_STYLE.CLIP;

    @Setter(AccessLevel.NONE)
    private int scrollSpeed = SCROLL_SPEED;

    @Getter(AccessLevel.NONE)
    @Setter(AccessLevel.NONE)
    private final CopyOnWriteArrayList<TextItem> items = new CopyOnWriteArrayList<>();

    /**
     * Creates a new ColouredTextPanel with default settings.
     */
    public ColouredTextPanel() {
        setDoubleBuffered(true);
        setOpaque(true);
    }

    @Override
    public void invalidate() {
        contentDirty = true;
        super.invalidate();
    }

    @Override
    protected void paintComponent(Graphics g) {
        super.paintComponent(g);

        // Recreate cache if needed
        if (contentDirty || cachedContent == null) {
            createCachedContent();
            contentDirty = false;
        }

        // Handle scrolling timer
        if (displayStyle == DISPLAY_STYLE.SCROLL && totalTextWidth > getWidth()) {
            if (scrollTimer == null) {
                scrollTimer = new Timer(SCROLL_DELAY, e -> {
                    scrollPosition += scrollSpeed;
                    if (scrollPosition >= totalTextWidth) {
                        scrollPosition = 0;
                    }
                    repaint();
                });
                scrollTimer.start();
            }
            else if (!scrollTimer.isRunning()) {
                scrollTimer.start();
            }
        }

        // Stop scrolling if not needed
        if (displayStyle != DISPLAY_STYLE.SCROLL || totalTextWidth <= getWidth()) {
            if (scrollTimer != null && scrollTimer.isRunning()) {
                scrollTimer.stop();
                scrollPosition = 0;
            }
        }

        Graphics2D g2 = (Graphics2D) g;
        if (displayStyle == DISPLAY_STYLE.SCROLL && totalTextWidth > getWidth()) {
            // Draw from cached image
            g2.drawImage(cachedContent, -scrollPosition, 0, null);

            // Draw second copy if needed
            if (scrollPosition > 0) {
                g2.drawImage(cachedContent, totalTextWidth - scrollPosition, 0, null);
            }
        }
        else {
            // Normal rendering
            g2.drawImage(cachedContent, 0, 0, null);
        }
    }

    /**
     * Creates the cached content image with all text drawn.
     */
    private void createCachedContent() {
        // If there's nothing to draw, skip
        if (totalTextWidth == 0 || totalTextHeight == 0) {
            return;
        }
        // Ensure we have a valid graphics configuration
        GraphicsConfiguration gc = getGraphicsConfiguration();
        if (gc != null) {
            cachedContent = gc.createCompatibleImage(
                totalTextWidth,
                Math.max(totalTextHeight, getHeight()),
                Transparency.TRANSLUCENT
            );
        }
        else {
            cachedContent = new BufferedImage(
                totalTextWidth,
                Math.max(totalTextHeight, getHeight()),
                BufferedImage.TYPE_INT_ARGB
            );
        }

        // Set the rendering hints for quality
        Graphics2D g2 = cachedContent.createGraphics();
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
        g2.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
        g2.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);

        // Match panel background and font
        g2.setFont(getFont());
        g2.setColor(getFontColor());
        g2.setBackground(getBackground());

        // Draw the text items
        for (TextItem item : items) {
            g2.setFont(item.font);
            g2.setColor(item.color);
            g2.drawString(item.text, item.x, item.y + item.font.getSize());
        }
        g2.dispose();
    }

    @Override
    public void setFont(Font font) {
        super.setFont(font);
        fontFamily = font.getFamily();
        fontSize = font.getSize2D();
        fontBold = font.isBold();
        fontItalic = font.isItalic();
    }

    @Override
    public Dimension getPreferredSize() {
        return displayStyle == DISPLAY_STYLE.FIT
                ? new Dimension(totalTextWidth, totalTextHeight)
                : new Dimension(super.getPreferredSize().width, super.getPreferredSize().height);
    }

    @Override
    public Dimension getMaximumSize() {
        return displayStyle == DISPLAY_STYLE.FIT
                ? new Dimension(totalTextWidth, totalTextHeight)
                : new Dimension(super.getMaximumSize().width, super.getMaximumSize().height);
    }

    @Override
    public Dimension getMinimumSize() {
        return displayStyle == DISPLAY_STYLE.FIT
                ? new Dimension(totalTextWidth, totalTextHeight)
                : new Dimension(super.getMinimumSize().width, super.getMinimumSize().height);
    }

    /**
     * Stops any ongoing scrolling of text
     */
    public void stopScrolling() {
        if (scrollTimer != null) {
            scrollTimer.stop();
        }
    }

    /**
     * Prints the given text with the current styling settings
     * at the current position.
     *
     * @param text The text to print.
     */
    public void print(String text) {
        TextItem item = new TextItem(text, this);
        items.add(item);

        // Move cursor to the end of the printed text
        FontMetrics fm = getFontMetrics(item.font);
        currentX += fm.stringWidth(text);

        // Mark cache as dirty
        contentDirty = true;

        if (displayStyle == DISPLAY_STYLE.FIT) {
            revalidate();
        }
        repaint();
    }

    /**
     * Clears the panel and resets the cursor position.
     */
    public void cls() {
        items.clear();
        currentX = 0;
        currentY = 0;
        totalTextWidth = 0;
        totalTextHeight = 0;
        fontBold = false;
        fontItalic = false;
        contentDirty = true;

        if (displayStyle == DISPLAY_STYLE.FIT) {
            revalidate();
        }
        repaint();
    }

    /**
     * Sets the display style.
     *
     * @param displayStyle The display style to set.
     */
    public void setDisplayStyle(DISPLAY_STYLE displayStyle) {
        DISPLAY_STYLE oldStyle = this.displayStyle;
        this.displayStyle = displayStyle;

        // If switching to/from GROW mode, revalidate
        if (oldStyle != displayStyle && (oldStyle == DISPLAY_STYLE.FIT || displayStyle == DISPLAY_STYLE.FIT)) {
            revalidate();
        }
        repaint();
    }

    /**
     * Sets the scroll speed for scrolling text.
     *
     * @param scrollSpeed The scroll speed to set.
     */
    public void setScrollSpeed(int scrollSpeed) {
        this.scrollSpeed = scrollSpeed;
        repaint();
    }

    /**
     * Represents a text item with its styling and position.
     */
    private static class TextItem {
        String text;
        int x, y, width, height;
        Font font;
        Color color;

        /**
         * Creates a new TextItem with the given text and styling.
         *
         * @param text  The text to print.
         * @param panel The panel to get styling from.
         */
        private TextItem(String text, ColouredTextPanel panel) {
            this.text = text;
            x = panel.getCurrentX();
            y = panel.getCurrentY();

            int style = Font.PLAIN;
            if (panel.isFontBold()) {
                style |= Font.BOLD;
            }
            if (panel.isFontItalic()) {
                style |= Font.ITALIC;
            }
            font = new Font(panel.getFontFamily(), style, 13).deriveFont(panel.getFontSize());
            color = panel.getFontColor();
            width = panel.getFontMetrics(font).stringWidth(text);
            height = panel.getFontMetrics(font).getHeight();
            panel.totalTextWidth = Math.max(x + width, panel.totalTextWidth);
            panel.totalTextHeight = Math.max(y + height, panel.totalTextHeight);
        }
    }
}
