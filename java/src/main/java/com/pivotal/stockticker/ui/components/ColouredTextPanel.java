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
import lombok.experimental.Delegate;
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

    @Delegate
    private final SettingsComponent<ColouredTextPanel> helper = new SettingsComponent<>(this);

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
    private boolean contiguousBackground = false;

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

    /**
     * Creates a ColouredTextPanel aligned to the right and of
     * default size.
     *
     * @return A configured ColouredTextPanel instance.
     */
    public static ColouredTextPanel create() {
        return new ColouredTextPanel();
    }

    @Override
    public void invalidate() {
        log.debug("Invalidating the layout and marking content dirty");
        contentDirty = true;
        super.invalidate();
    }

    @Override
    protected void paintComponent(Graphics g) {
        log.debug("Painting component");
        super.paintComponent(g);

        // Recreate cache if needed
        if (contentDirty || cachedContent == null) {
            createCachedContent();
            contentDirty = false;
        }

        // If no content to display, we're done (super.paintComponent already cleared it)
        if (cachedContent == null || items.isEmpty()) {
            return;
        }

        // Handle scrolling timer
        if (displayStyle == DISPLAY_STYLE.SCROLL && totalTextWidth > getWidth()) {
            log.debug("Creating scroll timer to scroll text");
            if (scrollTimer == null) {
                scrollTimer = new Timer(SCROLL_DELAY, e -> {
                    try {
                        scrollPosition += scrollSpeed;
                        if (scrollPosition >= totalTextWidth) {
                            scrollPosition = 0;
                        }
                        repaint();
                    }
                    catch (Exception ex) {
                        log.error("Error occurred during scroll timer action", ex);
                    }
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
                log.debug("Scroll timer has been stopped and position reset to zero");
                scrollTimer.stop();
                scrollPosition = 0;
            }
        }

        // Create a Graphics2D context for better rendering control when
        // we are scrolling
        Graphics2D g2 = (Graphics2D) g;
        if (displayStyle == DISPLAY_STYLE.SCROLL && totalTextWidth > getWidth()) {

            // Draw from cached image
            log.debug("Drawing the image from the cached version with scroll position {}", scrollPosition);
            g2.drawImage(cachedContent, -scrollPosition, 0, null);

            // Draw second copy if needed
            if (scrollPosition > 0) {
                log.debug("Drawing the image again to account for rotation");
                g2.drawImage(cachedContent, totalTextWidth - scrollPosition, 0, null);
            }
        }
        else {
            // Normal rendering
            log.debug("Drawing the cached image in normal mode");
            g2.drawImage(cachedContent, 0, 0, null);
        }
    }

    /**
     * Creates the cached content image with all text drawn.
     */
    private void createCachedContent() {
        log.debug("Creating cached content image");

        // Determine the size for the cached image
        int width = Math.max(1, totalTextWidth > 0 ? totalTextWidth : getWidth());
        int height = Math.max(1, totalTextHeight > 0 ? totalTextHeight : getHeight());

        // Ensure we have a valid graphics configuration
        GraphicsConfiguration gc = getGraphicsConfiguration();
        if (gc != null) {
            log.debug("Creating compatible image for cached content width:{} height:{}", width, height);
            cachedContent = gc.createCompatibleImage(width, height, Transparency.TRANSLUCENT);
        }
        else {
            log.debug("Creating buffered image for cached content width:{} height:{}", width, height);
            cachedContent = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
        }

        // Set the rendering hints for quality
        Graphics2D g2 = cachedContent.createGraphics();
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
        g2.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
        g2.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);

        // Clear the background
        g2.setComposite(AlphaComposite.Clear);
        g2.fillRect(0, 0, width, height);
        g2.setComposite(AlphaComposite.SrcOver);

        // If there's nothing to draw, we're done
        if (items.isEmpty()) {
            g2.dispose();
            return;
        }

        // Match panel background and font
        g2.setFont(getFont());
        g2.setColor(getFontColor());
        g2.setBackground(getBackground());

        // Draw the text items
        for (int i = 0; i < items.size(); i++) {
            TextItem item = items.get(i);
            log.debug("Drawing text item {}", item);
            g2.setFont(item.font);
            g2.setColor(item.background);

            // If this is a contiguous block then we need to look back to
            // all previous contiguous items with the same background color
            if (item.contiguousBackground) {
                int x = item.x;
                int y = item.y;
                width = item.width;
                height = item.height;

                // Find the start of the contiguous block
                int prev = i;
                while (prev > 0 && items.get(prev - 1).contiguousBackground && items.get(prev - 1).background == item.background) {
                    prev--;
                    x = items.get(prev).x;
                    y = items.get(prev).y;
                    width += item.x - items.get(prev).x;
                    height += item.y - items.get(prev).y;
                }

                // Fill the background rectangle
                g2.fillRect(x, y, width, height);

                // Redraw all the text in the block
                for (int n = prev; n < i; n++) {
                    TextItem tmp = items.get(n);
                    g2.setFont(tmp.font);
                    g2.setColor(tmp.color);
                    g2.drawString(tmp.text, tmp.x, tmp.y + tmp.font.getSize());
                }
            }
            else {
                g2.fillRect(item.x, item.y, item.width, item.height);
            }

            g2.setColor(item.color);
            g2.drawString(item.text, item.x, item.y + item.font.getSize());
        }
        g2.dispose();
    }

    @Override
    public void setFont(Font font) {
        log.debug("Setting font to {}", font);
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

    @Override
    public int getHeight() {
        return displayStyle == DISPLAY_STYLE.FIT ? getPreferredSize().height :super.getHeight();
    }

    @Override
    public int getWidth() {
        return displayStyle == DISPLAY_STYLE.FIT ? getPreferredSize().width :super.getWidth();
    }

    /**
     * Stops any ongoing scrolling of text
     */
    public void stopScrolling() {
        log.debug("Stop scrolling timer");
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
        log.debug("Printing text {}", text);
        TextItem item = new TextItem(text, this);
        items.add(item);

        // Move cursor to the end of the printed text
        currentX = item.x + item.width;

        // Mark cache as dirty
        contentDirty = true;

        if (displayStyle == DISPLAY_STYLE.FIT) {
            setSize(getPreferredSize());
            revalidate();
        }
        repaint();
    }

    /**
     * Clears the panel and resets the cursor position.
     */
    public void cls() {
        log.debug("Clearing the panel");
        items.clear();
        currentX = 0;
        currentY = 0;
        totalTextWidth = 0;
        totalTextHeight = 0;
        fontBold = false;
        fontItalic = false;
        contentDirty = true;
        cachedContent = null;

        // If in FIT mode, revalidate to adjust size
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
        log.debug("Changed display style to {}", displayStyle);
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
        log.debug("Changed the scroll speed to {}", scrollSpeed);
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
        Color background;
        boolean contiguousBackground;

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
            background = panel.getBackground();
            contiguousBackground = panel.isContiguousBackground();
            width = panel.getFontMetrics(font).stringWidth(text) + 2;
            height = panel.getFontMetrics(font).getHeight();
            panel.totalTextWidth = Math.max(x + width, panel.totalTextWidth);
            panel.totalTextHeight = Math.max(y + height, panel.totalTextHeight);
            log.debug("Created text item: {} textheight:{}", this, panel.totalTextHeight);
        }

        @Override
        public String toString() {
            return text + " [" + x  + "," + y + "," + width + "," + height + "] " + color.toString();
        }
    }
}
