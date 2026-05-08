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
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
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

    /** Width of each pan thumb button in pixels. */
    public static final int PAN_BUTTON_WIDTH = 20;

    /** Number of pixels to move per pan timer tick. */
    private static final int PAN_SCROLL_SPEED = 5;

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
        FIT,     // Panel will grow to fit all text
        SCROLL,  // Panel will rotate the text if text overflows
        PAN      // Clip but allow panning to view more using thumb buttons
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

    /**
     * Current pan offset in pixels; preserved across {@link #cls()} calls so
     * the view position survives content refreshes.
     */
    private int panPosition = 0;

    private boolean contiguousBackground = false;

    @Getter(AccessLevel.NONE)
    @Setter(AccessLevel.NONE)
    private Timer scrollTimer;

    /** Timer that drives continuous panning while a thumb button is held. */
    @Getter(AccessLevel.NONE)
    @Setter(AccessLevel.NONE)
    private Timer panTimer;

    @Setter(AccessLevel.NONE)
    private DISPLAY_STYLE displayStyle = DISPLAY_STYLE.CLIP;

    @Setter(AccessLevel.NONE)
    private int scrollSpeed = SCROLL_SPEED;

    @Getter(AccessLevel.NONE)
    @Setter(AccessLevel.NONE)
    private final CopyOnWriteArrayList<TextItem> items = new CopyOnWriteArrayList<>();

    /** Left (scroll-back) thumb button for PAN mode. */
    @Getter(AccessLevel.NONE)
    @Setter(AccessLevel.NONE)
    private final JButton panLeftButton;

    /** Right (scroll-forward) thumb button for PAN mode. */
    @Getter(AccessLevel.NONE)
    @Setter(AccessLevel.NONE)
    private final JButton panRightButton;

    /**
     * Creates a new ColouredTextPanel with default settings.
     */
    public ColouredTextPanel() {
        setDoubleBuffered(true);
        setOpaque(true);
        setLayout(null);

        panLeftButton  = createPanButton("◀", -PAN_SCROLL_SPEED);
        panRightButton = createPanButton("▶",  PAN_SCROLL_SPEED);

        add(panLeftButton);
        add(panRightButton);
    }

    /**
     * Creates a pan thumb button that scrolls content by {@code delta} pixels per timer tick.
     *
     * @param label The button label (a Unicode arrow character).
     * @param delta Positive moves right, negative moves left.
     * @return The configured {@link JButton}.
     */
    private JButton createPanButton(String label, int delta) {
        JButton btn = new JButton(label);
        btn.setVisible(false);
        btn.setFocusable(false);
        btn.setMargin(new Insets(0, 0, 0, 0));
        btn.setBorder(BorderFactory.createEmptyBorder());
        btn.setBackground(getBackground());
        btn.setFont(btn.getFont().deriveFont(8f));
        btn.setBackground(Color.BLACK);
        btn.setForeground(Color.LIGHT_GRAY);

        MouseAdapter ma = new MouseAdapter() {
            @Override
            public void mousePressed(MouseEvent e) {
                startPanning(delta);
            }

            @Override
            public void mouseReleased(MouseEvent e) {
                stopPanning();
            }

            @Override
            public void mouseExited(MouseEvent e) {
                stopPanning();
            }
        };
        btn.addMouseListener(ma);
        return btn;
    }

    /**
     * Starts the pan timer that shifts the content by {@code delta} pixels each tick.
     *
     * @param delta Pixels to move per tick; positive scrolls right, negative scrolls left.
     */
    private void startPanning(int delta) {
        log.debug("Starting pan timer with delta {}", delta);
        if (panTimer != null && panTimer.isRunning()) {
            panTimer.stop();
        }
        panTimer = new Timer(SCROLL_DELAY, e -> {
            try {
                int maxPan = Math.max(0, totalTextWidth - contentWidth());
                panPosition = Math.max(0, Math.min(panPosition + delta, maxPan));
                updatePanButtons();
                repaint();
            }
            catch (Exception ex) {
                log.error("Error occurred during pan timer action", ex);
            }
        });
        panTimer.start();
    }

    /**
     * Stops the ongoing pan timer.
     */
    private void stopPanning() {
        log.debug("Stopping pan timer");
        if (panTimer != null) {
            panTimer.stop();
        }
    }

    /**
     * Returns the usable content width inside any pan buttons.
     *
     * @return Available pixels for content display.
     */
    private int contentWidth() {
        boolean panButtonsVisible = displayStyle == DISPLAY_STYLE.PAN && totalTextWidth > super.getWidth();
        return panButtonsVisible ? super.getWidth() - PAN_BUTTON_WIDTH * 2 : super.getWidth();
    }

    /**
     * Positions and shows or hides the pan thumb buttons depending on whether
     * the display style is {@link DISPLAY_STYLE#PAN} and the content overflows.
     * This must be called on the Swing thread.
     */
    private void updatePanButtons() {
        boolean show = displayStyle == DISPLAY_STYLE.PAN && totalTextWidth > super.getWidth();
        int h = super.getHeight();

        if (show) {
            int w = super.getWidth();
            panLeftButton.setBounds(0, 0, PAN_BUTTON_WIDTH, h);
            panRightButton.setBounds(w - PAN_BUTTON_WIDTH, 0, PAN_BUTTON_WIDTH, h);

            // Enable/disable based on current pan position
            int maxPan = Math.max(0, totalTextWidth - (w - PAN_BUTTON_WIDTH * 2));
            panLeftButton.setVisible(panPosition > 0);
            panRightButton.setVisible(panPosition < maxPan);
        }
        else {
            panLeftButton.setVisible(false);
            panRightButton.setVisible(false);
        }
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

    /**
     * Repositions and updates the visibility of the pan buttons whenever Swing
     * runs a layout pass on this component (e.g. after a window resize). This
     * ensures the buttons have correct bounds and visibility before they are
     * painted, avoiding the one-frame lag that would occur if this were done
     * only inside {@link #paintComponent}.
     */
    @Override
    public void doLayout() {
        super.doLayout();
        updatePanButtons();
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
            updatePanButtons();
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

        // Create a Graphics2D context for better rendering control
        Graphics2D g2 = (Graphics2D) g;

        if (displayStyle == DISPLAY_STYLE.SCROLL && totalTextWidth > getWidth()) {

            // Draw from cached image
            log.debug("Drawing the image from the cached version with scroll position {}", scrollPosition);
            g2.drawImage(cachedContent, -scrollPosition, 0, null);

            // Draw second copy if needed for seamless rotation
            if (scrollPosition > 0) {
                log.debug("Drawing the image again to account for rotation");
                g2.drawImage(cachedContent, totalTextWidth - scrollPosition, 0, null);
            }
        }
        else if (displayStyle == DISPLAY_STYLE.PAN && totalTextWidth > super.getWidth()) {

            // Clamp pan position to valid range
            int maxPan = Math.max(0, totalTextWidth - contentWidth());
            panPosition = Math.max(0, Math.min(panPosition, maxPan));

            // Clip to the content area so drawing does not bleed under the buttons
            Shape oldClip = g2.getClip();
            g2.setClip(PAN_BUTTON_WIDTH, 0, contentWidth(), super.getHeight());
            log.debug("Drawing panned image at offset {} (panPosition={})", PAN_BUTTON_WIDTH - panPosition, panPosition);
            g2.drawImage(cachedContent, PAN_BUTTON_WIDTH - panPosition, 0, null);
            g2.setClip(oldClip);

            updatePanButtons();
        }
        else {
            // Normal rendering (CLIP, FIT, or PAN without overflow)
            log.debug("Drawing the cached image in normal mode");
            g2.drawImage(cachedContent, 0, 0, null);
            updatePanButtons();
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
     * The pan position is intentionally preserved so that the view
     * offset survives content refreshes in PAN mode.
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
     * Sets the display style and updates pan button visibility accordingly.
     *
     * @param displayStyle The display style to set.
     */
    public void setDisplayStyle(DISPLAY_STYLE displayStyle) {
        log.debug("Changed display style to {}", displayStyle);
        DISPLAY_STYLE oldStyle = this.displayStyle;
        this.displayStyle = displayStyle;

        // If leaving PAN mode, stop any in-progress pan and hide buttons
        if (oldStyle == DISPLAY_STYLE.PAN && displayStyle != DISPLAY_STYLE.PAN) {
            stopPanning();
            panPosition = 0;
        }

        // If switching to/from GROW mode, revalidate
        if (oldStyle != displayStyle && (oldStyle == DISPLAY_STYLE.FIT || displayStyle == DISPLAY_STYLE.FIT)) {
            revalidate();
        }
        updatePanButtons();
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
