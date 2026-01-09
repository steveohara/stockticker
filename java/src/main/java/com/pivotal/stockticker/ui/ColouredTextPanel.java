package com.pivotal.stockticker.ui;

import lombok.AccessLevel;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import javax.swing.plaf.ComponentUI;
import java.awt.*;
import java.util.ArrayList;

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
    private int fontSize = 14;
    private String fontFamily = "Arial";
    private int totalTextWidth = 0;
    private int totalTextHeight = 0;

    @Getter(AccessLevel.NONE)
    @Setter(AccessLevel.NONE)
    private Timer scrollTimer;

    @Getter(AccessLevel.NONE)
    @Setter(AccessLevel.NONE)
    private int scrollPosition = 0;

    @Setter(AccessLevel.NONE)
    private DISPLAY_STYLE displayStyle = DISPLAY_STYLE.CLIP;

    @Setter(AccessLevel.NONE)
    private int scrollSpeed = SCROLL_SPEED;

    @Getter(AccessLevel.NONE)
    @Setter(AccessLevel.NONE)
    private final ArrayList<TextItem> items = new ArrayList<>();

    /**
     * Calls the UI delegate's paint method, if the UI delegate
     * is non-<code>null</code>.  We pass the delegate a copy of the
     * <code>Graphics</code> object to protect the rest of the
     * paint code from irrevocable changes
     * (for example, <code>Graphics.translate</code>).
     * <p>
     * If you override this in a subclass you should not make permanent
     * changes to the passed in <code>Graphics</code>. For example, you
     * should not alter the clip <code>Rectangle</code> or modify the
     * transform. If you need to do these operations you may find it
     * easier to create a new <code>Graphics</code> from the passed in
     * <code>Graphics</code> and manipulate it. Further, if you do not
     * invoke super's implementation you must honor the opaque property, that is
     * if this component is opaque, you must completely fill in the background
     * in an opaque color. If you do not honor the opaque property you
     * will likely see visual artifacts.
     * <p>
     * The passed in <code>Graphics</code> object might
     * have a transform other than the identify transform
     * installed on it.  In this case, you might get
     * unexpected results if you cumulatively apply
     * another transform.
     *
     * @param g the <code>Graphics</code> object to protect
     * @see #paint
     * @see ComponentUI
     */
    @Override
    protected void paintComponent(Graphics g) {
        super.paintComponent(g);
        int drawCount = 1;

        // Handle the scrolling style if needed
        if (displayStyle == DISPLAY_STYLE.SCROLL && totalTextWidth > getWidth()) {
            drawCount++;
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

        // Get the graphics context
        Graphics2D g2 = (Graphics2D) g;
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);

        // Draw all text items
        for (int i = 0; i < drawCount; i++) {
            for (TextItem item : items) {
                int startX = (i * totalTextWidth) + item.x - scrollPosition;
                g2.setFont(item.font);
                g2.setColor(item.color);
                g2.drawString(item.text, startX, item.y + item.font.getSize());
                log.debug("Drawing text '{}' at ({}, {}) with font: {}", item.text, startX, item.y + item.font.getSize(), item.font);
            }
        }
    }

    @Override
    public void setFont(Font font) {
        super.setFont(font);
        this.fontFamily = font.getFamily();
        this.fontSize = font.getSize();
        this.fontBold = font.isBold();
        this.fontItalic = font.isItalic();
    }

    /**
     * If the <code>preferredSize</code> has been set to a
     * non-<code>null</code> value just returns it.
     * If the UI delegate's <code>getPreferredSize</code>
     * method returns a non <code>null</code> value then return that;
     * otherwise defer to the component's layout manager.
     *
     * @return the value of the <code>preferredSize</code> property
     * @see #setPreferredSize
     * @see ComponentUI
     */
    @Override
    public Dimension getPreferredSize() {
        return displayStyle == DISPLAY_STYLE.FIT
                ? new Dimension(totalTextWidth, totalTextHeight)
                : new Dimension(super.getPreferredSize().width, totalTextHeight);
    }

    /**
     * If the maximum size has been set to a non-<code>null</code> value
     * just returns it.  If the UI delegate's <code>getMaximumSize</code>
     * method returns a non-<code>null</code> value then return that;
     * otherwise defer to the component's layout manager.
     *
     * @return the value of the <code>maximumSize</code> property
     * @see #setMaximumSize
     * @see ComponentUI
     */
    @Override
    public Dimension getMaximumSize() {
        return displayStyle == DISPLAY_STYLE.FIT
                ? new Dimension(totalTextWidth, totalTextHeight)
                : new Dimension(super.getMaximumSize().width, totalTextHeight);
    }

    /**
     * If the minimum size has been set to a non-<code>null</code> value
     * just returns it.  If the UI delegate's <code>getMinimumSize</code>
     * method returns a non-<code>null</code> value then return that; otherwise
     * defer to the component's layout manager.
     *
     * @return the value of the <code>minimumSize</code> property
     * @see #setMinimumSize
     * @see ComponentUI
     */
    @Override
    public Dimension getMinimumSize() {
        return displayStyle == DISPLAY_STYLE.FIT
                ? new Dimension(totalTextWidth, totalTextHeight)
                : new Dimension(super.getMinimumSize().width, totalTextHeight);
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
        if (oldStyle != displayStyle &&
            (oldStyle == DISPLAY_STYLE.FIT || displayStyle == DISPLAY_STYLE.FIT)) {
            revalidate();
        }
        repaint();    }

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
            font = new Font(panel.getFontFamily(),
                    (panel.isFontBold() ? Font.BOLD : Font.PLAIN) |
                    (panel.isFontItalic() ? Font.ITALIC : Font.PLAIN),
                    panel.getFontSize());
            color = panel.getFontColor();
            width = panel.getFontMetrics(font).stringWidth(text);
            height = panel.getFontMetrics(font).getHeight();
            panel.totalTextWidth = Math.max(x + width, panel.totalTextWidth);
            panel.totalTextHeight = Math.max(y + height, panel.totalTextHeight);
        }
    }
}
