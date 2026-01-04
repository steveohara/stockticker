package com.pivotal.stockticker.ui;

import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import java.awt.*;
import java.util.ArrayList;

@Slf4j
@Getter
@Setter
public class ColouredTextPanel extends JPanel {
    private int currentX = 0;
    private int currentY = 0;
    private Color fontColor = Color.WHITE;
    private boolean fontBold = false;
    private boolean fontItalic = false;
    private int fontSize = 14;
    private String fontFamily = "Arial";

    private final ArrayList<TextItem> items = new ArrayList<>();

    @Override
    protected void paintComponent(Graphics g) {
        super.paintComponent(g);

        // Get the graphics context
        Graphics2D g2 = (Graphics2D) g;
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING,
                            RenderingHints.VALUE_ANTIALIAS_ON);

        // Draw all text items
        for (TextItem item : items) {
            g2.setFont(item.font);
            g2.setColor(item.color);
            g2.drawString(item.text, item.x, item.y + item.font.getSize());
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
     * Prints the given text with the current styling settings
     * at the current position.
     *
     * @param text The text to print.
     */
    public void print(String text) {
        TextItem item = new TextItem(text, this);
        items.add(item);

        // Move cursor to beginning of next line
        FontMetrics fm = getFontMetrics(item.font);
        currentX += fm.stringWidth(text);
        repaint();
    }

    /**
     * Clears the panel and resets the cursor position.
     */
    public void cls() {
        items.clear();
        currentX = 0;
        currentY = 0;
        repaint();
    }

    // Store all printed text for repainting
    private static class TextItem {
        String text;
        int x, y;
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
        }
    }
}
