/*
 *
 * Copyright (c) 2026, 4NG and/or its affiliates. All rights reserved.
 * 4NG PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.Utils;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import javax.swing.text.*;

/**
 * A JTextField that converts input text to upper case, lower case,
 * keeps it numeric or leaves it unchanged based on the specified conversion type.
 */
@Slf4j
public class CapableTextField extends JTextField {

    /**
     * Enumeration for case conversion types.
     */
    public enum CONVERSION_TYPE {
        UPPER,
        LOWER,
        NUMERIC,
        NONE
    }

    private final CaseDocumentFilter filter;

    /**
     * Constructor
     */
    public CapableTextField() {
        filter = new CaseDocumentFilter(CONVERSION_TYPE.NONE);
        ((AbstractDocument) getDocument()).setDocumentFilter(filter);
    }

    /**
     * Constructor with specified conversion type.
     *
     * @param conversionType The case conversion type.
     */
    public CapableTextField(CONVERSION_TYPE conversionType) {
        filter = new CaseDocumentFilter(conversionType);
        ((AbstractDocument) getDocument()).setDocumentFilter(filter);
    }

    /**
     * Gets the current conversion type.
     *
     * @return The conversion type.
     */
    public CONVERSION_TYPE getConversionType() {
        return filter.getConversionType();
    }

    /**
     * Sets the conversion type.
     *
     * @param conversionType The conversion type to set.
     */
    public void setConversionType(CONVERSION_TYPE conversionType) {
        filter.setConversionType(conversionType);
    }

    /**
     * A DocumentFilter that allows only valid floating-point number input.
     */
    @Setter
    @Getter
    @AllArgsConstructor
    private static class CaseDocumentFilter extends DocumentFilter {

        // The case conversion type
        private CONVERSION_TYPE conversionType;

        @Override
        public void insertString(FilterBypass fb, int offset, String string, AttributeSet attr) throws BadLocationException {
            if (conversionType == CONVERSION_TYPE.NUMERIC) {
                String newText = getNewText(fb, offset, 0, string);
                if (!isAcceptable(newText)) {
                    return;
                }
            }
            string = caseConvert(string);
            super.insertString(fb, offset, string, attr);
        }

        @Override
        public void replace(FilterBypass fb, int offset, int length, String text, AttributeSet attrs) throws BadLocationException {
            if (conversionType == CONVERSION_TYPE.NUMERIC) {
                String newText = getNewText(fb, offset, length, text);
                if (!isAcceptable(newText)) {
                    return;
                }
            }
            text = caseConvert(text);
            super.replace(fb, offset, length, text, attrs);
        }

        @Override
        public void remove(FilterBypass fb, int offset, int length) throws BadLocationException {
            if (conversionType == CONVERSION_TYPE.NUMERIC) {
                String newText = getNewText(fb, offset, length, "");
                if (!isAcceptable(newText)) {
                    return;
                }
            }
            super.remove(fb, offset, length);
        }

        /**
         * Converts the case of the input text based on the conversion type.
         *
         * @param text The input text.
         * @return The converted text.
         */
        private String caseConvert(String text) {
            return switch (conversionType) {
                case CONVERSION_TYPE.UPPER -> text.toUpperCase();
                case CONVERSION_TYPE.LOWER -> text.toLowerCase();
                default -> text;
            };
        }

        /**
         * Constructs the new text after a proposed edit.
         *
         * @param fb     The FilterBypass
         * @param offset The offset of the edit
         * @param length The length of text to replace
         * @param text   The new text to insert
         * @return The resulting text after the edit
         * @throws BadLocationException If accessing the document fails
         */
        private String getNewText(FilterBypass fb, int offset, int length, String text) throws BadLocationException {
            Document doc = fb.getDocument();
            String oldText = doc.getText(0, doc.getLength());
            StringBuilder sb = new StringBuilder(oldText);
            sb.replace(offset, offset + length, text);
            return sb.toString();
        }

        /**
         * Checks if the given text is an acceptable floating-point number format.
         * Accepts "", "+", "-", ".", "+.", "-.", or anything that parses as a float
         *
         * @param text The text to check.
         * @return True if acceptable, false otherwise.
         */
        private boolean isAcceptable(String text) {
            if (text.isEmpty() ||
                    text.equals("+") ||
                    text.equals("-") ||
                    text.equals(".") ||
                    text.equals("+.") ||
                    text.equals("-.")) {
                return true;
            }
            if (!text.matches("[0-9.+-]+")) {
                return false;
            }
            try {
                Double.parseDouble(text);
                return true;
            }
            catch (NumberFormatException e) {
                return false;
            }
        }
    }

    /**
     * Gets the numeric value of the text field.
     *
     * @return The numeric value, or 0.0 if parsing fails.
     */
    public double getValue() {
        return Utils.parseDouble(getText(), 0);
    }

    /**
     * Sets the text of the text field to the string representation of the given double value.
     *
     * @param value The double value to set.
     */
    public void setText(double value) {
        setText(String.valueOf(value));
    }

}
