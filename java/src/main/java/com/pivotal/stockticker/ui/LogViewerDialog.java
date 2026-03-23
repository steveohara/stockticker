/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.ui;

import com.pivotal.stockticker.utils.LogCapture;
import lombok.extern.slf4j.Slf4j;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.charset.StandardCharsets;
import java.util.function.Consumer;

/**
 * Non-modal, always-on-top dialog that streams the content of the configured SLF4J log,
 * similar to the Unix "tail -f" command.
 *
 * When SLF4J is directed to a file the file is tailed directly. When it is directed to
 * System.out or System.err the dialog subscribes to the {@link LogCapture} ring-buffer
 * that is installed at application startup.
 */
@Slf4j
public class LogViewerDialog extends JDialog {

    private static final int INITIAL_READ_BYTES = 50 * 1024;
    private static final int POLL_INTERVAL_MS = 500;

    private final JTextArea textArea;
    private final JCheckBox autoScroll;

    // File-tail state
    private volatile boolean running = true;
    private Thread tailThread;

    // System.out/err capture state
    private Consumer<String> captureListener;

    /**
     * Creates and displays the log viewer dialog.
     *
     * @param parent Parent frame.
     * @return The dialog instance.
     */
    public static LogViewerDialog show(Frame parent) {
        LogViewerDialog dialog = new LogViewerDialog(parent);
        dialog.setVisible(true);
        return dialog;
    }

    /**
     * Constructs the dialog, builds all UI components, resolves the log source from the
     * SLF4J SimpleLogger system property, and starts either the capture subscription or
     * the file-tail thread as appropriate.
     *
     * @param parent The owning frame; used for relative positioning.
     */
    private LogViewerDialog(Frame parent) {
        super(parent, "Application Logs", false);
        setAlwaysOnTop(true);
        setSize(950, 550);
        setLocationRelativeTo(parent);
        setLayout(new BorderLayout(0, 0));

        // Header showing the log source
        String logFilePath = resolveLogFilePath();
        JLabel headerLabel = new JLabel(" Watching: " + logFilePath);
        headerLabel.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 11));
        headerLabel.setBorder(new EmptyBorder(4, 4, 4, 4));
        headerLabel.setBackground(new Color(60, 60, 60));
        headerLabel.setForeground(new Color(180, 180, 180));
        headerLabel.setOpaque(true);
        add(headerLabel, BorderLayout.NORTH);

        // Log text area — dark terminal style
        textArea = new JTextArea();
        textArea.setEditable(false);
        textArea.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 12));
        textArea.setBackground(new Color(20, 20, 20));
        textArea.setForeground(new Color(204, 204, 204));
        textArea.setCaretColor(new Color(204, 204, 204));
        textArea.setLineWrap(false);

        JScrollPane scrollPane = new JScrollPane(textArea);
        scrollPane.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        scrollPane.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);
        add(scrollPane, BorderLayout.CENTER);

        // Button panel
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 6));
        autoScroll = new JCheckBox("Auto-scroll", true);
        autoScroll.setToolTipText("Automatically scroll to the latest log entries");
        JButton clearButton = new JButton("Clear");
        clearButton.setToolTipText("Clear the displayed log content");
        clearButton.addActionListener(e -> textArea.setText(""));
        JButton closeButton = new JButton("Close");
        closeButton.addActionListener(e -> dispose());
        buttonPanel.add(autoScroll);
        buttonPanel.add(clearButton);
        buttonPanel.add(closeButton);
        add(buttonPanel, BorderLayout.SOUTH);

        // Start streaming
        if (isStreamDestination(logFilePath)) {
            startCapture();
        }
        else {
            startTailing(logFilePath);
        }

        addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent e) {
                stopStreaming();
            }
        });
    }

    // -----------------------------------------------------------------------
    // System.out / System.err capture
    // -----------------------------------------------------------------------

    /**
     * Returns {@code true} when the configured log destination is a standard stream
     * ({@code System.out} or {@code System.err}) rather than a file path, indicating
     * that {@link LogCapture} should be used instead of file tailing.
     *
     * @param logFilePath The value of the {@code org.slf4j.simpleLogger.logFile} property.
     * @return {@code true} if the destination is {@code System.out} or {@code System.err}.
     */
    private boolean isStreamDestination(String logFilePath) {
        return "System.out".equals(logFilePath) || "System.err".equals(logFilePath);
    }

    /**
     * Subscribes to {@link LogCapture}. Populates the text area from the ring buffer
     * immediately, then receives live updates via a listener callback.
     */
    private void startCapture() {
        LogCapture capture = LogCapture.getInstance();
        if (capture == null) {
            textArea.setText("[LogCapture was not installed — cannot display System.out logs]\n");
            return;
        }

        // Populate from the existing buffer first
        String existing = capture.getBuffer();
        if (!existing.isEmpty()) {
            textArea.setText(existing);
            scrollToBottom();
        }

        // Then register for live updates
        captureListener = text -> SwingUtilities.invokeLater(() -> {
            textArea.append(text);
            if (autoScroll.isSelected()) {
                scrollToBottom();
            }
        });
        capture.addListener(captureListener);
    }

    // -----------------------------------------------------------------------
    // File tail
    // -----------------------------------------------------------------------

    /**
     * Starts the background thread that tails the log file.
     */
    private void startTailing(String logFilePath) {
        File logFile = new File(logFilePath);
        if (!logFile.exists()) {
            textArea.setText("[Log file does not exist: " + logFilePath + "]\n"
                    + "[The file will be displayed automatically once it is created]\n");
        }

        tailThread = new Thread(() -> tailFile(logFile), "log-tail-thread");
        tailThread.setDaemon(true);
        tailThread.start();
    }

    /**
     * Runs the tail loop. Seeks near the end for the initial read then polls for new content.
     */
    private void tailFile(File logFile) {
        // Wait for the file to appear if it doesn't exist yet
        while (running && !logFile.exists()) {
            try {
                Thread.sleep(POLL_INTERVAL_MS);
            }
            catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                return;
            }
        }
        if (!running) {
            return;
        }

        try (RandomAccessFile raf = new RandomAccessFile(logFile, "r")) {

            // Seek near the end to show only recent content on first open
            long fileLength = raf.length();
            long startPos = Math.max(0, fileLength - INITIAL_READ_BYTES);
            raf.seek(startPos);

            // If we jumped into the middle of a line, discard the partial line
            if (startPos > 0) {
                raf.readLine();
            }

            appendNewContent(raf);

            while (running) {
                try {
                    Thread.sleep(POLL_INTERVAL_MS);
                }
                catch (InterruptedException e) {
                    Thread.currentThread().interrupt();
                    break;
                }

                // Handle log rotation: if the file has shrunk, reset to the beginning
                if (raf.length() < raf.getFilePointer()) {
                    raf.seek(0);
                }

                appendNewContent(raf);
            }
        }
        catch (IOException e) {
            final String msg = e.getMessage();
            SwingUtilities.invokeLater(() ->
                    textArea.append("\n[Error reading log file: " + msg + "]\n"));
        }
    }

    /**
     * Reads any new bytes since the last position and appends them to the text area.
     */
    private void appendNewContent(RandomAccessFile raf) throws IOException {
        long available = raf.length() - raf.getFilePointer();
        if (available <= 0) {
            return;
        }
        byte[] bytes = new byte[(int) available];
        int read = raf.read(bytes);
        if (read > 0) {
            String text = new String(bytes, 0, read, StandardCharsets.UTF_8);
            SwingUtilities.invokeLater(() -> {
                textArea.append(text);
                if (autoScroll.isSelected()) {
                    scrollToBottom();
                }
            });
        }
    }

    // -----------------------------------------------------------------------
    // Shared helpers
    // -----------------------------------------------------------------------

    /**
     * Moves the text area caret to the very end of its document, causing the scroll pane
     * to snap to the latest content. Must be called on the Swing event dispatch thread.
     */
    private void scrollToBottom() {
        textArea.setCaretPosition(textArea.getDocument().getLength());
    }

    /**
     * Reads the SLF4J SimpleLogger log destination from the system property
     * {@code org.slf4j.simpleLogger.logFile}, defaulting to {@code System.out} when the
     * property is absent (matching SimpleLogger's own default behaviour).
     *
     * @return The configured log file path, or {@code "System.out"} if not set.
     */
    private String resolveLogFilePath() {
        return System.getProperty("org.slf4j.simpleLogger.logFile", "System.out");
    }

    /**
     * Stops all active streaming. Interrupts the file-tail thread if running, and
     * deregisters the {@link LogCapture} listener if one was registered. Safe to call
     * multiple times.
     */
    private void stopStreaming() {
        // Stop file tail
        running = false;
        if (tailThread != null) {
            tailThread.interrupt();
        }

        // Deregister capture listener
        if (captureListener != null) {
            LogCapture capture = LogCapture.getInstance();
            if (capture != null) {
                capture.removeListener(captureListener);
            }
            captureListener = null;
        }
    }

    @Override
    public void dispose() {
        stopStreaming();
        super.dispose();
    }
}
