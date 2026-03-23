/*
 *
 * Copyright (c) 2026, Pivotal Solutions and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.stockticker.utils;

import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.function.Consumer;

/**
 * Intercepts System.out and System.err by installing tee PrintStreams that forward all
 * output to the original streams while also buffering it and notifying registered listeners.
 * Install once at application startup via {@link #install()}.
 */
public class LogCapture {

    /** Maximum number of characters retained in the in-memory ring buffer (~512 KB). */
    private static final int MAX_BUFFER_CHARS = 512 * 1024;

    private static LogCapture instance;

    private final StringBuilder buffer = new StringBuilder();
    private final List<Consumer<String>> listeners = new CopyOnWriteArrayList<>();

    /**
     * Private constructor. Replaces {@code System.out} and {@code System.err} with tee
     * streams that forward to the originals while capturing all output.
     */
    private LogCapture() {
        System.setOut(tee(System.out));
        System.setErr(tee(System.err));
    }

    /**
     * Installs the capture streams. Safe to call multiple times; only the first call has effect.
     */
    public static synchronized void install() {
        if (instance == null) {
            instance = new LogCapture();
        }
    }

    /**
     * Returns the singleton instance, or {@code null} if {@link #install()} has not been called.
     */
    public static LogCapture getInstance() {
        return instance;
    }

    /**
     * Returns the current ring-buffer contents as a single String.
     */
    public synchronized String getBuffer() {
        return buffer.toString();
    }

    /**
     * Registers a listener that is called on the thread that produced the output whenever
     * new text arrives. Listeners should be short and non-blocking; use
     * {@code SwingUtilities.invokeLater} inside the listener for any UI work.
     */
    public void addListener(Consumer<String> listener) {
        listeners.add(listener);
    }

    /**
     * Removes a previously registered listener.
     */
    public void removeListener(Consumer<String> listener) {
        listeners.remove(listener);
    }

    // -----------------------------------------------------------------------
    // Internal helpers
    // -----------------------------------------------------------------------

    /**
     * Wraps {@code original} in a tee {@link PrintStream} that writes every byte to both
     * {@code original} and the shared capture buffer via {@link #accept(byte[], int, int)}.
     *
     * @param original The stream to forward output to.
     * @return A new {@link PrintStream} that tees writes to {@code original} and the buffer.
     */
    private PrintStream tee(PrintStream original) {
        OutputStream teeStream = new OutputStream() {
            @Override
            public void write(int b) throws IOException {
                original.write(b);
                accept(new byte[]{(byte) b}, 0, 1);
            }

            @Override
            public void write(byte[] b, int off, int len) throws IOException {
                original.write(b, off, len);
                accept(b, off, len);
            }

            @Override
            public void flush() throws IOException {
                original.flush();
            }
        };
        return new PrintStream(teeStream, true, StandardCharsets.UTF_8);
    }

    /**
     * Converts the raw bytes to a UTF-8 string, appends it to the ring buffer (trimming the
     * oldest content at a line boundary if the buffer exceeds {@link #MAX_BUFFER_CHARS}),
     * then notifies all registered listeners with the new text.
     *
     * @param b   The byte array containing the new output.
     * @param off Offset within {@code b} at which the data starts.
     * @param len Number of bytes to read from {@code b}.
     */
    private synchronized void accept(byte[] b, int off, int len) {
        String text = new String(b, off, len, StandardCharsets.UTF_8);

        buffer.append(text);

        // Trim to keep the buffer within the size limit
        int excess = buffer.length() - MAX_BUFFER_CHARS;
        if (excess > 0) {
            // Remove from the start, but stay on a line boundary to avoid broken lines
            int cut = buffer.indexOf("\n", excess);
            buffer.delete(0, cut >= 0 ? cut + 1 : excess);
        }

        for (Consumer<String> listener : listeners) {
            listener.accept(text);
        }
    }
}
