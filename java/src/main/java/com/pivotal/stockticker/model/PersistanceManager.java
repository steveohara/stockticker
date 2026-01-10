package com.pivotal.stockticker.model;

import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import net.bytebuddy.ByteBuddy;
import net.bytebuddy.implementation.MethodDelegation;
import net.bytebuddy.implementation.bind.annotation.*;
import net.bytebuddy.matcher.ElementMatchers;

import java.awt.*;
import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.Arrays;
import java.util.Base64;
import java.util.concurrent.Callable;
import java.util.prefs.BackingStoreException;
import java.util.prefs.Preferences;

/**
 * Application settings and configuration
 * These changes will be persisted to settings.json automatically
 */
@Slf4j
abstract public class PersistanceManager {

    // Root node for all preferences
    public static final String ROOT_NODE = "/stockticker/";

    // Preferences instance
    private Preferences prefs;

    // Auto-save flag - if true, changes are automatically saved to preferences
    @Getter
    @Setter
    private boolean autoSave = true;

    /**
     * Loads the field value from preferences.
     *
     * @param field The field to save.
     */
    private void loadField(Field field) {

        // Ignore non-setable fields
        if (Modifier.isStatic(field.getModifiers())) {
            return;
        }

        field.setAccessible(true);
        try {
            Class<?> type = field.getType();
            String key = field.getName();
            Object target = field.getDeclaringClass().cast(this);

            if (type == String.class) {
                String value = prefs.get(key, field.get(target) == null ? "" : field.get(target).toString());
                field.set(target, value);
            }
            else if (type == int.class || type == Integer.class) {
                int value = prefs.getInt(key, field.get(target) == null ? 0 : field.getInt(target));
                field.set(target, value);
            }
            else if (type == long.class || type == Long.class) {
                long value = prefs.getLong(key, field.get(target) == null ? 0 : field.getLong(target));
                field.set(target, value);
            }
            else if (type == boolean.class || type == Boolean.class) {
                boolean value = prefs.getBoolean(key, field.get(target) != null && field.getBoolean(target));
                field.set(target, value);
            }
            else if (type == double.class || type == Double.class) {
                double value = prefs.getDouble(key, field.get(target) == null ? 0.0 : field.getDouble(target));
                field.set(target, value);
            }
            else if (type == Color.class || type == Font.class) {
                Object value = deserializeObject(prefs.get(key, null));
                if (value != null) {
                    field.set(target, value);
                }
            }
        }
        catch (IllegalAccessException e) {
            throw new RuntimeException("Failed to set field: " + field.getName(), e);
        }
        catch (NumberFormatException e) {
            throw new RuntimeException("Invalid default value for field: " + field.getName(), e);
        }
    }

    /**
     * Saves the field value to preferences.
     *
     * @param field The field to save.
     * @param value Value to save
     */
    protected void saveField(Field field, Object value) {

        // Ignore non-setable fields
        if (Modifier.isStatic(field.getModifiers())) {
            return;
        }

        field.setAccessible(true);
        try {
            Class<?> type = field.getType();
            String key = field.getName();

            if (value == null && type != String.class) {
                prefs.remove(key);
            }
            else if (type == String.class) {
                prefs.put(key, value == null ? "" : field.get(this).toString());
            }
            else if (type == int.class || type == Integer.class) {
                prefs.putInt(key, (int)value);
            }
            else if (type == long.class || type == Long.class) {
                prefs.putLong(key, (long)value);
            }
            else if (type == boolean.class || type == Boolean.class) {
                prefs.putBoolean(key, (boolean)value);
            }
            else if (type == double.class || type == Double.class) {
                prefs.putDouble(key, (double)value);
            }
            else if (type == Color.class || type == Font.class) {
                String serialized = serializeObject((Serializable) value);
                prefs.put(key, serialized);
            }
        }
        catch (IllegalAccessException e) {
            throw new RuntimeException("Failed to save field: " + field.getName(), e);
        }
        catch (NumberFormatException e) {
            throw new RuntimeException("Invalid default value for field: " + field.getName(), e);
        }
    }

    /**
     * Serializes an object to a Base64 encoded string.
     *
     * @param obj The object to serialize.
     * @return The Base64 encoded string.
     */
    private static String serializeObject(Serializable obj) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ObjectOutputStream oos = new ObjectOutputStream(baos);
            oos.writeObject(obj);
            oos.close();
            return Base64.getEncoder().encodeToString(baos.toByteArray());
        }
        catch (IOException e) {
            throw new RuntimeException("Failed to serialize object", e);
        }
    }

    /**
     * Deserializes an object from a Base64 encoded string.
     *
     * @param str The Base64 encoded string.
     * @return The deserialized object.
     */
    private static Object deserializeObject(String str) {
        if (str == null || str.isEmpty()) {
            return null;
        }
        try {
            byte[] data = Base64.getDecoder().decode(str);
            ObjectInputStream ois = new ObjectInputStream(new ByteArrayInputStream(data));
            Object obj = ois.readObject();
            ois.close();
            return obj;
        }
        catch (IOException | ClassNotFoundException e) {
            log.error("Failed to deserialize object", e);
            return null;
        }
    }

    /**
     * Checks if a key exists in preferences.
     *
     * @param key The key to check.
     * @return True if the key exists, false otherwise.
     */
    public boolean keyExists(String key) {
        try {
            return Arrays.asList(prefs.keys()).contains(key);
        }
        catch (BackingStoreException e) {
            return false;
        }
    }

    /**
     * Creates a proxy instance of this class so that we can intercept method calls.
     *
     * @param clazz The class to create a proxy for.
     * @param prefs Preferences to save the values to/from.
     * @param autoSave Indicates whether to load existing values from storage upon creation and enable auto-saving on changes.
     *
     * @return A proxy instance of this class.
     */
    protected static <T> T createProxy(Class<T> clazz, Preferences prefs, boolean autoSave) throws Exception {
        T instance = new ByteBuddy()
                .subclass(clazz)
                .method(ElementMatchers.nameStartsWith("set"))
                .intercept(MethodDelegation.to(ChangeTrackingInterceptor.class))
                .make()
                .load(clazz.getClassLoader())
                .getLoaded()
                .getDeclaredConstructor()
                .newInstance();

        // Load from storage
        ((PersistanceManager) instance).prefs = prefs;
        ((PersistanceManager) instance).setAutoSave(autoSave);
        for (Field field : instance.getClass().getSuperclass().getDeclaredFields()) {
            ((PersistanceManager) instance).loadField(field);
        }
        return instance;
    }

    /**
     * Saves all fields to storage into the current instance.
     */
    public void saveToStorage() {
        for (Field field : getClass().getSuperclass().getDeclaredFields()) {
            try {
                this.saveField(field, field.get(this));
            }
            catch (IllegalAccessException e) {
                throw new RuntimeException("Failed to save field: " + field.getName(), e);
            }
        }
    }

    /**
     * SettingsManager class to handle persisting settings changes
     */
    public static class ChangeTrackingInterceptor {
        @RuntimeType
        public static Object intercept(@This Object self,
                                       @Origin Method method,
                                       @AllArguments Object[] args,
                                       @SuperCall Callable<?> zuper) throws Exception {

            // Set the field value first
            Object result = zuper.call();

            // If we are auto-saving, persist the change
            if (((PersistanceManager) self).isAutoSave()) {
                if (args != null && method.getName().startsWith("set")) {
                    String fieldName = method.getName().substring(3);

                    // Find the corresponding field name
                    for (Field field : self.getClass().getSuperclass().getDeclaredFields()) {
                        if (field.getName().equalsIgnoreCase(fieldName)) {
                            field.setAccessible(true);
                            ((PersistanceManager) self).saveField(field, args[0]);
                            break;
                        }
                    }
                    log.debug("Field changed: {}", fieldName);
                }
            }
            return result;
        }
    }

}
