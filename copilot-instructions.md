## Architecture
This project is a minimalist Java application using standard Swing.

The targeted java version is 21, and the plugin is built using the Maven build system.

## Coding Style
- For every Java file, ensure **all classes and methods (public, protected, package, and private)** include Javadoc headers describing purpose, parameters, return values, and exceptions. Skip only members annotated with `@Override`.
- Keep Javadoc concise and meaningful; document each parameter, return type, and any thrown exceptions.
- Do not remove existing user-written comments or annotations while adding Javadoc.
- Replace incomplete or placeholder Javadoc with accurate descriptions based on the code's functionality.
- If a method is overridden from a superclass or interface, do not add Javadoc unless it provides additional information beyond the inherited documentation.
- Ensure that Javadoc is grammatically correct and free of spelling errors, as it will be used for generating documentation and improving code readability.
