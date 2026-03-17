# Copilot instructions for `stockticker`

## Project shape

This repository contains two separate desktop implementations of the same product:

- `java/`: the current Java 21 Swing application.
- `visualbasic/`: the legacy Visual Basic 6 Windows application.

The root `pom.xml` is only an aggregator for those two modules. For most work, run Maven against a specific module instead of the repo root.

At a high level, the Java implementation is a minimalist Java 21 Swing desktop application built with Maven.

Important existing guidance to preserve:

- `copilot-instructions.md`: Java files are expected to have Javadoc on all classes and methods except `@Override` members.
- `CONTRIBUTING.md`: follow existing style, keep VB6 headers/comments thorough, and do not introduce extra Windows dependencies such as OCX/OLE DLLs.

## Build, test, and lint commands

### Java module

Run commands from the repo root unless noted otherwise.

```bash
# compile the Java app
mvn -pl java compile

# run the committed Java test suite
mvn -pl java test

# run a single Java test class
mvn -pl java -Dtest=PriceTest test

# run a single Java test method
mvn -pl java -Dtest=PriceTest#testMethod test

# package the Java app, including shaded jar, Javadocs, and jpackage output for the current OS
mvn -pl java package

# generate Javadocs explicitly
mvn -pl java javadoc:javadoc
```

Notes:

- `mvn -pl java test` is currently valid but there are no committed tests under `java/src/test/java` yet.
- `java/pom.xml` does not configure a dedicated lint/static-analysis tool such as Checkstyle, PMD, or SpotBugs.
- `mvn -pl java package` runs the shade build and then `jpackage`; on macOS it produces a `.pkg`, on Windows it produces an `.exe`.

### Visual Basic 6 module

```bash
# validate/build the VB6 application on Windows with VB6 installed
mvn -pl visualbasic package
```

Notes:

- This build is Windows-only. `visualbasic/pom.xml` invokes `C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe`.
- The VB6 build rewrites version fields in `psmain.bas`, `psTicker.vbp`, and even the version line in `README.md` as part of the Maven pipeline. Treat those values as build-managed, not hand-maintained.
- There is no automated test or lint setup for the VB6 module in this repository.

## High-level architecture

### Big picture

The Java and VB6 modules are parallel implementations of the same stock ticker idea: a narrow desktop ticker UI backed by persisted settings, a stored symbol list, scheduled market-data refreshes, and optional summary views. They do not share code, but they do mirror the same domain concepts: settings, symbols/holdings, live prices, exchange rates, summary statistics, alarms/colors/fonts, and launch URLs.

### Java architecture

The main entry point is `java/src/main/java/com/pivotal/stockticker/App.java`, which sets FlatLaf and opens `TickerBar`.

`TickerBar` is the hub of the Java app:

- owns the core managers: `SettingsManager`, `SymbolsManager`, `PricesManager`, `ExchangeRatesManager`
- builds the scrolling ticker and summary UI
- reacts to `SettingsForm` and `SymbolsForm` changes through the `CallbackInterface`
- reloads state and redraws when settings or symbol data changes

The main layers are:

- `model/`: persisted state such as `Price`, `ExchangeRate`, `SymbolTransaction`, `SettingsManager`
- `service/`: orchestration and storage-aware managers such as `PricesManager`, `SymbolsManager`, `ExchangeRatesManager`, plus export/backup helpers
- `service/apis/`: provider adapters tried in sequence to fill missing market data
- `ui/` and `ui/components/`: Swing windows, dialogs, and custom controls

Data flow in practice:

1. `SymbolsManager` loads the stored holdings/symbol definitions.
2. `PricesManager` and `ExchangeRatesManager` derive the symbol/currency universe from those holdings.
3. API adapters fetch missing data and update persisted model objects.
4. `TickerBar` rebuilds the visible ticker, summary, and day-summary panels from `LivePrice.getLivePrices(...)`.

### Java persistence model

The most important non-obvious piece is `PersistanceManager`.

- Java state is stored in `java.util.prefs.Preferences` under `/stockticker/...`, not in JSON files.
- `PersistanceManager.createProxyInstance(...)` uses ByteBuddy to create autosaving proxies for model objects.
- Setter calls on proxied models trigger persistence through `ChangeTrackingInterceptor`.

That means persistence behavior depends on using the existing model objects and their setters. If you bypass that pattern, autosave will not happen.

### Java refresh model

`PricesManager` and `ExchangeRatesManager` both own schedulers. `TickerBar` starts or restarts them after initialization and after settings/symbol changes. When symbols change, `TickerBar` reloads symbols, then calls:

- `prices.replacePrices(symbols.getAllSymbolCodes(false))`
- `rates.replaceExchangeRates(symbols.getAllCurrencyCodes(false))`

Preserve that relationship when changing symbol-management behavior.

### Visual Basic 6 architecture

The VB6 entry point is `visualbasic/src/main/vb6/psmain.bas`, which defines version constants, registry keys, global helpers, timer plumbing, and startup logic. The main UI lives in `psmain.frm`.

The VB6 implementation is organized around:

- `.frm` forms for presentation (`psmain.frm`, `pssettings.frm`, `pssymbols.frm`, `frmAlarm.frm`, `frmTooltip.frm`, `pspreview.frm`)
- `.cls` classes for domain/infrastructure (`pssymbol.cls`, `psstock.cls`, `psRegistry.cls`, `psChart.cls`, `psRegion.cls`, `JsonBag.cls`)
- `.bas` modules for shared procedures and platform integration (`psmain.bas`, `psnetwork.bas`, `psdata.bas`, `psgdiplus.bas`, `psgen.bas`)

Persistence is registry-centric:

- `psmain.bas` defines the canonical registry key names.
- `psRegistry.cls` wraps the Win32 registry APIs and also handles export/import-style operations.
- `psmain.frm` reads registry-backed settings to drive colors, font, summary visibility, URLs, and scrolling behavior.

## Key conventions

### Java conventions

- Carry forward the existing Javadoc rule from the repo-level `copilot-instructions.md`: all classes and methods should have meaningful Javadoc unless the member is `@Override`.
- For every Java file, ensure all classes and methods (public, protected, package, and private) include Javadoc headers describing purpose, parameters, return values, and exceptions. Skip only members annotated with `@Override`.
- Keep Javadoc concise and meaningful, and make sure it is grammatically correct and free of spelling errors.
- Do not remove existing user-written comments or annotations while adding Javadoc.
- Replace incomplete or placeholder Javadoc with accurate descriptions based on the code's functionality.
- If a method overrides a superclass or interface member, do not add Javadoc unless it adds useful information beyond the inherited documentation.
- The repository uses Lombok (`@Slf4j`, getters/setters) heavily. Match the current style instead of expanding boilerplate manually.
- Persisted models follow the `PersistanceManager` proxy pattern. Prefer the existing factory/getter methods such as `SettingsManager.getInstance()` and `Price.getPrice(...)` over direct construction.
- `SymbolsManager` tracks new/modified/deleted holdings separately and writes them with `persistChanges()`. Do not replace that with ad hoc preference writes.
- UI updates are expected to return to the Swing thread via `SwingUtilities.invokeLater(...)`.

### Java UI work

- Treat `TickerBar` as the coordinator for Java UI changes. Settings and symbol edits flow back into it through `CallbackInterface.changed(...)`, then it reloads state and redraws the ticker/summary panels.
- The Java UI is not plain layout-manager-driven Swing. Custom widgets such as `SettingsLabel` and `ColouredTextPanel` delegate positioning/sizing helpers through `SettingsComponent`, and many dialogs size themselves from rendered content. Match that pattern instead of introducing a different layout approach in the middle of an existing screen.
- `ColouredTextPanel` is the core rendering primitive for ticker text and summaries. It owns cached painting, scrolling, contiguous background blocks, and text positioning. For ticker/summary visual changes, prefer extending `ColouredTextPanel` usage rather than replacing it with standard labels/tables.
- Preview/summary dialogs (`StockPanel`, `SummaryPanel`, `DaySummaryPanel`) are lightweight `JDialog`s with auto-hide `Timer`s that watch mouse position relative to both the dialog and the owning ticker region. Preserve that behavior when changing hover previews.
- Settings dialogs follow a consistent pattern: `loadFromSettings(...)` populates controls, `Utils.attachChangeListeners(...)` enables dirty-state tracking, `onOK()` writes back through `SettingsManager`, and the caller callback triggers redraw/reload work.
- `Utils.attachChangeListeners(...)` is the shared way to observe nested Swing form controls. Reuse it for new settings fields instead of wiring one-off listeners everywhere.
- UI colors, fonts, and summary visibility come from `SettingsManager`; use those values rather than hard-coded Swing defaults so the ticker, previews, and summaries stay visually consistent.
- When debugging UI geometry, prefer the existing `Utils.dumpAllBounds(...)` helper before introducing ad hoc debug code.

### Visual Basic 6 conventions

- Reuse the existing `REG_*` constants and `cRegistry`/`psRegistry.cls` access patterns rather than inventing new registry-key strings inline.
- Keep the current naming style: `ps*` modules/classes for core features and `frm*` for auxiliary forms.
- Do not hand-edit build-managed version values in `psmain.bas`, `psTicker.vbp`, or the README version line unless you are intentionally changing the Maven build process.
- Do not add new external Windows runtime dependencies; the existing code is designed to rely on standard Windows/VB6 capabilities.
