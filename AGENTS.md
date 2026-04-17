# AGENTS.md

## StockTicker AI Agent Guide

This document summarizes essential knowledge for AI agents working productively in the StockTicker codebase. It covers architecture, workflows, and project-specific conventions that differ from typical open source projects.

---

### 1. Big Picture Architecture

- **Dual Implementation:**
  - `java/`: Modern Java 21 Swing desktop app (actively developed).
  - `visualbasic/`: Legacy Visual Basic 6 Windows app (maintenance only).
  - Both implement the same stock ticker concept: a narrow desktop ticker UI, persisted settings, symbol list, scheduled market-data refresh, and summary views. No code is shared, but domain concepts are mirrored.

- **Java Structure:**
  - **Entry Point:** `java/src/main/java/com/pivotal/stockticker/App.java` (sets FlatLaf, opens `TickerBar`).
  - **Core Hub:** `TickerBar` owns managers (`SettingsManager`, `SymbolsManager`, `PricesManager`, `ExchangeRatesManager`), builds UI, and coordinates state reload/redraw.
  - **Layers:**
    - `model/`: persisted state (e.g., `Price`, `ExchangeRate`, `SettingsManager`)
    - `service/`: orchestration, storage-aware managers, export/backup
    - `service/apis/`: provider adapters for market data
    - `ui/`, `ui/components/`: Swing windows, dialogs, custom controls

- **Persistence:**
  - Java state is stored in `java.util.prefs.Preferences` under `/stockticker/...`.
  - `PersistanceManager.createProxyInstance(...)` (uses ByteBuddy) creates autosaving proxies; setters trigger persistence.
  - **Important:** Always use provided model objects and their setters for state changes—bypassing this disables autosave.

- **Refresh Model:**
  - `PricesManager` and `ExchangeRatesManager` own schedulers. `TickerBar` restarts them after settings/symbol changes.
  - Symbol changes: `prices.replacePrices(symbols.getAllSymbolCodes(false))`, `rates.replaceExchangeRates(symbols.getAllCurrencyCodes(false))`.

- **VB6 Structure:**
  - Entry: `visualbasic/src/main/vb6/psmain.bas` (version constants, registry keys, startup logic).
  - UI: `.frm` forms; domain: `.cls` classes; shared/platform: `.bas` modules.
  - Persistence is registry-centric via `psRegistry.cls`.

---

### 2. Developer Workflows

- **Java Build/Test:**
  - Compile: `mvn -pl java compile`
  - Test: `mvn -pl java test` (no committed tests yet)
  - Single test class: `mvn -pl java -Dtest=PriceTest test`
  - Package (shaded jar, Javadocs, jpackage): `mvn -pl java package`
  - Javadocs: `mvn -pl java javadoc:javadoc`
  - **Note:** Run Maven against the module, not the repo root.

- **VB6 Build:**
  - Windows-only: `mvn -pl visualbasic package` (invokes VB6, rewrites version fields in several files)

---

### 3. Project-Specific Conventions

- **Java:**
  - All classes/methods require meaningful Javadoc (except `@Override`).
  - Use Lombok for boilerplate (`@Slf4j`, getters/setters).
  - Persisted models: use factory/getter methods (e.g., `SettingsManager.getInstance()`, `Price.getPrice(...)`).
  - `SymbolsManager` tracks changes and persists with `persistChanges()`.
  - UI updates must return to the Swing thread (`SwingUtilities.invokeLater(...)`).
  - Custom widgets (e.g., `ColouredTextPanel`) are preferred for ticker/summary rendering.
  - Use `Utils.attachChangeListeners(...)` for Swing form dirty-state tracking.
  - UI colors/fonts/visibility come from `SettingsManager` (not hard-coded).

- **VB6:**
  - Reuse `REG_*` constants and registry access patterns.
  - Naming: `ps*` for core modules/classes, `frm*` for forms.
  - Do not hand-edit build-managed version fields.
  - No new external Windows dependencies.

---

### 4. Integration Points & External Dependencies

- **Java:**
  - Uses ByteBuddy for persistence proxies.
  - Market data fetched via `service/apis/` adapters.
  - No dedicated static analysis tool configured.

- **VB6:**
  - Registry is the main persistence mechanism.
  - Build process rewrites version fields automatically.

---

### 5. Key Files & Directories

- `.github/copilot-instructions.md`: Project conventions and architecture.
- `README.md`: Project overview and philosophy.
- `java/src/main/java/com/pivotal/stockticker/`: Main Java sources.
- `visualbasic/src/main/vb6/`: Main VB6 sources.
- `CONTRIBUTING.md`: Style and contribution rules.

---

**For more details, see `.github/copilot-instructions.md` and the WIKI.**

