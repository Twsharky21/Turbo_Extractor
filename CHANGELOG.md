# Changelog

All notable changes to Turbo Extractor are documented here.

## v1.1.0 — 2026-02-23

### Added
- **Tooltips** on all editor and destination labels — hover any field label for a concise explanation of what it does and whether it's optional.
- **Edge-case test coverage** — 29 new tests covering rules on booleans, numeric coercion, keep-mode + rules combos, batch zero-row interleaving, destination management, and full-pipeline combinations.
- `main.py` entry point at project root.
- `README.md`, `LICENSE` (MIT), and `CHANGELOG.md`.

### Files added
- `gui/tooltip.py`
- `tests/test_tooltip.py` (6 tests)
- `tests/test_edge_gaps.py` (29 tests)
- `main.py`
- `LICENSE`
- `README.md`
- `CHANGELOG.md`
- `docs/` — screenshots

### Files modified
- `gui/ui_build.py` — tooltip imports and label wiring
- `gui/mixins/editor_mixin.py` — exported tooltip text constants for rules headers

---

## v1.0.0 — 2026-02-23

### Core extraction pipeline
- Load XLSX or CSV sources with configurable sheet selection.
- Row selection (ranges, sparse lists, or all rows).
- Rules-based row filtering: include/exclude with equals, contains, `<`, `>` operators. AND/OR combinators. Rules reference absolute source columns.
- Column selection (ranges, non-adjacent, or all columns).
- Two paste modes: **Pack Together** (dense, no gaps) and **Keep Format** (preserves column spacing within bounding box).
- Source start row offset for skipping headers.
- Destination placement: explicit row/column anchoring or automatic append mode.
- Collision detection on target columns only — gap columns from Keep Format never block. Enables automatic merge of overlapping bounding boxes.
- Batch execution with shared in-memory workbook cache, fail-fast on first error, and per-item disk saves for crash safety.

### GUI
- Tree-based project structure: Sources → Recipes → Sheets.
- Right-panel sheet editor with live model sync.
- Scrollable rules section with dynamic add/remove.
- Right-click context menus: rename recipes/sheets (inline edit), save/load/set/reset source templates.
- Animated throbber spinner during runs (threaded extraction keeps UI responsive).
- Rich run report dialog with color-coded results, monospace layout, and copy-to-clipboard.
- Move up/down for sources, recipes, and sheets.
- Destination browse with no overwrite confirmation prompt.
- Auto-capitalize column letter inputs.
- Title Case UI labels throughout.
- Custom window icon support.

### Templates
- Save/load source templates as portable JSON (excludes source file path).
- Default template: set once, auto-applied to every new source.
- Full round-trip fidelity for all SheetConfig fields including rules, paste mode, and destination settings.

### Autosave
- Debounced (1.2s) + periodic (45s) autosave to `~/.turbo_extractor_v3/autosave.json`.
- Atomic writes (write to `.tmp`, then `os.replace`) — no corrupt saves on crash.
- Configurable path via `TURBO_AUTOSAVE_PATH` environment variable.
- Auto-loads on startup when env var is set.

### Error handling
- Structured `AppError` with error codes: `DEST_BLOCKED`, `BAD_SPEC`, `SOURCE_READ_FAILED`, `SHEET_NOT_FOUND`, `INVALID_RULE`, `FILE_LOCKED`, `SAVE_FAILED`, `MISSING_DEST_PATH`, `MISSING_SOURCE_PATH`.
- `friendly_message()` produces plain-English one-liners for every error code. Never exposes raw tracebacks.
- Permission errors on source/destination files produce actionable "close the file" messages.

### Architecture
- Clean module separation: `core/` (pipeline logic, zero GUI dependency) and `gui/` (Tkinter UI).
- GUI split into focused mixins: `ReportMixin`, `TreeMixin`, `EditorMixin`, `ThrobberMixin`.
- `core/engine.py` re-export shim preserves backward compatibility after runner/batch split.
- Writer never writes `None` cells — prevents openpyxl phantom-cell bug that corrupts append scans.
- Planner operates on target columns only — documented in `core/landing.py` docstring.

### Testing
- 370+ tests across 16 test files.
- Combinatorial matrix tests (source type × paste mode × column spec × row spec × rules × destination config).
- Merge-mode, autosave, and template round-trip coverage.
- GUI wiring tests (Tkinter): tree operations, editor sync, report formatting, inline rename, reorder.
