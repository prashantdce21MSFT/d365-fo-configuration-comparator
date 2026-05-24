# Changelog

## 0.8.4 — 2026-05-24

- **Disclaimer surfaced inside the extension**: status-bar warning indicator (always visible), activation toast (until acknowledged), Output channel banner (auto-opens on first run), new `D365 FO Config Compare: Show Disclaimer` command, and disclaimer block in walkthrough Step 1. Makes it explicit this is an independent community tool — not a Microsoft product, provided AS IS, not for production.
- **README**: prominent disclaimer banner at the top.
- **Packaging**: excluded `Test/**` from the published `.vsix` (was shipping ~40 sample xlsx files).
- **Publisher ID**: fixed casing to `PrashantVerma` to match the Marketplace account.

## 0.8.3 — 2026-05-18

- **Playwright — posting-matrix forms**: new `_JS_POSTING_TYPES` loop iterates the left-side radio list on forms like `InventPosting`, capturing every posting type's matrix (verified: 59 posting types, 103+ rows on `InventPosting`).
- **Playwright — grid column alignment**: matrix cells now align to headers by position (with empty structural columns dropped) so the `Main account` segmented column populates correctly.
- **Playwright — per-tab Excel sheets** for matrix forms with a leading `Posting type` column + auto-filter.
- **Playwright — master-detail iteration**: new `--per-row-detail` flag (default cap `--max-records 500`) clicks each row in the top grid and captures the detail panel into `Records` + `Record grids` sheets.
- **Fields sheet readability**: excludes grid-cell inputs, radio widgets, and page-chrome search boxes; adds `Label` (UI text) and `Field (internal)` columns; better section detection; deduped top-fields vs first-tab fields; auto-filter enabled.
- **Form-path validator**: fixed missing-first-character bug by switching to `press_sequentially()` with focus settle.
- **Backend advisor**: `validate_form_paths` now emits an `ADVICE|<path>|<backend>|<flags>|<reason>` line per validated form. The Extract picker pre-selects the recommended backend and auto-passes `--per-row-detail` when advised. Recommendations persist in `formPaths` settings.
- **Empty posting-type placeholders**: posting types with 0 rows are still listed in the per-tab sheet so consultants can see the full radio list.

## 0.7.0 — 2026-05-13

- **Deep MCP extraction**: recursively walks every Tab / FastTab / Group / ReferenceGroup, merging nested fields and grids. Per-tab progress is logged.
- **Force-opens every tab** via `form_open_or_close_tab` (falls back to `form_click_control`) so previously-empty tabs now populate.
- **Full grid pagination** for every grid discovered at any nesting level (up to 200 pages, deduped).
- **Hidden-control sweep**: runs `form_find_controls` across ~146 alphanumeric + domain terms (Code, Name, Account, Vend, Cust, Item, …) to surface fields the tab walk missed. Output shows `Sweep added N new field(s)`.
- README rewrite and walkthrough polish.

## 0.6.x — 2026-05-13

- Multi-select form extraction in the Extract command.
- Form-path validator hardened: notification dismiss, magnifier-button fallback, breadcrumb-filtered result picker, MI-change detection, ASCII-safe output.
- Per-environment 4-tuple unpack fix.

## 0.5.1 — 2026-05-13

- Onboarding now actively drives `az login` per environment (tries default tenant first; prompts for tenant only on failure).
- Onboarding probes each environment's `/mcp` endpoint and records `mcpAvailable`.
- If no environment supports MCP, `defaultBackend` is auto-set to `playwright`.
- Onboarding offers inline form-path configuration + Playwright validation as a final step (no more typing menu items per extraction).
- Onboarding is idempotent: re-running it lets you keep existing environments/form paths or extend them.

## 0.1.0 — 2026-05-12

- Initial release.
- Two extraction backends: D365 ERP MCP (via Azure CLI auth) and Playwright (Chrome CDP).
- Onboarding wizard.
- Single-form extraction with optional cross-environment diff Excel.
- Start-Chrome-with-CDP command.
