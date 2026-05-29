# SKILL.md — D365 FO Configuration Comparator

> **Domain knowledge for working on / with this VS Code extension.** Covers the business problem, the two extraction backends (MCP + Playwright), their internals, and the lessons learned shipping it.

---

## 1. Business Problem

Acme REGION (and many other D365 F&O programs) hit a recurring class of bug during SIT/UAT:

> *"It worked in DEV. Now in UAT it doesn't. Why?"*

The answer is almost always **configuration drift** — some form, parameter, or master-data record differs between the **Golden** environment and **UAT / SIT / PROD-like** environments. These mismatches:

- cause hard-to-diagnose functional bugs late in the cycle,
- consume multiple consultant-days per incident,
- and are invisible to backend diff tools because **the difference is at the form/parameter level, not the entity level.**

This extension's job is to make that drift **visible, side-by-side, in Excel, mapped to the UI form a functional consultant actually understands** — *before* test cycles begin.

### Output is for functional consultants, not developers

- Sheets are named after forms and tabs (e.g. `General`, `Vendors`, `VendGroupGrid`) — not technical entities.
- Field labels are shown, not just internal names.
- A `_DIFF.xlsx` workbook highlights mismatches (green = match, pink = differ, yellow = missing on one side).

---

## 2. Two Backends — Why Both Exist

| | **D365 ERP MCP** | **Playwright (Chrome CDP)** |
|---|---|---|
| **How** | JSON-RPC over HTTPS to `<env>/mcp` | Drives logged-in Chrome on `:9222` |
| **Auth** | Azure CLI bearer token (`az account get-access-token`) | Existing Chrome session |
| **Mode** | Headless | Visible browser, user can watch |
| **Best for** | Structured, form-aware, explainable extraction | Forms MCP can't open; visual verification; complex/segmented UI |
| **Output shape** | Tab-wise sheets aligned to F&O form model | Same shape, sourced from DOM |
| **Speed** | Slower per form (many MCP calls + sweep) | Faster on complex forms |
| **Risk** | Read-only by protocol design | UI-level — can accidentally trigger actions; must guard against Create/Edit/Delete |
| **Coverage gaps** | Misses segmented controls (financial dimensions), 25-row default page, ~160 KB response cap, limited fuzzy search | None significant; depends on visible DOM |

**Verdict:** Neither is universally better. The extension lets the user **pick at run-time per extraction** because:

- For master-data forms with simple tabs → MCP is enough and gives the cleanest output.
- For complex forms (parameters with dozens of FastTabs, financial-dimension segmented entries, etc.) → Playwright caught more diffs in real scenarios (e.g. 8 vs 5).

### Why not just compare Data Entities?

- Data entities are efficient and scale to thousands of rows.
- **But** they produce hundreds of technical columns that functional consultants can't map back to a form.
- Field-to-UI mapping is the whole value proposition. Hence: form-based extraction is non-optional.

### Governance

- AI is used **only for read/compare/extract**. Never for write.
- Humans review the DIFF Excel, decide the correct value, and apply fixes through controlled channels (ADO tickets, deployment pipelines, manual config).

---

## 3. MCP Backend — Internals

### 3.1 Authentication
```
az account get-access-token --resource https://<env>.dynamics.com
```
No app registration, no secret stored. Every run asks `az` fresh. The token is passed as `Authorization: Bearer <token>` to `<env>/mcp`.

### 3.2 MCP tools used

| Tool | Purpose |
|---|---|
| `form_open_menu_item` | Open a form by `mi=...` (the menu-item key from the URL) |
| `form_open_or_close_tab` | Force-open a Tab / FastTab (more reliable than clicking) |
| `form_click_control` | Fallback to activate a tab; also used with `actionId=LoadNextPage` for grid pagination |
| `form_find_controls` | Search controls by term — used for the sweep |
| `form_select_grid_row` | (Not yet ported) Activates per-row detail extraction |
| `form_close_form` | Cleanup |

### 3.3 The 5-stage extraction pipeline (`extract_mcp` in `python/extract.py`)

```
[1] Open form + initial recursive parse
       └── _parse_form walks Tab → TabPage → Group → FastTab → ReferenceGroup → Children
       └── Merges nested fields & grids into a top-level flat dict
       └── While preserving per-tab structure for sheet naming

[2] Expand every collapsed FastTab
       └── _expand_fasttabs(client, fasttabs) — click each

[3] Force-open every Tab
       └── form_open_or_close_tab (preferred) → form_click_control (fallback)
       └── Re-parse the active payload, attribute the new fields to that tab
       └── Subtract tab-attributed fields from "top" to avoid duplicates

[4] Paginate every grid
       └── _paginate_grid uses _find_grid_recursive (handles deeply nested grids)
       └── form_click_control(controlName=<grid>, actionId="LoadNextPage")
       └── Up to 200 pages, dedup via tuple keys of row values
       └── Triggered when grid reports has_next OR ≥25 rows (default page)

[5] Hidden-control sweep
       └── _find_controls_sweep iterates ~146 terms:
            a-z, 0-9, plus ~100 domain terms (Code, Name, Date, Account, Vend,
            Cust, Item, Posting, Default, Tax, Currency, ...)
       └── Adds any returned non-grid controls missing from steps 1-3
       └── Skips controls named SystemDefined*
       └── Output: "Sweep added N new field(s)"
```

### 3.4 Constraints to remember

- **~160 KB response cap** — keep responses small, request specific tabs not the whole form when possible.
- **25-row default grid page** — always assume pagination is needed.
- **`form_select_grid_row` requires `rowNumber` as a STRING** (a notorious foot-gun).
- **Sweep is the main perf cost** — running 146 `form_find_controls` calls adds latency. Worth it for completeness; consider trimming for time-critical batch runs.
- **Batch in chunks of ~10 forms**, not 100 — sessions/tokens get flaky on long runs.

### 3.5 What MCP still misses
- **Segmented controls** (financial dimensions multi-segment editors) — values often blank.
- **Hidden / conditionally-shown sections** beyond what the sweep finds.
- Forms not exposed by the MCP server at all → fall back to Playwright.

---

## 4. Playwright Backend — Internals

### 4.1 Connection

```python
playwright.chromium.connect_over_cdp("http://localhost:9222")
```

User runs `D365 FO Configuration Comparator: Start Chrome with Remote Debugging`, which:
1. Closes any running Chrome.
2. Relaunches with `--remote-debugging-port=9222 --user-data-dir=<profile>`.
3. User signs in to D365 once in that Chrome. Playwright attaches to the same session.

### 4.2 Form open
- Navigate to `<env>/?cmp=<le>&mi=<menuItem>`.
- Wait for the form chrome to render.

### 4.3 What gets scraped

- **Tabs**: enumerated via `[role="tab"]` and the F&O FastTab CSS class set.
- **Fields**: `_JS_TABS` evaluator pulls input/combobox/checkbox label + value from each tab's panel.
- **Grids**: `_scroll_collect_grids` virtual-scrolls each grid container, collecting row cells as it goes. This bypasses the MCP 25-row cap entirely.
- **Action buttons** (`_ACTION_BUTTONS`): a controlled list — clicking is restricted to known-safe activators (tab expanders, grid scroll triggers). **Never** Save / New / Delete.

### 4.4 Safety guards

- Hard allowlist of clickable selectors.
- All interactions are read-style (focus, hover, scroll).
- The extension's prompt to the extractor explicitly forbids Create/Edit/Delete actions.

### 4.5 When Playwright wins

- Forms where MCP returns empty tabs.
- Forms with segmented controls — DOM has the rendered text even when MCP doesn't.
- When user wants to *see* the extraction happening.

---

## 5. Form-Path Validator

Users provide UI paths in human-readable form:
```
Accounts payable > Setup > Vendor groups
```
The validator (`validate_form_paths` in `python/extract.py`) discovers the corresponding `mi=` menu item by:

1. Navigate Chrome (CDP) to D365 home.
2. Dismiss any notification flyout (was poisoning the result list).
3. Click the magnifier (Search) button — fallback chain handles different aria-labels.
4. Type the leaf segment of the path.
5. Wait for results, **filter by breadcrumb** matching the parent segments.
6. Click the match, detect URL change, capture `mi=...`.
7. Persist `{path, menuItem, validated: true}` to `d365FormExtractor.formPaths` in `.vscode/settings.json`.

Robustness fixes shipped along the way:
- Notification dismiss before each path (otherwise the listitem selector picked toast items).
- Magnifier button fallback (`aria-label="Search"` vs `"Search for a page"`, `is_visible()` lying on icon buttons).
- MI-change detection (some clicks don't trigger SPA nav → result was being captured as `DefaultDashboard`).
- ASCII-only output (Windows console cp1252 choked on `✓`).
- 4-tuple env unpack fix.

---

## 6. Output Layout

For each `<form>` × `<environment>`:

```
<form>_<le>_<env>_<timestamp>.xlsx
   Form Info          caption, company, form name, timestamp
   Top Fields         name, label, value, type, mandatory, visible
   <TabName>          fields scoped to that tab
   <GridName>         paginated grid rows
```

When 2+ environments are selected:

```
<form>_<le>_DIFF_<timestamp>.xlsx
   Green   = match
   Pink    = differ
   Yellow  = present in one env, missing in another
```

---

## 7. Architecture

```
VS Code Extension (TypeScript)
├── src/extension.ts       — command registration, walkthrough wiring
├── src/extractor.ts       — multi-select picker, env picker, backend picker, spawns Python
├── src/onboarding.ts      — Python deps install, az login probe, env registration, /mcp probe
├── src/validator.ts       — orchestrates form-path validation
└── python/extract.py      — the actual extractor (single file, two backends)
        ├── MCP path:       extract_mcp + _parse_form + _paginate_grid + _find_controls_sweep
        ├── Playwright path: extract_playwright + _JS_TABS + _scroll_collect_grids + _ACTION_BUTTONS
        ├── Validator:      validate_form_paths
        └── Excel writer:   write_excel + write_diff
```

Python deps installed by onboarding: `requests`, `msal`, `openpyxl`, `playwright`.

---

## 8. Key Takeaways (the lessons we paid for)

1. **Functional consultants need form-shaped diffs, not entity-shaped diffs.** Data Entity comparison is technically possible but operationally useless for this audience.
2. **Neither MCP nor Playwright is universally better — let the user choose per form.**
3. **MCP's biggest hidden cost is the sweep.** It's necessary for completeness; budget for the calls.
4. **Tabs don't open by themselves.** `form_open_or_close_tab` per tab is mandatory or you get empty tabs.
5. **Every grid is paginated.** Always. Default page = 25. Plan for 200 pages.
6. **`form_select_grid_row` needs a string `rowNumber`.** Pass an int and you'll waste an afternoon.
7. **Playwright is a power tool — guard it.** Allowlist clicks. Forbid Create/Edit/Delete in the extractor prompt.
8. **Batch ~10 forms at a time**, not 100. Sessions get flaky; tokens expire; retries are cheaper.
9. **Path validation is fragile** — F&O search UX has many states (notifications, breadcrumbs, SPA navigation quirks). Validate once, cache the `mi=`, reuse forever.
10. **AI for read; humans for write.** All extraction is read-only. Fixes flow through ADO/manual change control. This boundary is the whole governance model.

---

## 9. Roadmap (not yet shipped)

- Port `_extract_row_details` from the original standalone extractor — per-row tab walk via `form_select_grid_row`. ~1-3 min/form but captures sub-grid detail forms.
- Smarter sweep — narrow the term set per form type to cut MCP latency.
- Optional: per-environment auth caching to skip `az` round-trip on consecutive forms.
- Optional: HTML diff report in addition to Excel.
