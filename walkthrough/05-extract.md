# Step 5 — Extract one or many forms

Run **Extract Form**:

1. Pick a **legal entity** (e.g. `usmf`, `my30`)
2. **Multi-select** one or more **forms** from your validated paths (Space to toggle, Enter to confirm)
3. Pick one or more **environments**
4. Pick a **backend**: D365 ERP MCP (headless, recommended) or Playwright (driven Chrome)

### What the MCP backend does (0.7.0+)

For each form, the deep extractor:

- Opens the form and parses the full control tree recursively
- Force-opens every Tab / FastTab so hidden ones populate
- Paginates every grid (up to 200 pages, deduped)
- Sweeps ~146 search terms via `form_find_controls` to find hidden fields

Per-tab progress is logged in the **D365 FO Configuration Comparator** output channel:

```
[3/7] Vendors: 45 fields, 1 grid(s)
Sweep added 12 new field(s) (total fields: 218).
```

### Output

- `<form>_<le>_<env>_<timestamp>.xlsx` per environment
- `<form>_<le>_DIFF_<timestamp>.xlsx` when 2+ environments selected (green = match, pink = differ, yellow = missing on one side)

Files land in `<workspace>/d365-extracts/` by default.
