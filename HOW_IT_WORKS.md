# How the Extension Works — Technical Flow

> Companion to [`MASTER_GUIDE.md`](MASTER_GUIDE.md). This file focuses **only** on
> what the shipped VS Code extension does at runtime — every command, every
> process, every file it touches.

---

## 1. High-level architecture

```
 ┌─────────────────────────── VS Code window ──────────────────────────┐
 │                                                                     │
 │   src/extension.ts          ← command registry + activation         │
 │     │                                                               │
 │     ├─ onboarding.ts        ← env config, az/python checks, MCP probe│
 │     ├─ chrome.ts            ← launch Chrome with --remote-debugging │
 │     ├─ formPaths.ts         ← multi-line paste → spawn python       │
 │     │                          (--validate-paths)                   │
 │     └─ extractor.ts         ← pick form(s) + env(s) → spawn python  │
 │                                (--backend mcp|playwright)           │
 │                                                                     │
 │   Settings (workspace)                                              │
 │     d365FormExtractor.environments  [{label,baseUrl,tenantId,mcpUrl}]│
 │     d365FormExtractor.formPaths     [{path,menuItem,validated}]     │
 │     d365FormExtractor.outputFolder                                  │
 │     d365FormExtractor.defaultBackend  mcp|playwright|ask            │
 │   SecretStorage                                                     │
 │     d365FormExtractor.mcpClientSecret   (optional, app-reg flow)    │
 └──────────────────────────────┬──────────────────────────────────────┘
                                │ spawn (stdio piped to Output channel)
                                ▼
                        python/extract.py
                ┌──────────────┴──────────────┐
                ▼                              ▼
        MCP backend                     Playwright backend
        (JSON-RPC over HTTPS)           (CDP → existing Chrome :9222)
                │                              │
                ▼                              ▼
        <baseUrl>/mcp                   logged-in D365 tab
        az access-token                 user already signed in
                │                              │
                └──────────────┬───────────────┘
                               ▼
                       openpyxl writer
                               ▼
                <workspace>/d365-extracts/*.xlsx
                (single-env extract or DeepCompare diff)
```

Everything in TypeScript is **orchestration only** — no D365 logic. All the
parsing, RPC, pagination, FastTab expansion, sweep, and Excel writing lives
in `python/extract.py`.

---

## 2. Activation & first-run

`activate(context)` in [src/extension.ts](src/extension.ts) runs on
`onStartupFinished`:

1. Creates one shared `OutputChannel` named **D365 FO Config Compare**.
2. If `globalState.welcomeShown` is false, schedules the walkthrough
   `prashant-verma.d365-form-extractor#d365FormExtractor.gettingStarted`
   to open after ~1.2 s, then flips the flag.
3. Registers all 7 commands:

| Command id                                | Handler              | Purpose                                   |
| ----------------------------------------- | -------------------- | ----------------------------------------- |
| `d365FormExtractor.welcome`               | `openWalkthrough`    | Re-open the walkthrough                   |
| `d365FormExtractor.scaffoldWorkspace`     | `scaffoldWorkspace`  | mkdir + `vscode.openFolder`               |
| `d365FormExtractor.onboard`               | `runOnboarding`      | Env config, Python deps, MCP probe        |
| `d365FormExtractor.startChromeCdp`        | `startChromeCdp`     | Launch Chrome on :9222                    |
| `d365FormExtractor.configureFormPaths`    | `configureFormPaths` | Validate UI paths → menu items            |
| `d365FormExtractor.extract`               | `runExtraction`      | Main path — pick form(s), env(s), backend |
| `d365FormExtractor.clearSettings`         | inline               | Wipe config + secret + onboarding flag    |

---

## 3. Command flows

### 3.1 `scaffoldWorkspace`
Asks for a folder (default `C:\D365-Form-Extractor`), `fs.mkdirSync` it +
`d365-extracts/` subdir, then `vscode.openFolder`. No D365 work.

### 3.2 `onboard` ([src/onboarding.ts](src/onboarding.ts))
Per-step:
1. **Python check** — `where python` (Win) / `which python`. If missing,
   prompts and aborts.
2. **`pip install`** — runs `python -m pip install --user openpyxl requests msal playwright`
   then `playwright install chromium`. Output streamed to channel.
3. **Azure CLI check** — `where az`. Warns if missing; not fatal (the user can
   still use Playwright backend).
4. **Environment loop** — for each env the user adds:
   - Label, base URL, optional tenant id.
   - Optional explicit `mcpUrl` (defaults to `${baseUrl}/mcp`).
   - **MCP probe**: HTTP `POST <mcpUrl>` with a stub `tools/list` JSON-RPC body
     and the user's `az account get-access-token --resource <baseUrl>` bearer.
     Result stored as `mcpAvailable: true|false` on the env entry.
5. Persists the array under `d365FormExtractor.environments` (Workspace scope).
6. Optionally collects an AAD app-reg (`tenantId`, `clientId`, secret). The
   secret is written via `context.secrets.store(...)`; everything else lives
   in `mcpAuth`. When both are present, the extract step injects them as
   `MCP_TENANT_ID` / `MCP_CLIENT_ID` / `MCP_CLIENT_SECRET` env vars to the
   Python child, and the Python side uses client-credentials instead of `az`.
7. Sets `globalState.onboarded = true`.

`ensureOnboarded` is a guard: if `onboarded` is false, prompts to run
onboarding first.

### 3.3 `startChromeCdp` ([src/chrome.ts](src/chrome.ts))
1. GET `http://localhost:9222/json/version` (2 s timeout). If it returns 200,
   bail out: CDP already running.
2. Resolve `chrome.exe`: setting → auto-detect (`Program Files`,
   `Program Files (x86)`, `LOCALAPPDATA`).
3. **Modal warning** — Chrome 136+ refuses `--remote-debugging-port` on the
   default user-data-dir, so we have to kill existing windows.
4. `taskkill /F /IM chrome.exe`, wait 1.5 s.
5. Build a dedicated profile dir:
   - `<workspace>/_chrome_profile` if there is a workspace, else
     `<tmp>/d365-chrome-profile`.
6. Spawn detached:
   ```
   chrome.exe --remote-debugging-port=9222 \
              --user-data-dir=<profile> \
              --no-first-run --no-default-browser-check
   ```
7. User signs in to D365 in that window. The profile persists, so subsequent
   launches skip sign-in.

### 3.4 `configureFormPaths` ([src/formPaths.ts](src/formPaths.ts))
1. Ask which environment to validate against.
2. Ask for legal entity (`cmp`), default `usmf`.
3. Open an **untitled scratch document** prefilled with comments + any
   previously-saved paths.
4. User edits — no save required.
5. Auto-call `startChromeCdp` (no-op if already up).
6. Spawn:
   ```
   python python/extract.py --validate-paths \
       --env <label>=<baseUrl> \
       --le <cmp> \
       --paths <path1>\x1f<path2>\x1f...
   ```
   (Unit-separator `\u001f` so paths can contain `>` and spaces.)
7. Read stdout line-by-line; parse `RESULT|<path>|<menuItem|>|<error|>`.
8. Merge into `formPaths` setting (Workspace scope) — replace by path.

### 3.5 `extract` ([src/extractor.ts](src/extractor.ts))
1. `ensureOnboarded` gate.
2. Ask for legal entity (e.g. `my30`).
3. Build the form list:
   - Saved `formPaths` with `validated && menuItem` → multi-select QuickPick.
   - Selection of N paths → N jobs (`{ form: menuItem, formPath: path }`).
   - If no validated paths, fall back to: free-text menu item + optional path.
4. Multi-select environments QuickPick (defaults all checked).
5. Backend selection:
   - Setting `defaultBackend` = `mcp|playwright|ask`.
   - If `ask`, show 2-item QuickPick.
6. **MCP pre-flight** (`ensureAzLogin`): for every selected env, run
   `az account get-access-token --resource <baseUrl>` (+ `--tenant <id>` when
   set). On failure, offers to open a terminal with
   `az login --scope <baseUrl>/.default --allow-no-subscriptions` and a Retry
   modal. Auto-captures the tenant id (`az account show`) if the env didn't
   have one.
7. Resolve output folder: setting → `<workspace>/d365-extracts` → `<tmp>`.
8. For each job, spawn:
   ```
   python python/extract.py \
       --backend <mcp|playwright> \
       --form <menuItem> \
       --le <cmp> \
       --out-dir <outFolder> \
       [--form-path <path>] \
       --env <label1>=<baseUrl1>|<tenant1>|<mcpUrl1> \
       --env <label2>=<baseUrl2>|<tenant2>|<mcpUrl2> \
       [--diff]                ← added when 2+ envs picked
   ```
   `--diff` is the signal to also write a `DeepCompare_*.xlsx`.
   Optional env vars `MCP_TENANT_ID|CLIENT_ID|CLIENT_SECRET` are injected if a
   client-credentials app reg is configured.
9. Stdout/stderr piped live into the Output channel.
   Cancellation token kills the child.
10. After all jobs: summary toast + "Open Folder" button.

---

## 4. The Python script — `python/extract.py`

`main()` parses argparse with three modes:
- `--validate-paths` → call `validate_form_paths(base_url, le, paths)`
- `--backend mcp`    → `extract_mcp(client, form, le)`
- `--backend playwright` → `extract_playwright(base_url, label, form, le)`
- For both extract modes, when `--diff` is set, writes a comparison Excel
  in addition to per-env files.

### 4.1 Authentication
- `az_token(resource, tenant)` — shells out to
  `az account get-access-token --resource <resource> [--tenant <tenant>]`.
  Returns the bearer string. Cached per-process via dict on `resource`.
- `_client_credentials_token(resource)` — if `MCP_TENANT_ID` /
  `MCP_CLIENT_ID` / `MCP_CLIENT_SECRET` are all present in env, does a
  direct `POST https://login.microsoftonline.com/<tenant>/oauth2/v2.0/token`
  with `client_credentials`. Used by automated/CI scenarios.

### 4.2 MCP client (`McpClient`)
- Single instance per env per run.
- `_rpc(method, params, rpc_id)` — POSTs JSON-RPC to `<mcpUrl>`:
  ```json
  {"jsonrpc":"2.0","id":<n>,"method":"tools/call",
   "arguments":{"name":"form_open_menu_item",...}}
  ```
- Headers: `Authorization: Bearer <token>`, `Content-Type: application/json`,
  `Accept: text/event-stream` (server streams SSE).
- Response handling: parse last `data: ` line, JSON-decode `result.content[0].text`.
- `tool_json(name, args)` — thin wrapper that returns the parsed JSON payload.

### 4.3 `extract_mcp` — the 5-stage deep extractor
```
Stage 1 — OPEN
  form_open_menu_item(menuItemName=<form>, legalEntity=<le>)
    → returns the whole control tree as one JSON blob.
  _parse_form() walks it once and collects:
      data = { fields:{...}, grids:{...}, tabs:{...}, fasttabs:{...} }

Stage 2 — EXPAND FASTTABS
  For each fasttab still recorded with state="Collapsed":
    form_click_control(controlName=<headerName>, actionId="Click")
  Walk again, fold in any newly-exposed nested controls.

Stage 3 — FORCE-OPEN TABS
  For every Tab/TabPage discovered:
    form_open_or_close_tab(tabName=<n>, tabAction="Open")
       ← tabAction is REQUIRED; "Activate" doesn't reveal lazy children.
  Re-parse to merge.

Stage 4 — PAGINATE GRIDS
  For every grid:
    while form_click_control(controlName=<grid>, actionId="LoadNextPage")
      returns more rows: keep appending.
    Deduplicate by row identity.

Stage 5 — 146-TERM SWEEP
  _FIND_CONTROLS_TERMS = a-z (26) + 120 high-signal keywords
                        (parameters, ledger, item, cust, vend, ...)
  For each term:
    form_find_controls(controlSearchTerm=<term>)
      → returns any control whose name/label CONTAINS <term>.
  Merge with the parsed tree — catches controls that live behind
  conditional renders, dynamic groups, hidden FastTabs, etc.

Finally:
  form_close_form(...)        ← polite cleanup
  return data
```

### 4.4 `extract_playwright` — CDP path
1. `chromium.connect_over_cdp("http://localhost:9222")`.
2. Pick (or open) a context targeting the env's base URL.
3. Navigate to `?cmp=<le>&mi=<form>`.
4. `page.evaluate(_JS_EXPAND_FASTTABS)` — clicks every
   `[class*="fastTab"] button[aria-expanded="false"]`.
5. For each visible tab strip → click + expand FastTabs again.
6. `_scroll_collect_grids(page)` — virtual-scroll grids: scroll until
   `scrollTop` stops advancing, dedupe rows by visible cell signature.
7. Returns the same shape as MCP (`fields/grids/tabs/fasttabs`).

### 4.5 `validate_form_paths`
1. Connect over CDP, open one tab.
2. For each path (e.g. `Sales and marketing > Setup > Parameters`):
   - Resolve module from segment 1.
   - Navigate, expand the sidenav, click each segment by accessible name.
   - Read final URL → extract `mi=<menuItem>`.
3. Print `RESULT|<path>|<menuItem|>|<error|>` per line.

### 4.6 Output
- `write_excel(data, env, ...)` → one sheet per area:
  `Fields`, `FastTabs`, `Tabs`, `Grids[name]` (one per grid).
- `write_diff_excel(per_env, ...)` → when 2+ envs were extracted in a single
  run:
  - Side-by-side sheets with `MATCH` (green) / `DIFF` (orange) flag column.
  - Grids matched by row PK heuristic, fields by control name.
  - Filename pattern `DeepCompare_<form>_<yyyymmdd_HHMMSS>.xlsx`.

---

## 5. State & where it lives

| What                              | Where                                              |
| --------------------------------- | -------------------------------------------------- |
| Env list (label, baseUrl, tenant) | Workspace settings → `d365FormExtractor.environments` |
| Validated form paths              | Workspace settings → `d365FormExtractor.formPaths`  |
| Output folder override            | Either scope → `d365FormExtractor.outputFolder`    |
| App-reg id/tenant (optional)      | Global settings → `d365FormExtractor.mcpAuth`      |
| App-reg **secret** (optional)     | `SecretStorage` key `d365FormExtractor.mcpClientSecret` |
| Onboarded flag                    | `globalState.onboarded`                            |
| Walkthrough-shown flag            | `globalState.welcomeShown`                         |
| Chrome profile (CDP)              | `<workspace>/_chrome_profile/` (or temp)           |
| Extracts (xlsx)                   | `<workspace>/d365-extracts/` by default            |

`clearSettings` removes all of the above except the Chrome profile and the
extract files (those are user data).

---

## 6. Process / IPC model

- One Node host process per VS Code window.
- One short-lived `python` child per extract job (and per validate run).
- TS ↔ Py contract is **stdout-only**:
  - extract.py prints freeform progress (forwarded verbatim).
  - validate-paths additionally prints structured `RESULT|...` lines.
  - exit code 0 = success, non-zero = failure.
- No file-based handoff, no temp JSON. The Excel files are the durable artifact.

---

## 7. End-to-end flow (the happy path)

```
User                  Extension (TS)              Python                 D365
 │ open VS Code ────► activate ──► walkthrough
 │ Run Onboarding ──► runOnboarding ──► pip install
 │                                    ──► az check
 │ add ENV1, ENV2 ──►  (probe /mcp via https) ◄──── 200 OK
 │ Start Chrome   ──► chrome.exe --remote-debugging-port=9222
 │ sign in to D365 inside that Chrome
 │ Configure Form Paths ─► open scratch doc
 │ paste 4 paths, click Validate
 │                  ─► spawn extract.py --validate-paths
 │                                       ─► CDP → resolve mi for each ─► Chrome
 │                  ◄── RESULT lines
 │                  ─► save formPaths
 │ Extract Form ──► multi-pick forms + envs + backend=mcp
 │                  ─► ensureAzLogin (per env)
 │                                       ─► spawn extract.py --backend mcp
 │                                          --env ENV1 --env ENV2 --diff
 │                                                   ─► McpClient ENV1 ─► /mcp
 │                                                       Stage1..5
 │                                                   ─► McpClient ENV2 ─► /mcp
 │                                                       Stage1..5
 │                                                   ─► write_excel x2
 │                                                   ─► write_diff_excel
 │                  ◄── exit 0
 │ ◄── "Extraction complete" toast → Open Folder
```

---

## 8. Failure modes & how the extension handles them

| Symptom                                          | Where it's caught                                          | What happens                                                  |
| ------------------------------------------------ | ---------------------------------------------------------- | ------------------------------------------------------------- |
| `python` not on PATH                             | `onboarding.which`                                         | Modal; abort                                                  |
| `pip install` fails                              | onboarding stream                                          | Surfaced in Output; user retries                              |
| `/mcp` returns 401                               | onboarding probe + `ensureAzLogin`                         | `mcpAvailable=false` → user nudged to Playwright              |
| `az account get-access-token` fails              | `ensureAzLogin`                                            | Offers `az login` in a terminal + Retry                       |
| Chrome 136 refuses default user-data-dir         | `startChromeCdp`                                           | Uses dedicated `_chrome_profile/`                             |
| Chrome already running on different port         | `startChromeCdp`                                           | Modal: close & relaunch                                       |
| Form path → wrong menu item                      | `validate_form_paths`                                      | `RESULT|...|error` recorded; row marked invalid               |
| MCP `tools/call` returns SSE error               | `McpClient._rpc`                                           | Raises; caller logs + non-zero exit                           |
| User cancels mid-extract                         | `withProgress` cancellation token                          | `child.kill()`                                                |
| One form fails in a batch                        | extractor.ts try/catch per job                             | `failCount++`; continue; summary lists failures               |
| App-reg client secret rotated                    | `getClientCredentialsToken`                                | Falls back to `az` token if env vars unset                    |

---

## 9. Build & ship

- TypeScript compiled by `esbuild` (see `scripts.build` in `package.json`)
  to `dist/extension.js`.
- Python script bundled as-is under `python/extract.py` (no build).
- Packaging: `npx --yes @vscode/vsce package --no-yarn` produces
  `d365-form-extractor-<version>.vsix`.
- Install: `code --install-extension <file>.vsix --force`.
- Beta status: `package.json.preview = true`, `displayName = "D365 FO Config
  Compare (Beta)"`, BETA ribbon baked into `media/icon-beta.png`.

---

## 10. File reference

| File                                         | Role                                       |
| -------------------------------------------- | ------------------------------------------ |
| [`package.json`](package.json)               | Manifest, commands, settings, walkthrough  |
| [`src/extension.ts`](src/extension.ts)       | Activation, command registry               |
| [`src/onboarding.ts`](src/onboarding.ts)     | Deps, env config, MCP probe, secret store  |
| [`src/chrome.ts`](src/chrome.ts)             | CDP launcher                               |
| [`src/formPaths.ts`](src/formPaths.ts)       | Path → menu-item validator UI              |
| [`src/extractor.ts`](src/extractor.ts)       | Extract orchestrator                       |
| [`src/welcome.ts`](src/welcome.ts)           | Walkthrough helpers                        |
| [`python/extract.py`](python/extract.py)     | Both backends + validator + Excel writer   |
| [`media/icon-beta.png`](media/icon-beta.png) | Marketplace icon                           |
| [`walkthrough/*.md`](walkthrough/)           | Walkthrough step bodies                    |
| [`archive/`](archive/)                       | Original Python pipeline + docs (frozen)   |

See [`MASTER_GUIDE.md`](MASTER_GUIDE.md) for the broader story (business
problem, gotchas, governance) and [`SKILL.md`](SKILL.md) for the
domain-knowledge cheat sheet.
