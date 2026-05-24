# D365 FO Config Compare

> 🚧 **Beta** — actively developed. Expect rapid iteration; please report issues.

> ⚠️ **Disclaimer — use at your own risk.** This is an independent, community-built tool. It is **not a standard out-of-the-box Microsoft product**, not affiliated with, endorsed by, or supported by Microsoft. The software is provided **AS IS** under the MIT license, with no warranty of any kind. **Not recommended for production environments** — validate against UAT or a read-replica first. "Microsoft", "Dynamics 365", and related names are trademarks of Microsoft Corporation, used here only for descriptive purposes.

Extract **Microsoft Dynamics 365 Finance & Operations** form configurations from one or more environments and compare them side-by-side in Excel — straight from VS Code.

> 📘 **Read the [Master Guide](MASTER_GUIDE.md)** for the complete story: business problem, full architecture, both backends in deep detail, every gotcha we hit, and the original Python pipeline (preserved in [`archive/`](archive/)).

---

## What it does

1. You point it at one or more D365 F&O environments.
2. You pick one or more forms (e.g. `smmParameters`, `VendGroup`).
3. It walks **every tab, every FastTab, every grid (with pagination)**, sweeps for hidden controls, and writes one Excel per environment.
4. If you picked 2+ environments, it also produces a **`_DIFF.xlsx`** with field-level match / mismatch highlighting.

## Two backends — pick at run-time

| Backend | When to use | Auth |
|---|---|---|
| **D365 ERP MCP** *(recommended, headless, fast)* | Forms exposed by the F&O MCP server. Deep extraction: all tabs + grids paginated + hidden-control sweep. | Azure CLI bearer token (`az account get-access-token`) |
| **Playwright (Chrome CDP)** | Visual verification, or forms MCP can't open. Drives your logged-in Chrome on port 9222. | Your existing Chrome session |

---

## Quick start

1. Install the `.vsix`.
2. **Command Palette → `D365: Getting Started`** — opens the walkthrough.
3. Run the walkthrough steps in order:
   - **Create / open a workspace folder** (e.g. `C:\D365-Form-Extractor`)
   - **Run Onboarding** — installs Python deps, checks `az login`, registers environments, probes `/mcp`
   - **Start Chrome with Remote Debugging** *(only needed for Playwright backend or path validation)*
   - **Configure Form Paths** — paste UI paths like `Accounts payable > Setup > Vendor groups`; the validator finds the menu item (`mi=VendGroup`)
   - **Extract Form** — pick one or many forms, pick environments, pick backend, watch the Output panel
4. Excels land in `<workspace>/d365-extracts/`.

---

## Commands (Ctrl+Shift+P → "D365") — run in this order

| # | Command | What it does |
|---|---|---|
| 1 | `D365 FO Config Compare: Getting Started` | Opens the step-by-step walkthrough |
| 2 | `D365 FO Config Compare: Create / Open Workspace Folder` | Scaffolds `.vscode/settings.json` |
| 3 | `D365 FO Config Compare: Run Onboarding` | Python deps + `az login` + environment registration + MCP probe |
| 4 | `D365 FO Config Compare: Start Chrome with Remote Debugging` | Closes Chrome, relaunches with `--remote-debugging-port=9222` |
| 5 | `D365 FO Config Compare: Configure Form Paths` | Paste UI paths; auto-discovers `mi=...` |
| 6 | `D365 FO Config Compare: Extract Form` | Multi-select forms + environments, then **pick a backend at run-time: D365 ERP MCP (headless, recommended) or Playwright (Chrome CDP)** |
| — | `D365 FO Config Compare: Clear Saved Settings` | Reset everything (use only to start over) |

---

## What you get in the output

For each form + environment:

```
<form>_<env>.xlsx
  Sheet: Form Info        caption, company, form name
  Sheet: Top Fields       name, label, value, type, mandatory, visible
  Sheet: <TabName>        fields scoped to that tab
  Sheet: <GridName>       grid rows with all columns
```

When 2+ environments are picked, an extra workbook:

```
<form>_DIFF_<timestamp>.xlsx     OK rows in green, DIFF rows highlighted, missing fields flagged per env
```

---

## Prerequisites

- Windows 10/11
- Python 3.10+
- Azure CLI (`az`) signed in with read access to the D365 environment(s)
- Google Chrome *(only for Playwright backend or path validation)*

Onboarding installs the Python deps for you: `requests`, `msal`, `openpyxl`, `playwright`.

---

## Authentication

```
az account get-access-token --resource https://<env>.dynamics.com
```

That's it. The extension never stores tokens or secrets; every run asks `az` fresh.

---

## Settings reference

| Setting | Purpose |
|---|---|
| `d365FormExtractor.pythonPath` | Python executable (default: `python` on PATH) |
| `d365FormExtractor.chromePath` | `chrome.exe` path (auto-detected) |
| `d365FormExtractor.environments` | `[{label, baseUrl, tenantId?, mcpUrl?, mcpAvailable?}]` |
| `d365FormExtractor.formPaths` | `[{path, menuItem, validated, error?}]` — populated by the validator |
| `d365FormExtractor.outputFolder` | Where Excels are written (default `<workspace>/d365-extracts`) |
| `d365FormExtractor.defaultBackend` | `mcp` \| `playwright` \| `ask` |

---

## Tips

- **First time?** Run `D365: Getting Started` — the walkthrough wires everything in order.
- **MCP "0 fields on some tabs"?** 0.7.0+ opens every tab, paginates every grid, and sweeps for hidden controls. Watch the Output panel for per-tab counts and the `Sweep added N field(s)` summary.
- **Path validation fails?** Make sure Chrome on `:9222` is signed in to D365 and on the dashboard. Re-run `Configure Form Paths`.
- **Want to compare envs?** Pick 2+ in the environment quick-pick during Extract — the DIFF workbook is generated automatically.

## License

MIT
