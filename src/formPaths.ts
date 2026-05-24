import * as vscode from "vscode";
import * as path from "path";
import { spawn } from "child_process";
import { startChromeCdp } from "./chrome";

export interface FormPathEntry {
    path: string;
    menuItem?: string;
    validated?: boolean;
    error?: string;
    /** Backend recommendation derived from a Playwright DOM probe during validation. */
    recommendedBackend?: "mcp" | "playwright";
    /** Comma-joined flag list, e.g. "per-row-detail". Empty when no extras. */
    recommendedFlags?: string;
    /** Human-readable rationale, e.g. "Posting matrix \u2014 11 radio-driven sub-grids". */
    recommendedReason?: string;
}

export async function configureFormPaths(context: vscode.ExtensionContext, out: vscode.OutputChannel): Promise<void> {
    const cfg = vscode.workspace.getConfiguration("d365FormExtractor");
    const envs = cfg.get<{ label: string; baseUrl: string }[]>("environments") || [];
    if (envs.length === 0) {
        vscode.window.showErrorMessage("Configure environments first (Run Onboarding).");
        return;
    }

    // Pick env to validate against
    const envPick = await vscode.window.showQuickPick(
        envs.map(e => ({ label: e.label, description: e.baseUrl, env: e })),
        { title: "Validate form paths against which environment?", ignoreFocusOut: true }
    );
    if (!envPick) return;

    // Optional legal entity
    const le = (await vscode.window.showInputBox({
        title: "Legal entity (cmp) for navigation",
        value: "usmf",
        ignoreFocusOut: true,
    }))?.trim() || "usmf";

    // Open scratch doc for user to paste paths
    const existing = cfg.get<FormPathEntry[]>("formPaths") || [];
    const placeholder =
        "# Paste one form path per line. Lines starting with # are ignored.\n" +
        "# Example:\n" +
        "# Production control > Setup > Production journal names\n" +
        "# Production control > Setup > Production > Production pools\n\n" +
        existing.map(e => e.path).join("\n");
    const doc = await vscode.workspace.openTextDocument({ content: placeholder, language: "plaintext" });
    await vscode.window.showTextDocument(doc, { preview: false });

    const choice = await vscode.window.showInformationMessage(
        "Edit the form paths in the open document (no need to save), then click Validate. Chrome with CDP will start automatically if not already running.",
        { modal: false },
        "Validate", "Cancel"
    );
    if (choice !== "Validate") return;

    const lines = doc.getText().split(/\r?\n/).map(l => l.trim()).filter(l => l && !l.startsWith("#"));
    if (lines.length === 0) { vscode.window.showWarningMessage("No form paths to validate."); return; }

    // Auto-launch Chrome CDP if needed (no-op when already running).
    out.show(true);
    await startChromeCdp(out);

    out.show(true);
    out.appendLine(`\n=== Validating ${lines.length} form path(s) against ${envPick.env.label} ===`);

    const pythonPath = cfg.get<string>("pythonPath") || "python";
    const scriptPath = path.join(context.extensionPath, "python", "extract.py");
    const args = [
        scriptPath, "--validate-paths",
        "--env", `${envPick.env.label}=${envPick.env.baseUrl}`,
        "--le", le,
        "--paths", lines.join("\u001f"), // unit separator
    ];

    const results: FormPathEntry[] = await vscode.window.withProgress(
        { location: vscode.ProgressLocation.Notification, title: `Validating form paths via Playwright...` },
        () => new Promise((resolve, reject) => {
            const child = spawn(pythonPath, args, { shell: false });
            let buf = "";
            child.stdout?.on("data", d => { const s = d.toString(); out.append(s); buf += s; });
            child.stderr?.on("data", d => out.append(d.toString()));
            child.on("error", reject);
            child.on("close", (code) => {
                if (code !== 0) return reject(new Error(`validate exited with code ${code}`));
                // Parse RESULT lines: RESULT|path|menuItem|error
                // Parse ADVICE lines: ADVICE|path|backend|flags|reason
                const out: FormPathEntry[] = [];
                const adviceByPath = new Map<string, { backend?: "mcp" | "playwright"; flags?: string; reason?: string }>();
                for (const ln of buf.split(/\r?\n/)) {
                    if (ln.startsWith("ADVICE|")) {
                        const parts = ln.split("|");
                        const p = parts[1];
                        const b = (parts[2] || "").trim();
                        const backend = (b === "mcp" || b === "playwright") ? b : undefined;
                        adviceByPath.set(p, { backend, flags: parts[3] || "", reason: parts[4] || "" });
                        continue;
                    }
                    if (!ln.startsWith("RESULT|")) continue;
                    const parts = ln.split("|");
                    out.push({ path: parts[1], menuItem: parts[2] || undefined, validated: !!parts[2], error: parts[3] || undefined });
                }
                // Attach advice to matching RESULT entries.
                for (const r of out) {
                    const a = adviceByPath.get(r.path);
                    if (!a) continue;
                    r.recommendedBackend = a.backend;
                    r.recommendedFlags = a.flags || undefined;
                    r.recommendedReason = a.reason || undefined;
                }
                resolve(out);
            });
        })
    );

    // Merge with existing (replace by path)
    const merged = [...existing];
    for (const r of results) {
        const idx = merged.findIndex(x => x.path === r.path);
        if (idx >= 0) merged[idx] = r; else merged.push(r);
    }
    await cfg.update("formPaths", merged, vscode.ConfigurationTarget.Workspace);

    const ok = results.filter(r => r.validated).length;
    const fail = results.length - ok;
    out.appendLine(`\nValidated: ${ok} OK, ${fail} failed.`);
    vscode.window.showInformationMessage(`Form paths: ${ok} validated, ${fail} failed. See Output for details.`);
}
