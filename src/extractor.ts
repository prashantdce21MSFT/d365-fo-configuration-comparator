import * as vscode from "vscode";
import * as path from "path";
import { spawn } from "child_process";

function runCmd(cmd: string, args: string[]): Promise<{ code: number; stdout: string; stderr: string }> {
    return new Promise((resolve) => {
        const p = spawn(cmd, args, { shell: process.platform === "win32" });
        let stdout = "", stderr = "";
        p.stdout?.on("data", d => stdout += d.toString());
        p.stderr?.on("data", d => stderr += d.toString());
        p.on("close", code => resolve({ code: code ?? -1, stdout, stderr }));
        p.on("error", () => resolve({ code: -1, stdout, stderr }));
    });
}

async function ensureAzLogin(envs: { label: string; baseUrl: string; tenantId?: string }[], out: vscode.OutputChannel): Promise<boolean> {
    out.appendLine("\n[az] Verifying access for each environment...");
    for (const e of envs) {
        const resource = e.baseUrl.replace(/\/$/, "");
        let attempt = 0;
        while (true) {
            const args = ["account", "get-access-token", "--resource", resource, "--query", "accessToken", "-o", "tsv"];
            if (e.tenantId) args.push("--tenant", e.tenantId);
            const tok = await runCmd("az", args);
            if (tok.code === 0 && tok.stdout.trim()) {
                // Auto-capture the tenant we just used (if not set)
                if (!e.tenantId) {
                    const who = await runCmd("az", ["account", "show", "--query", "tenantId", "-o", "tsv"]);
                    if (who.code === 0 && who.stdout.trim()) {
                        e.tenantId = who.stdout.trim();
                        const cfg = vscode.workspace.getConfiguration("d365FormExtractor");
                        const all = cfg.get<any[]>("environments") || [];
                        const idx = all.findIndex(x => x.label === e.label);
                        if (idx >= 0) { all[idx].tenantId = e.tenantId; await cfg.update("environments", all, vscode.ConfigurationTarget.Workspace); }
                    }
                }
                out.appendLine(`  [${e.label}] OK${e.tenantId ? ` (tenant ${e.tenantId})` : ""}`);
                break;
            }
            attempt++;
            if (attempt > 1) {
                vscode.window.showErrorMessage(`Still cannot get token for ${e.label}. Aborting.`);
                return false;
            }
            const err = (tok.stderr || "").split("\n")[0].trim();
            const choice = await vscode.window.showWarningMessage(
                `Sign-in required for ${e.label} (${resource}). ${err}`,
                "Sign in", "Skip env", "Cancel"
            );
            if (choice === "Cancel" || !choice) return false;
            if (choice === "Skip env") break;
            const term = vscode.window.createTerminal(`az login: ${e.label}`);
            term.show();
            term.sendText(`az login --scope ${resource}/.default --allow-no-subscriptions`);
            await vscode.window.showInformationMessage(
                `Complete sign-in for ${e.label} in the terminal, then click Retry.`,
                { modal: false }, "Retry"
            );
        }
    }
    return true;
}

export async function runExtraction(
    context: vscode.ExtensionContext,
    out: vscode.OutputChannel
): Promise<void> {
    const cfg = vscode.workspace.getConfiguration("d365FormExtractor");
    const envs = cfg.get<{ label: string; baseUrl: string }[]>("environments") || [];
    if (envs.length === 0) {
        vscode.window.showErrorMessage("No D365 environments configured. Run onboarding first.");
        return;
    }

    // Pick legal entity
    const le = await vscode.window.showInputBox({
        title: "Legal Entity",
        prompt: "Legal entity code (e.g. my30, sg60)",
        placeHolder: "my30",
        ignoreFocusOut: true,
        validateInput: (v) => (v && v.trim() ? null : "Required"),
    });
    if (!le) return;

    // Pick form: use saved form paths if available, else free text.
    // Multi-select supported — selecting N validated paths runs N extractions sequentially.
    const savedPaths = cfg.get<{
        path: string;
        menuItem?: string;
        validated?: boolean;
        recommendedBackend?: "mcp" | "playwright";
        recommendedFlags?: string;
        recommendedReason?: string;
    }[]>("formPaths") || [];
    type FormJob = {
        form: string;
        formPath?: string;
        recommendedBackend?: "mcp" | "playwright";
        recommendedFlags?: string;
        recommendedReason?: string;
    };
    const jobs: FormJob[] = [];
    const validated = savedPaths.filter(p => p.validated && p.menuItem);
    if (validated.length > 0) {
        const items = validated.map(p => {
            const adviceTag = p.recommendedBackend
                ? `  • ${p.recommendedBackend === "mcp" ? "MCP" : "Playwright"}${p.recommendedFlags ? " + " + p.recommendedFlags : ""}`
                : "";
            return {
                label: `$(check) ${p.path}`,
                description: `mi=${p.menuItem}${adviceTag}`,
                detail: p.recommendedReason || undefined,
                mi: p.menuItem!,
                pathTxt: p.path,
                recommendedBackend: p.recommendedBackend,
                recommendedFlags: p.recommendedFlags,
                recommendedReason: p.recommendedReason,
                picked: false,
            };
        });
        const picks = await vscode.window.showQuickPick(items, {
            title: "Pick form(s)",
            placeHolder: "Select one or more validated form paths (Space to toggle, Enter to confirm). Pick none to enter manually.",
            canPickMany: true,
            ignoreFocusOut: true,
        });
        if (picks === undefined) return; // cancelled
        for (const p of picks) jobs.push({
            form: p.mi,
            formPath: p.pathTxt,
            recommendedBackend: p.recommendedBackend,
            recommendedFlags: p.recommendedFlags,
            recommendedReason: p.recommendedReason,
        });
    }
    if (jobs.length === 0) {
        const f = await vscode.window.showInputBox({
            title: "Form / Menu Item",
            prompt: "Menu item name (mi=…), e.g. smmParameters",
            placeHolder: "smmParameters",
            ignoreFocusOut: true,
            validateInput: (v) => (v && v.trim() ? null : "Required"),
        });
        if (!f) return;
        const fp = await vscode.window.showInputBox({
            title: "Form Path (optional)",
            prompt: "Human-readable navigation path",
            ignoreFocusOut: true,
        });
        jobs.push({ form: f.trim(), formPath: fp || undefined });
    }

    // Choose environments to extract (1 or both)
    const envPicks = await vscode.window.showQuickPick(
        envs.map((e) => ({ label: e.label, description: e.baseUrl, picked: true, env: e })),
        {
            canPickMany: true,
            title: "Select environments",
            placeHolder: "Pick one or more environments",
            ignoreFocusOut: true,
        }
    );
    if (!envPicks || envPicks.length === 0) return;

    // Backend choice — if all selected jobs share a single recommendation,
    // use it as the default; otherwise show the picker.
    let backend = cfg.get<string>("defaultBackend") || "ask";
    const recBackends = new Set(jobs.map(j => j.recommendedBackend).filter(Boolean));
    const consensusBackend: "mcp" | "playwright" | undefined =
        recBackends.size === 1 ? (jobs.find(j => j.recommendedBackend)!.recommendedBackend) : undefined;
    if (backend === "ask") {
        const items = [
            {
                label: "$(database) D365 ERP MCP",
                description: "Headless. Uses 'az account get-access-token'.",
                id: "mcp",
                picked: consensusBackend === "mcp",
            },
            {
                label: "$(browser) Playwright (Chrome CDP)",
                description: "Drives your logged-in Chrome on port 9222.",
                id: "playwright",
                picked: consensusBackend === "playwright",
            },
        ];
        if (consensusBackend) {
            const idx = items.findIndex(it => it.id === consensusBackend);
            const reason = jobs.find(j => j.recommendedReason)?.recommendedReason;
            items[idx].description += `  — ✨ recommended${reason ? ": " + reason : ""}`;
        }
        const pick = await vscode.window.showQuickPick(items, { title: "Extraction backend", ignoreFocusOut: true });
        if (!pick) return;
        backend = pick.id;
    }

    // Pre-flight: verify az login for MCP backend
    if (backend === "mcp") {
        out.show(true);
        const ok = await ensureAzLogin(envPicks.map(p => p.env), out);
        if (!ok) return;
    }

    // Resolve output folder
    let outFolder = cfg.get<string>("outputFolder") || "";
    if (!outFolder) {
        const ws = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath;
        outFolder = ws ? path.join(ws, "d365-extracts") : path.join(require("os").tmpdir(), "d365-extracts");
    }

    // Build python args (per-job inside loop)
    const pythonPath = cfg.get<string>("pythonPath") || "python";
    const scriptPath = path.join(context.extensionPath, "python", "extract.py");

    out.show(true);
    out.appendLine(`\n=== Extracting ${jobs.length} form(s) for ${le} via ${backend} ===`);

    let okCount = 0, failCount = 0;
    const failed: string[] = [];

    for (let i = 0; i < jobs.length; i++) {
        const { form, formPath, recommendedFlags } = jobs[i];
        const args: string[] = [
            scriptPath,
            "--backend", backend,
            "--form", form.trim(),
            "--le", le.trim().toLowerCase(),
            "--out-dir", outFolder,
        ];
        if (formPath) args.push("--form-path", formPath);
        // Auto-pass --per-row-detail when Playwright was recommended with that flag.
        if (backend === "playwright" && (recommendedFlags || "").split(",").map(s => s.trim()).includes("per-row-detail")) {
            args.push("--per-row-detail");
        }
        for (const e of envPicks) {
            const t = (e.env as any).tenantId as string | undefined;
            const m = (e.env as any).mcpUrl as string | undefined;
            let val = `${e.env.label}=${e.env.baseUrl}`;
            if (t || m) val += `|${t || ""}`;
            if (m) val += `|${m}`;
            args.push("--env", val);
        }
        if (envPicks.length > 1) args.push("--diff");

        out.appendLine(`\n--- [${i + 1}/${jobs.length}] ${form} ${formPath ? `(${formPath})` : ""} ---`);
        out.appendLine(`Command: ${pythonPath} ${args.map((a) => (a.includes(" ") ? `"${a}"` : a)).join(" ")}\n`);

        try {
            await vscode.window.withProgress(
                { location: vscode.ProgressLocation.Notification, title: `Extracting ${form} (${i + 1}/${jobs.length}, ${backend})...`, cancellable: true },
                async (_progress, token) => {
                    const mcpAuth = cfg.get<{ tenantId: string; clientId: string } | undefined>("mcpAuth");
                    const clientSecret = mcpAuth ? await context.secrets.get("d365FormExtractor.mcpClientSecret") : undefined;
                    const childEnv: NodeJS.ProcessEnv = { ...process.env };
                    if (mcpAuth && clientSecret) {
                        childEnv.MCP_TENANT_ID = mcpAuth.tenantId;
                        childEnv.MCP_CLIENT_ID = mcpAuth.clientId;
                        childEnv.MCP_CLIENT_SECRET = clientSecret;
                    }
                    return new Promise<void>((resolve, reject) => {
                        const child = spawn(pythonPath, args, { shell: false, env: childEnv });
                        token.onCancellationRequested(() => child.kill());
                        child.stdout?.on("data", (d) => out.append(d.toString()));
                        child.stderr?.on("data", (d) => out.append(d.toString()));
                        child.on("error", reject);
                        child.on("close", (code) => {
                            if (code === 0) resolve();
                            else reject(new Error(`extract.py exited with code ${code}`));
                        });
                    });
                }
            );
            okCount++;
        } catch (e: any) {
            failCount++;
            failed.push(`${form}: ${e?.message ?? e}`);
            out.appendLine(`  [error] ${form}: ${e?.message ?? e}`);
        }
    }

    out.appendLine(`\n=== Done: ${okCount} OK, ${failCount} failed. Files in: ${outFolder} ===`);
    if (failed.length) {
        for (const f of failed) out.appendLine(`  - ${f}`);
    }
    const reveal = await vscode.window.showInformationMessage(
        `Extraction complete: ${okCount} OK, ${failCount} failed.`,
        "Open Folder",
        "Dismiss"
    );
    if (reveal === "Open Folder") {
        vscode.env.openExternal(vscode.Uri.file(outFolder));
    }
}
