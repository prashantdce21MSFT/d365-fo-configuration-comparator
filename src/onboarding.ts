import * as vscode from "vscode";
import * as https from "https";
import { spawn } from "child_process";

const SECRET_KEY_CLIENT_SECRET = "d365FormExtractor.mcpClientSecret";

interface EnvEntry {
    label: string;
    baseUrl: string;
    tenantId?: string;
    mcpUrl?: string;
    mcpAvailable?: boolean;
}

interface McpAuth {
    tenantId: string;
    clientId: string;
}

async function which(cmd: string): Promise<boolean> {
    return new Promise((resolve) => {
        const p = spawn(process.platform === "win32" ? "where" : "which", [cmd], { shell: false });
        p.on("close", (code) => resolve(code === 0));
        p.on("error", () => resolve(false));
    });
}

async function run(cmd: string, args: string[]): Promise<{ code: number; stdout: string; stderr: string }> {
    return new Promise((resolve) => {
        const p = spawn(cmd, args, { shell: process.platform === "win32" });
        let stdout = "";
        let stderr = "";
        p.stdout?.on("data", (d) => (stdout += d.toString()));
        p.stderr?.on("data", (d) => (stderr += d.toString()));
        p.on("close", (code) => resolve({ code: code ?? -1, stdout, stderr }));
        p.on("error", () => resolve({ code: -1, stdout, stderr }));
    });
}

function postForm(url: string, form: Record<string, string>, timeoutMs = 15000):
    Promise<{ status: number; body: string }> {
    return new Promise((resolve) => {
        try {
            const u = new URL(url);
            const data = Buffer.from(
                Object.entries(form)
                    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
                    .join("&"),
                "utf8"
            );
            const req = https.request({
                method: "POST",
                hostname: u.hostname,
                port: u.port || 443,
                path: u.pathname + u.search,
                headers: {
                    "Content-Type": "application/x-www-form-urlencoded",
                    "Content-Length": data.length,
                },
                timeout: timeoutMs,
            }, (res) => {
                let buf = "";
                res.on("data", (c) => (buf += c.toString()));
                res.on("end", () => resolve({ status: res.statusCode || 0, body: buf }));
            });
            req.on("error", () => resolve({ status: 0, body: "" }));
            req.on("timeout", () => { req.destroy(); resolve({ status: 0, body: "" }); });
            req.write(data);
            req.end();
        } catch { resolve({ status: 0, body: "" }); }
    });
}

export async function getClientCredentialsToken(
    auth: McpAuth,
    secret: string,
    resource: string
): Promise<{ token?: string; error?: string }> {
    const res = await postForm(
        `https://login.microsoftonline.com/${auth.tenantId}/oauth2/v2.0/token`,
        {
            grant_type: "client_credentials",
            client_id: auth.clientId,
            client_secret: secret,
            scope: resource.replace(/\/$/, "") + "/.default",
        }
    );
    if (res.status === 200) {
        try {
            const j = JSON.parse(res.body);
            if (j.access_token) return { token: j.access_token as string };
            return { error: "response missing access_token" };
        } catch { return { error: "invalid JSON in token response" }; }
    }
    let err = `HTTP ${res.status}`;
    try {
        const j = JSON.parse(res.body);
        if (j.error_description) err += `: ${j.error_description.split("\n")[0]}`;
        else if (j.error) err += `: ${j.error}`;
    } catch { /* ignore */ }
    return { error: err };
}

function postJson(
    url: string,
    body: any,
    headers: Record<string, string>,
    timeoutMs = 15000
): Promise<{ status: number; body: string }> {
    return new Promise((resolve) => {
        try {
            const u = new URL(url);
            const data = Buffer.from(JSON.stringify(body), "utf8");
            const req = https.request(
                {
                    method: "POST",
                    hostname: u.hostname,
                    port: u.port || 443,
                    path: u.pathname + u.search,
                    headers: {
                        "Content-Type": "application/json",
                        "Content-Length": data.length,
                        ...headers,
                    },
                    timeout: timeoutMs,
                },
                (res) => {
                    let buf = "";
                    res.on("data", (c) => (buf += c.toString()));
                    res.on("end", () => resolve({ status: res.statusCode || 0, body: buf }));
                }
            );
            req.on("error", () => resolve({ status: 0, body: "" }));
            req.on("timeout", () => {
                req.destroy();
                resolve({ status: 0, body: "" });
            });
            req.write(data);
            req.end();
        } catch {
            resolve({ status: 0, body: "" });
        }
    });
}

async function tryGetToken(
    baseUrl: string,
    tenantId?: string
): Promise<{ ok: boolean; tenantId?: string; err: string }> {
    const args = ["account", "get-access-token", "--resource", baseUrl, "--query", "accessToken", "-o", "tsv"];
    if (tenantId) args.push("--tenant", tenantId);
    const r = await run("az", args);
    if (r.code === 0 && r.stdout.trim()) {
        let tid = tenantId;
        if (!tid) {
            const who = await run("az", ["account", "show", "--query", "tenantId", "-o", "tsv"]);
            if (who.code === 0) tid = who.stdout.trim() || undefined;
        }
        return { ok: true, tenantId: tid, err: "" };
    }
    return { ok: false, err: (r.stderr || r.stdout).split("\n")[0].trim() };
}

async function azLoginFor(env: EnvEntry, out: vscode.OutputChannel): Promise<boolean> {
    // Force a fresh interactive sign-in per environment (no cached-token shortcut).
    out.appendLine(`  [${env.label}] forcing fresh az login for ${env.baseUrl} ...`);

    let tenantArg: string | undefined = env.tenantId;
    const useTenant = await vscode.window.showQuickPick(
        [
            { label: tenantArg ? `$(check) Use saved tenant (${tenantArg})` : "$(globe) Use default tenant", id: "default" },
            { label: "$(edit) Enter tenant ID / domain…", id: "ask" },
        ],
        {
            title: `Sign-in for ${env.label}`,
            placeHolder: `${env.baseUrl}`,
            ignoreFocusOut: true,
        }
    );
    if (!useTenant) return false;
    if (useTenant.id === "ask") {
        const t = await vscode.window.showInputBox({
            title: `Tenant ID or domain for ${env.label}`,
            prompt: "AAD tenant GUID or domain (e.g. contoso.onmicrosoft.com). Leave empty for default.",
            value: env.tenantId || "",
            ignoreFocusOut: true,
        });
        tenantArg = t?.trim() || undefined;
    }

    const term = vscode.window.createTerminal(`az login: ${env.label}`);
    term.show();
    const tenantFlag = tenantArg ? ` --tenant ${tenantArg}` : "";
    term.sendText(`az login --scope ${env.baseUrl}/.default --allow-no-subscriptions${tenantFlag}`);

    // Loop: modal dialog blocks until user finishes browser sign-in or cancels.
    while (true) {
        const action = await vscode.window.showInformationMessage(
            `Complete sign-in for ${env.label} in the terminal '${term.name}'.\n\nClick 'I'm done' AFTER you finish the browser sign-in.`,
            { modal: true },
            "I'm done",
            "Skip this env"
        );
        if (action === "Skip this env" || !action) {
            out.appendLine(`  [${env.label}] sign-in skipped.`);
            return false;
        }
        const res = await tryGetToken(env.baseUrl, tenantArg);
        if (res.ok) {
            env.tenantId = res.tenantId || tenantArg;
            out.appendLine(`  [${env.label}] OK${env.tenantId ? ` (tenant ${env.tenantId})` : ""}`);
            return true;
        }
        out.appendLine(`  [${env.label}] token check failed: ${res.err}`);
        const retry = await vscode.window.showWarningMessage(
            `Sign-in for ${env.label} not detected yet. ${res.err || ""}`,
            { modal: true },
            "Retry",
            "Skip this env"
        );
        if (retry !== "Retry") return false;
    }
}

async function probeMcp(env: EnvEntry, auth: McpAuth | undefined, secret: string | undefined, out: vscode.OutputChannel): Promise<boolean> {
    const mcpUrl = env.mcpUrl?.trim() || `${env.baseUrl}/mcp`;
    const mcpOrigin = (() => { try { const u = new URL(mcpUrl); return `${u.protocol}//${u.host}`; } catch { return env.baseUrl; } })();

    let token: string | undefined;
    if (auth && secret) {
        const r = await getClientCredentialsToken(auth, secret, mcpOrigin);
        if (!r.token) {
            out.appendLine(`  [${env.label}] MCP probe: app-reg token failed (${r.error}). Falling back to az.`);
        } else {
            token = r.token;
        }
    }
    if (!token) {
        const args = [
            "account", "get-access-token", "--resource", mcpOrigin,
            ...(env.tenantId ? ["--tenant", env.tenantId] : []),
            "--query", "accessToken", "-o", "tsv",
        ];
        const tokRes = await run("az", args);
        if (tokRes.code !== 0 || !tokRes.stdout.trim()) {
            out.appendLine(`  [${env.label}] MCP probe (${mcpUrl}): no token for ${mcpOrigin}`);
            return false;
        }
        token = tokRes.stdout.trim();
    }

    const res = await postJson(
        mcpUrl,
        {
            jsonrpc: "2.0",
            id: 1,
            method: "initialize",
            params: {
                protocolVersion: "2024-11-05",
                capabilities: {},
                clientInfo: { name: "d365-form-extractor-probe", version: "1.0" },
            },
        },
        { Authorization: `Bearer ${token}`, Accept: "application/json, text/event-stream" }
    );
    if (res.status === 200) { out.appendLine(`  [${env.label}] MCP (${mcpUrl}): available`); return true; }
    if (res.status === 401 || res.status === 403) {
        out.appendLine(`  [${env.label}] MCP (${mcpUrl}): ${res.status} (endpoint not enabled, audience rejected, or app/user lacks API permission/role)`);
    } else if (res.status === 0) {
        out.appendLine(`  [${env.label}] MCP (${mcpUrl}): unreachable`);
    } else {
        out.appendLine(`  [${env.label}] MCP (${mcpUrl}): HTTP ${res.status}`);
    }
    return false;
}

async function addEnvironmentsLoop(envs: EnvEntry[]): Promise<void> {
    while (true) {
        const choice = await vscode.window.showQuickPick(
            [
                { label: "$(add) Add environment", id: "add" },
                { label: "$(check) Done", id: "done" },
                { label: "$(trash) Remove all", id: "clear" },
            ],
            {
                title: "D365 Environments",
                placeHolder: "Current: " + (envs.length ? envs.map((e) => e.label).join(", ") : "none"),
                ignoreFocusOut: true,
            }
        );
        if (!choice || choice.id === "done") break;
        if (choice.id === "clear") {
            envs.length = 0;
            continue;
        }
        const label = await vscode.window.showInputBox({
            prompt: "Environment label (e.g. UAT, Config, Production)",
            ignoreFocusOut: true,
        });
        if (!label) continue;
        const baseUrl = await vscode.window.showInputBox({
            prompt: "Base URL (no trailing slash)",
            placeHolder: "https://<your-env>.sandbox.operations.dynamics.com",
            ignoreFocusOut: true,
            validateInput: (v) => (v?.startsWith("https://") ? null : "Must start with https://"),
        });
        if (!baseUrl) continue;
        const mcpUrl = await vscode.window.showInputBox({
            prompt: "MCP endpoint URL (optional, leave empty if unknown or to use <baseUrl>/mcp)",
            placeHolder: "https://org541b3d42.operations.dynamics.com/mcp",
            ignoreFocusOut: true,
            validateInput: (v) => (!v || v.startsWith("https://") ? null : "Must start with https://"),
        });
        envs.push({
            label,
            baseUrl: baseUrl.replace(/\/$/, ""),
            ...(mcpUrl?.trim() ? { mcpUrl: mcpUrl.trim() } : {}),
        });
    }
}

export async function ensureOnboarded(
    context: vscode.ExtensionContext,
    out: vscode.OutputChannel
): Promise<boolean> {
    const done = context.globalState.get<boolean>("onboarded", false);
    if (done) return true;
    const choice = await vscode.window.showInformationMessage(
        "D365 FO Configuration Comparator is not configured yet. Run onboarding now?",
        "Run Onboarding",
        "Cancel"
    );
    if (choice !== "Run Onboarding") return false;
    return await runOnboarding(context, out);
}

export async function runOnboarding(
    context: vscode.ExtensionContext,
    out: vscode.OutputChannel
): Promise<boolean> {
    out.show(true);
    out.appendLine("=== D365 FO Configuration Comparator — Onboarding ===");

    const TOTAL = 8;
    const cfg = () => vscode.workspace.getConfiguration("d365FormExtractor");

    // [1/7] Python
    out.appendLine(`\n[1/${TOTAL}] Checking Python...`);
    let pythonPath = cfg().get<string>("pythonPath") || "python";
    let pyOk = await run(pythonPath, ["--version"]);
    if (pyOk.code !== 0) {
        const input = await vscode.window.showInputBox({
            prompt: "Python (3.10+) not found on PATH. Enter full path to python.exe",
            placeHolder: "C:\\Python313\\python.exe",
            ignoreFocusOut: true,
        });
        if (!input) return false;
        pythonPath = input;
        await cfg().update("pythonPath", pythonPath, vscode.ConfigurationTarget.Global);
        pyOk = await run(pythonPath, ["--version"]);
    }
    out.appendLine(`  Python: ${(pyOk.stdout || pyOk.stderr).trim()}`);

    // [2/7] Python deps
    out.appendLine(`\n[2/${TOTAL}] Ensuring Python dependencies (requests, msal, openpyxl, playwright)...`);
    const pipRes = await run(pythonPath, [
        "-m", "pip", "install", "--quiet", "requests", "msal", "openpyxl", "playwright",
    ]);
    if (pipRes.code !== 0) {
        out.appendLine("  WARNING: pip install failed:\n" + pipRes.stderr);
        const cont = await vscode.window.showWarningMessage(
            "Some Python packages failed to install. Continue anyway?",
            "Continue",
            "Cancel"
        );
        if (cont !== "Continue") return false;
    } else {
        out.appendLine("  OK");
    }

    // [3/7] Azure CLI present?
    out.appendLine(`\n[3/${TOTAL}] Checking Azure CLI...`);
    const azPresent = await which("az");
    if (!azPresent) {
        const inst = await vscode.window.showWarningMessage(
            "Azure CLI ('az') is required for D365 ERP MCP authentication. Install from https://aka.ms/installazurecliwindows then re-run onboarding. Continue with Playwright only?",
            "Continue (PW only)",
            "Cancel"
        );
        if (inst !== "Continue (PW only)") return false;
        out.appendLine("  Skipped (az not installed). Forcing defaultBackend=playwright.");
        await cfg().update("defaultBackend", "playwright", vscode.ConfigurationTarget.Global);
    } else {
        const accRes = await run("az", ["account", "show", "--query", "user.name", "-o", "tsv"]);
        const currentUser = (accRes.code === 0) ? accRes.stdout.trim() : "";
        if (currentUser) {
            out.appendLine(`  Currently signed in as: ${currentUser}`);
        } else {
            out.appendLine("  Not signed in to az.");
        }
        const items = currentUser
            ? [
                { label: `$(check) Continue as ${currentUser}`, id: "keep" },
                { label: "$(sign-in) Sign in again (same or different account)", id: "login" },
                { label: "$(sign-out) Sign out, then sign in", id: "logout" },
              ]
            : [
                { label: "$(sign-in) Sign in now (az login)", id: "login" },
                { label: "$(debug-step-over) Skip (sign in later per environment)", id: "skip" },
              ];
        const pick = await vscode.window.showQuickPick(items, {
            title: "Azure CLI sign-in",
            placeHolder: "Choose how to authenticate az",
            ignoreFocusOut: true,
        });
        if (!pick) return false;
        if (pick.id === "logout") {
            await run("az", ["logout"]);
            out.appendLine("  Signed out of az.");
        }
        if (pick.id === "logout" || pick.id === "login") {
            const tenantInput = await vscode.window.showInputBox({
                title: "Tenant (optional)",
                prompt: "Tenant GUID or domain to sign in to. Leave empty for default.",
                ignoreFocusOut: true,
            });
            const tenant = tenantInput?.trim();
            const term = vscode.window.createTerminal("az login");
            term.show();
            term.sendText(`az login${tenant ? ` --tenant ${tenant}` : ""}`);
            // Modal so it doesn't auto-dismiss while the browser flow is running.
            while (true) {
                const done = await vscode.window.showInformationMessage(
                    `Complete sign-in in the terminal '${term.name}'.\n\nClick 'I'm done' AFTER you finish the browser sign-in.`,
                    { modal: true },
                    "I'm done",
                    "Cancel"
                );
                if (done !== "I'm done") { out.appendLine("  Sign-in cancelled."); break; }
                const after = await run("az", ["account", "show", "--query", "user.name", "-o", "tsv"]);
                if (after.code === 0 && after.stdout.trim()) {
                    out.appendLine(`  Signed in as: ${after.stdout.trim()}`);
                    break;
                }
                const retry = await vscode.window.showWarningMessage(
                    "Still not signed in. Retry?",
                    { modal: true }, "Retry", "Skip"
                );
                if (retry !== "Retry") { out.appendLine("  WARNING: still not signed in."); break; }
            }
        } else {
            out.appendLine(`  Continuing as ${currentUser}.`);
        }
    }

    // [4/8] App registration for MCP (optional)
    out.appendLine(`\n[4/${TOTAL}] MCP app registration (optional)...`);
    const existingAuth = cfg().get<McpAuth | undefined>("mcpAuth");
    const hasSecret = !!(await context.secrets.get(SECRET_KEY_CLIENT_SECRET));
    let mcpAuth: McpAuth | undefined = existingAuth || undefined;
    {
        const items: (vscode.QuickPickItem & { id: string })[] = [];
        if (mcpAuth && hasSecret) {
            items.push({ label: `$(check) Keep existing app reg (clientId=${mcpAuth.clientId.slice(0, 8)}\u2026)`, id: "keep" });
            items.push({ label: "$(edit) Replace app registration", id: "set" });
            items.push({ label: "$(trash) Remove app registration", id: "clear" });
        } else {
            items.push({ label: "$(key) Configure app registration (tenant + client id + secret)", id: "set" });
            items.push({ label: "$(debug-step-over) Skip (use az CLI auth only)", id: "skip" });
        }
        const pick = await vscode.window.showQuickPick(items, {
            title: "MCP app registration",
            placeHolder: "Provide an Azure AD app reg for MCP, or skip to use the az CLI session.",
            ignoreFocusOut: true,
        });
        if (!pick) return false;
        if (pick.id === "clear") {
            await cfg().update("mcpAuth", undefined, vscode.ConfigurationTarget.Global);
            await context.secrets.delete(SECRET_KEY_CLIENT_SECRET);
            mcpAuth = undefined;
            out.appendLine("  App registration cleared.");
        } else if (pick.id === "set") {
            const tenantId = (await vscode.window.showInputBox({
                title: "App registration \u2014 Tenant ID",
                value: mcpAuth?.tenantId || "",
                prompt: "AAD tenant GUID where the app registration lives.",
                ignoreFocusOut: true,
                validateInput: (v) => (v?.trim() ? null : "Required"),
            }))?.trim();
            if (!tenantId) return false;
            const clientId = (await vscode.window.showInputBox({
                title: "App registration \u2014 Client (Application) ID",
                value: mcpAuth?.clientId || "",
                prompt: "Application (client) ID GUID.",
                ignoreFocusOut: true,
                validateInput: (v) => (v?.trim() ? null : "Required"),
            }))?.trim();
            if (!clientId) return false;
            const clientSecret = (await vscode.window.showInputBox({
                title: "App registration \u2014 Client secret",
                prompt: "Stored securely in VS Code SecretStorage \u2014 never written to settings.json.",
                password: true,
                ignoreFocusOut: true,
                validateInput: (v) => (v?.trim() ? null : "Required"),
            }))?.trim();
            if (!clientSecret) return false;
            mcpAuth = { tenantId, clientId };
            await cfg().update("mcpAuth", mcpAuth, vscode.ConfigurationTarget.Global);
            await context.secrets.store(SECRET_KEY_CLIENT_SECRET, clientSecret);
            out.appendLine(`  Saved app reg (tenant ${tenantId}, client ${clientId.slice(0, 8)}\u2026). Secret stored in SecretStorage.`);
        } else {
            out.appendLine("  Skipped (will use az CLI).");
        }
    }

    // [5/8] Environments
    out.appendLine(`\n[5/${TOTAL}] Configuring D365 environments...`);
    const existing = (cfg().get<EnvEntry[]>("environments") || []).slice();
    let envs: EnvEntry[] = existing;
    if (envs.length > 0) {
        out.appendLine(`  Existing: ${envs.map((e) => `${e.label} → ${e.baseUrl}`).join(", ")}`);
        const keep = await vscode.window.showQuickPick(
            [
                { label: "$(check) Keep existing", id: "keep" },
                { label: "$(add) Add more", id: "add" },
                { label: "$(refresh) Start over", id: "reset" },
            ],
            { title: `Environments already configured (${envs.length})`, ignoreFocusOut: true }
        );
        if (!keep) return false;
        if (keep.id === "reset") envs = [];
        if (keep.id !== "keep") await addEnvironmentsLoop(envs);
    } else {
        await addEnvironmentsLoop(envs);
    }
    await cfg().update("environments", envs, vscode.ConfigurationTarget.Global);
    out.appendLine(`  Saved ${envs.length} environment(s).`);
    if (envs.length === 0) {
        vscode.window.showWarningMessage("No environments configured. Add at least one to continue.");
        return false;
    }

    // [6/8] Verify each env can get a token from the current az session (no extra sign-in).
    out.appendLine(`\n[6/${TOTAL}] Verifying token access for each environment (using the az session from step 3)...`);
    if (azPresent) {
        const failed: { env: EnvEntry; err: string }[] = [];
        for (const e of envs) {
            const res = await tryGetToken(e.baseUrl, e.tenantId);
            if (res.ok) {
                e.tenantId = res.tenantId || e.tenantId;
                out.appendLine(`  [${e.label}] OK${e.tenantId ? ` (tenant ${e.tenantId})` : ""}`);
            } else {
                out.appendLine(`  [${e.label}] FAILED: ${res.err}`);
                failed.push({ env: e, err: res.err });
            }
        }
        await cfg().update("environments", envs, vscode.ConfigurationTarget.Global);
        if (failed.length > 0) {
            const msg = failed.map((f) => `${f.env.label}: ${f.err}`).join("\n");
            vscode.window.showWarningMessage(
                `Token check failed for ${failed.length} environment(s). The account from step 3 may not have access, or the env is in a different tenant.\n\n${msg}\n\nFix with 'az login --tenant <id>' in a terminal, then re-run onboarding.`,
                { modal: true },
                "OK"
            );
        }
    } else {
        out.appendLine("  Skipped (az not installed).");
    }

    // [7/8] MCP probe
    out.appendLine(`\n[7/${TOTAL}] Probing D365 ERP MCP endpoint on each environment...`);
    let anyMcp = false;
    const secret = mcpAuth ? await context.secrets.get(SECRET_KEY_CLIENT_SECRET) : undefined;
    if (azPresent || (mcpAuth && secret)) {
        for (const e of envs) {
            e.mcpAvailable = await probeMcp(e, mcpAuth, secret, out);
            anyMcp = anyMcp || !!e.mcpAvailable;
        }
        await cfg().update("environments", envs, vscode.ConfigurationTarget.Global);
        const currentBackend = cfg().get<string>("defaultBackend") || "ask";
        if (!anyMcp && currentBackend !== "playwright") {
            await cfg().update("defaultBackend", "playwright", vscode.ConfigurationTarget.Global);
            out.appendLine("  No environment supports MCP → defaultBackend set to 'playwright'.");
            vscode.window.showInformationMessage(
                "MCP not available on any environment. Using Playwright backend by default."
            );
        }
    } else {
        out.appendLine("  Skipped (az not installed).");
    }

    // [8/8] Form paths
    out.appendLine(`\n[8/${TOTAL}] Form paths...`);
    const existingPaths = cfg().get<any[]>("formPaths") || [];
    const validatedCount = existingPaths.filter((p) => p && p.validated).length;
    let runFormPaths = false;
    if (validatedCount > 0) {
        const c = await vscode.window.showQuickPick(
            [
                { label: `$(check) Keep ${validatedCount} validated path(s)`, id: "keep" },
                { label: "$(edit) Add / re-validate paths", id: "edit" },
            ],
            { title: "Form paths", ignoreFocusOut: true }
        );
        runFormPaths = c?.id === "edit";
    } else {
        const c = await vscode.window.showInformationMessage(
            "Configure and validate form paths now? (Recommended — avoids typing menu items per extraction.)",
            "Configure now",
            "Skip"
        );
        runFormPaths = c === "Configure now";
    }
    if (runFormPaths) {
        try {
            await vscode.commands.executeCommand("d365FormExtractor.configureFormPaths");
        } catch (e: any) {
            out.appendLine(`  configureFormPaths failed: ${e?.message ?? e}`);
        }
    } else {
        out.appendLine("  Skipped.");
    }

    // Output folder
    out.appendLine(`\nOutput folder...`);
    const ws = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath;
    const defOut = ws ? ws + "\\d365-extracts" : "";
    const currentOut = cfg().get<string>("outputFolder");
    if (!currentOut) {
        const outFolder = await vscode.window.showInputBox({
            prompt: "Where should extraction Excels be written?",
            value: defOut,
            ignoreFocusOut: true,
        });
        if (outFolder) await cfg().update("outputFolder", outFolder, vscode.ConfigurationTarget.Global);
    } else {
        out.appendLine(`  ${currentOut} (already set)`);
    }

    await context.globalState.update("onboarded", true);
    out.appendLine("\n=== Onboarding complete ===");
    vscode.window.showInformationMessage(
        "D365 FO Configuration Comparator onboarding complete. Run 'D365: Extract Form' to begin."
    );
    return true;
}
