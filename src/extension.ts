import * as vscode from "vscode";
import * as fs from "fs";
import * as path from "path";
import { runOnboarding, ensureOnboarded } from "./onboarding";
import { runExtraction } from "./extractor";
import { startChromeCdp } from "./chrome";
import { configureFormPaths } from "./formPaths";

const EXT_ID = "PrashantVerma.d365-form-extractor";
const WALKTHROUGH_ID = `${EXT_ID}#d365FormExtractor.gettingStarted`;

const DISCLAIMER_LINES = [
    "============================================================",
    " D365 FO Config Compare  (Beta)  --  USE AT YOUR OWN RISK",
    "============================================================",
    " This is an INDEPENDENT, COMMUNITY-BUILT tool.",
    " It is NOT a standard out-of-the-box Microsoft product,",
    " not affiliated with, endorsed by, or supported by Microsoft.",
    " Provided AS IS under the MIT license, with no warranty.",
    " NOT recommended for production environments --",
    " validate against UAT or a read-replica first.",
    " 'Microsoft' and 'Dynamics 365' are trademarks of",
    " Microsoft Corporation, used here only descriptively.",
    "============================================================",
];

async function openWalkthrough() {
    await vscode.commands.executeCommand("workbench.action.openWalkthrough", WALKTHROUGH_ID, false);
}

async function scaffoldWorkspace() {
    const folders = vscode.workspace.workspaceFolders;
    if (folders && folders.length > 0) {
        const r = await vscode.window.showInformationMessage(
            `Workspace already open: ${folders[0].uri.fsPath}`,
            "Open Different Folder", "Keep This"
        );
        if (r !== "Open Different Folder") return;
    }
    const input = await vscode.window.showInputBox({
        title: "Workspace folder",
        value: "C:\\D365-Form-Extractor",
        prompt: "Folder to create (if missing) and open as workspace",
        ignoreFocusOut: true,
    });
    if (!input) return;
    try { fs.mkdirSync(input, { recursive: true }); } catch { /* ignore */ }
    try { fs.mkdirSync(path.join(input, "d365-extracts"), { recursive: true }); } catch { /* ignore */ }
    await vscode.commands.executeCommand("vscode.openFolder", vscode.Uri.file(input), { forceNewWindow: false });
}

export function activate(context: vscode.ExtensionContext) {
    const out = vscode.window.createOutputChannel("D365 FO Config Compare");
    context.subscriptions.push(out);
    for (const line of DISCLAIMER_LINES) { out.appendLine(line); }
    out.appendLine("");

    // Persistent status-bar indicator -- always visible, clickable for full disclaimer
    const statusItem = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 100);
    statusItem.text = "$(warning) D365 Config Compare (Beta — not a Microsoft product)";
    statusItem.tooltip = "Independent community tool. NOT a Microsoft product. Provided AS IS under MIT, no warranty. Not for production. Click for details.";
    statusItem.command = "d365FormExtractor.showDisclaimer";
    statusItem.backgroundColor = new vscode.ThemeColor("statusBarItem.warningBackground");
    statusItem.show();
    context.subscriptions.push(statusItem);

    // Show disclaimer toast on every activation until the user acknowledges it
    if (!context.globalState.get<boolean>("disclaimerAcknowledged")) {
        out.show(true);
        vscode.window.showWarningMessage(
            "D365 FO Config Compare is an independent community tool (NOT a Microsoft product). Provided AS IS, no warranty. Not recommended for production — use UAT/read-replica first.",
            { modal: false },
            "I Understand", "View Details"
        ).then(choice => {
            if (choice === "I Understand") {
                context.globalState.update("disclaimerAcknowledged", true);
            } else if (choice === "View Details") {
                out.show(true);
            }
        });
    }

    // Open walkthrough on first activation
    if (!context.globalState.get<boolean>("welcomeShown")) {
        setTimeout(() => { openWalkthrough().catch(() => {}); context.globalState.update("welcomeShown", true); }, 1200);
    }

    context.subscriptions.push(
        vscode.commands.registerCommand("d365FormExtractor.showDisclaimer", async () => {
            out.show(true);
            const r = await vscode.window.showWarningMessage(
                "D365 FO Config Compare — independent community tool, NOT a Microsoft product. Provided AS IS under MIT, no warranty. Not recommended for production environments.",
                { modal: true, detail: DISCLAIMER_LINES.join("\n") },
                "I Understand"
            );
            if (r === "I Understand") {
                await context.globalState.update("disclaimerAcknowledged", true);
            }
        }),
        vscode.commands.registerCommand("d365FormExtractor.onboard", async () => {
            await runOnboarding(context, out);
        }),
        vscode.commands.registerCommand("d365FormExtractor.extract", async () => {
            try {
                const onboarded = await ensureOnboarded(context, out);
                if (!onboarded) return;
                await runExtraction(context, out);
            } catch (err: any) {
                vscode.window.showErrorMessage(`Extraction failed: ${err?.message ?? err}`);
                out.appendLine(`ERROR: ${err?.stack ?? err}`);
            }
        }),
        vscode.commands.registerCommand("d365FormExtractor.startChromeCdp", async () => {
            await startChromeCdp(out);
        }),
        vscode.commands.registerCommand("d365FormExtractor.clearSettings", async () => {
            const cfg = vscode.workspace.getConfiguration("d365FormExtractor");
            await cfg.update("environments", [], vscode.ConfigurationTarget.Global);
            await cfg.update("environments", undefined, vscode.ConfigurationTarget.Workspace);
            await cfg.update("formPaths", undefined, vscode.ConfigurationTarget.Workspace);
            await cfg.update("outputFolder", undefined, vscode.ConfigurationTarget.Workspace);
            await cfg.update("outputFolder", undefined, vscode.ConfigurationTarget.Global);
            await cfg.update("mcpAuth", undefined, vscode.ConfigurationTarget.Global);
            await context.secrets.delete("d365FormExtractor.mcpClientSecret");
            await context.globalState.update("onboarded", false);
            await context.globalState.update("welcomeShown", false);
            const r = await vscode.window.showInformationMessage(
                "D365 FO Config Compare settings cleared. Start fresh now?",
                "Start Getting Started", "Later"
            );
            if (r === "Start Getting Started") {
                await vscode.commands.executeCommand("d365FormExtractor.welcome");
            }
        }),
        vscode.commands.registerCommand("d365FormExtractor.configureFormPaths", async () => {
            try {
                await configureFormPaths(context, out);
            } catch (err: any) {
                vscode.window.showErrorMessage(`Configure Form Paths failed: ${err?.message ?? err}`);
                out.appendLine(`ERROR: ${err?.stack ?? err}`);
            }
        }),
        vscode.commands.registerCommand("d365FormExtractor.welcome", async () => {
            await openWalkthrough();
        }),
        vscode.commands.registerCommand("d365FormExtractor.scaffoldWorkspace", async () => {
            await scaffoldWorkspace();
        })
    );
}

export function deactivate() {}
