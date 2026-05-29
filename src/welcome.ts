import * as vscode from "vscode";
import * as fs from "fs";
import * as path from "path";

/**
 * First-run guided setup. Walks the user through prerequisites:
 *  1) Create / open a workspace folder
 *  2) Run onboarding (Python deps, environments, output)
 *  3) Start Chrome with remote debugging + sign in to D365
 *  4) Configure form paths
 *  5) Run extraction
 */
export async function showWelcome(context: vscode.ExtensionContext): Promise<void> {
    const steps: { id: string; title: string; detail: string; action: () => Promise<void> }[] = [
        {
            id: "workspace",
            title: "1. Open a workspace folder",
            detail: "All output Excel files and validated form paths are saved in your workspace.",
            action: async () => {
                const folders = vscode.workspace.workspaceFolders;
                if (folders && folders.length > 0) {
                    vscode.window.showInformationMessage(`Workspace already open: ${folders[0].uri.fsPath}`);
                    return;
                }
                const def = "C:\\D365-Form-Extractor";
                const input = await vscode.window.showInputBox({
                    title: "Folder path to create / open",
                    value: def,
                    prompt: "A folder where extracts and config will live",
                    ignoreFocusOut: true,
                });
                if (!input) return;
                try { fs.mkdirSync(input, { recursive: true }); } catch (e) { /* ignore */ }
                try { fs.mkdirSync(path.join(input, "d365-extracts"), { recursive: true }); } catch { }
                await vscode.commands.executeCommand("vscode.openFolder", vscode.Uri.file(input), { forceNewWindow: false });
            },
        },
        {
            id: "onboard",
            title: "2. Run onboarding",
            detail: "Installs Python deps (requests, playwright, openpyxl), checks Azure CLI, registers your D365 environments.",
            action: async () => { await vscode.commands.executeCommand("d365FormExtractor.onboard"); },
        },
        {
            id: "chrome",
            title: "3. Launch Chrome (for Playwright backend)",
            detail: "Starts Chrome with remote debugging on :9222. Sign in to D365 there.",
            action: async () => { await vscode.commands.executeCommand("d365FormExtractor.startChromeCdp"); },
        },
        {
            id: "paths",
            title: "4. Configure form paths (optional but recommended)",
            detail: "Paste UI paths like 'Production control > Setup > Production journal names'. The tool will use Playwright to discover and save the menu item name.",
            action: async () => { await vscode.commands.executeCommand("d365FormExtractor.configureFormPaths"); },
        },
        {
            id: "extract",
            title: "5. Extract a form",
            detail: "Pick a form path (or enter menu item manually) and the environments to compare.",
            action: async () => { await vscode.commands.executeCommand("d365FormExtractor.extract"); },
        },
    ];

    while (true) {
        const items: (vscode.QuickPickItem & { run?: () => Promise<void> })[] = steps.map(s => ({
            label: s.title,
            detail: s.detail,
            run: s.action,
        }));
        items.push({ label: "$(check) Close", run: undefined });
        const pick = await vscode.window.showQuickPick(items, {
            title: "D365 FO Configuration Comparator — Getting Started",
            placeHolder: "Pick a step to run, or Close",
            ignoreFocusOut: true,
            matchOnDetail: true,
        });
        if (!pick || !pick.run) break;
        try { await pick.run(); } catch (e: any) {
            vscode.window.showErrorMessage(`Step failed: ${e?.message ?? e}`);
        }
    }
    await context.globalState.update("welcomeShown", true);
}
