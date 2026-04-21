import * as vscode from 'vscode';
import { loadContexts, QlikContext } from './contexts';
import { createClient, QlikClient, QlikApp } from './qlikClient';
import type { Space } from '@qlik/api/spaces';
import { parseScript, serializeScript } from './scriptModel';
import { QlikScriptFS } from './scriptFS';
import { QlikScriptTreeProvider, SectionItem } from './treeProvider';
import { QlikHistoryContentProvider } from './historyContentProvider';
import { QlikHistoryProvider, HistoryVersionItem, HistorySectionItem } from './historyProvider';
import { QlikCompletionProvider, QlikHoverProvider, validateSections, validateDocument } from './languageProvider';

// ── Module-level state ─────────────────────────────────────────────────────

let currentContext: QlikContext | undefined;
let currentClient: QlikClient | undefined;
let currentAppId: string | undefined;

const scriptFS = new QlikScriptFS();
const treeProvider = new QlikScriptTreeProvider();
const historyContentProvider = new QlikHistoryContentProvider();
const historyProvider = new QlikHistoryProvider(historyContentProvider);
const reloadOutput = vscode.window.createOutputChannel('Qlik Cloud Reload', 'log');
const diagCollection = vscode.languages.createDiagnosticCollection('qlikscript');

// ── Activation ─────────────────────────────────────────────────────────────

export function activate(context: vscode.ExtensionContext): void {
  // Register virtual filesystem
  context.subscriptions.push(
    vscode.workspace.registerFileSystemProvider(QlikScriptFS.SCHEME, scriptFS, {
      isCaseSensitive: true,
      isReadonly: false,
    }),
  );

  // Register history content provider (read-only, for diff editor left side)
  context.subscriptions.push(
    vscode.workspace.registerTextDocumentContentProvider(
      QlikHistoryContentProvider.SCHEME,
      historyContentProvider,
    ),
  );

  // Register tree views
  const treeView = vscode.window.createTreeView('qlikScriptSections', {
    treeDataProvider: treeProvider,
    showCollapseAll: false,
  });
  context.subscriptions.push(treeView);

  const historyView = vscode.window.createTreeView('qlikScriptHistory', {
    treeDataProvider: historyProvider,
    showCollapseAll: false,
  });
  context.subscriptions.push(historyView);

  // Language intelligence: completions, hover docs, diagnostics
  context.subscriptions.push(
    diagCollection,
    vscode.languages.registerCompletionItemProvider(
      { language: 'qlikscript' },
      new QlikCompletionProvider(treeProvider),
      '$',
    ),
    vscode.languages.registerHoverProvider(
      { language: 'qlikscript' },
      new QlikHoverProvider(),
    ),
    // Validate local .qvs files on open and change
    vscode.workspace.onDidOpenTextDocument(doc => {
      if (doc.languageId === 'qlikscript' && doc.uri.scheme === 'file') {
        validateDocument(doc, diagCollection);
      }
    }),
    vscode.workspace.onDidChangeTextDocument(e => {
      if (e.document.languageId === 'qlikscript' && e.document.uri.scheme === 'file') {
        validateDocument(e.document, diagCollection, 500);
      }
    }),
    vscode.workspace.onDidCloseTextDocument(doc => {
      if (doc.uri.scheme === 'file') diagCollection.delete(doc.uri);
    }),
  );

  // Update tree title to show current context / app
  const updateTitle = () => {
    const parts: string[] = [];
    if (currentContext) parts.push(currentContext.name);
    const appName = treeProvider.getAppName();
    if (appName) parts.push(appName);
    treeView.title = parts.length ? `Script: ${parts.join(' › ')}` : 'Script Sections';
    if (treeProvider.isDirty()) treeView.badge = { value: 1, tooltip: 'Unsaved changes' };
    else treeView.badge = undefined;
  };

  // Listen for section file saves → mark dirty
  context.subscriptions.push(
    scriptFS.onSectionWritten(uri => {
      // URI: qlikscript://<sectionId>/<appId>/<name>.qvs
      const parts = uri.path.split('/').filter(Boolean);
      const appId = parts[0];
      const sectionId = uri.authority;
      if (sectionId) {
        treeProvider.markDirty(sectionId);
        // Sync the in-memory body back into the section store
        const body = scriptFS.readSectionBody(appId, sectionId);
        const sections = treeProvider.getSections();
        const s = sections.find(x => x.id === sectionId);
        if (s && body !== undefined) s.body = body;
      }
      if (appId) validateSections(treeProvider.getSections(), scriptFS, appId, diagCollection, 500);
      updateTitle();
      vscode.commands.executeCommand('setContext', 'qlikcloud.appLoaded', true);
    }),
  );

  // ── Commands ─────────────────────────────────────────────────────────────

  context.subscriptions.push(
    vscode.commands.registerCommand('qlikcloud.selectContext', cmdSelectContext),
    vscode.commands.registerCommand('qlikcloud.openApp', cmdOpenApp),
    vscode.commands.registerCommand('qlikcloud.saveScript', cmdSaveScript),
    vscode.commands.registerCommand('qlikcloud.openSection', cmdOpenSection),
    vscode.commands.registerCommand('qlikcloud.addSection', cmdAddSection),
    vscode.commands.registerCommand('qlikcloud.deleteSection', cmdDeleteSection),
    vscode.commands.registerCommand('qlikcloud.moveSectionUp', cmdMoveUp),
    vscode.commands.registerCommand('qlikcloud.moveSectionDown', cmdMoveDown),
    vscode.commands.registerCommand('qlikcloud.renameSection', cmdRenameSection),
    vscode.commands.registerCommand('qlikcloud.refreshHistory', cmdRefreshHistory),
    vscode.commands.registerCommand('qlikcloud.revertToVersion', cmdRevertToVersion),
    vscode.commands.registerCommand('qlikcloud.openHistoryDiff', cmdOpenHistoryDiff),
    vscode.commands.registerCommand('qlikcloud.reloadApp', cmdReloadApp),
  );

  // Status-bar item
  const statusBar = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 10);
  statusBar.command = 'qlikcloud.selectContext';
  statusBar.tooltip = 'Click to select Qlik context';
  statusBar.text = '$(account) Qlik: (no context)';
  statusBar.show();
  context.subscriptions.push(statusBar);

  const refreshStatus = () => {
    statusBar.text = currentContext
      ? `$(account) Qlik: ${currentContext.name}`
      : '$(account) Qlik: (no context)';
    updateTitle();
  };

  // Wrap select-context to also refresh status
  const origSelect = cmdSelectContext;
  context.subscriptions.push(
    vscode.commands.registerCommand('_qlikcloud.refreshStatus', refreshStatus),
  );

  // Initial state
  vscode.commands.executeCommand('setContext', 'qlikcloud.appLoaded', false);

  // ── Helper: refresh status after context change ──────────────────────────
  // We patch statusBar update into the command registration below.
  // (The commands defined above capture a closure; we update statusBar in the
  //  selectContext implementation via the module-level currentContext ref.)

  // Expose refreshStatus so commands can call it
  (globalThis as Record<string, unknown>).__qlikRefreshStatus = refreshStatus;
}

export function deactivate(): void {}

// ── Command implementations ────────────────────────────────────────────────

async function cmdSelectContext(): Promise<void> {
  const { contexts, currentContext: defaultCtx } = loadContexts();
  if (contexts.length === 0) {
    vscode.window.showErrorMessage('No Qlik contexts found in ~/.qlik/contexts.yml');
    return;
  }

  const items: vscode.QuickPickItem[] = contexts.map(c => ({
    label: c.name,
    description: c.server,
    detail: c.serverType === 'cloud' ? '$(cloud) Cloud' : `$(server) ${c.serverType}`,
    picked: c.name === defaultCtx,
  }));

  const picked = await vscode.window.showQuickPick(items, {
    title: 'Select Qlik Context',
    placeHolder: 'Choose a context from ~/.qlik/contexts.yml',
  });

  if (!picked) return;

  currentContext = contexts.find(c => c.name === picked.label)!;
  currentClient = createClient(currentContext.hostConfig);

  const refresh = (globalThis as Record<string, unknown>).__qlikRefreshStatus as (() => void) | undefined;
  refresh?.();

  vscode.window.showInformationMessage(`Connected to: ${currentContext.server}`);
}

async function cmdOpenApp(): Promise<void> {
  if (!currentClient || !currentContext) {
    const select = 'Select Context';
    const choice = await vscode.window.showErrorMessage('No Qlik context selected.', select);
    if (choice === select) await cmdSelectContext();
    return;
  }

  // Step 1: pick spaces
  let spaces: Space[] = [];
  await vscode.window.withProgress(
    { location: vscode.ProgressLocation.Notification, title: 'Loading spaces…', cancellable: false },
    async () => {
      spaces = await currentClient!.getSpaces();
    },
  );

  const spaceItems: vscode.QuickPickItem[] = [
    { label: 'All spaces', description: 'Search across all accessible spaces', picked: true },
    { label: '── Personal ──', description: 'Apps not in any shared space', kind: vscode.QuickPickItemKind.Separator },
    { label: 'Personal', description: 'My personal space' },
    { label: '── Shared & Managed ──', kind: vscode.QuickPickItemKind.Separator },
    ...spaces
      .filter(s => s.type === 'shared' || s.type === 'managed' || s.type === 'data')
      .map(s => ({ label: s.name, description: `${s.type} · ${s.id}` })),
  ];

  const pickedSpaces = await vscode.window.showQuickPick(spaceItems, {
    title: 'Filter by Space (optional)',
    placeHolder: 'Select spaces to search in, or pick "All spaces"',
    canPickMany: true,
  });

  if (pickedSpaces === undefined) return; // cancelled

  // Resolve space IDs
  let spaceIds: string[] = [];
  const allSelected = pickedSpaces.some(p => p.label === 'All spaces') || pickedSpaces.length === 0;

  if (!allSelected) {
    for (const p of pickedSpaces) {
      if (p.label === 'Personal') {
        spaceIds.push('__personal__');
      } else if (p.description) {
        const match = p.description.match(/·\s+(\S+)$/);
        if (match) spaceIds.push(match[1]);
      }
    }
  }

  // Step 2: search apps with live filter
  const qp = vscode.window.createQuickPick<vscode.QuickPickItem & { app?: QlikApp }>();
  qp.title = 'Open Qlik App';
  qp.placeholder = 'Type to search app names…';
  qp.busy = false;

  let debounceTimer: ReturnType<typeof setTimeout> | undefined;

  const loadApps = (query: string) => {
    qp.busy = true;
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(async () => {
      try {
        const apps = await currentClient!.searchApps(spaceIds, query);
        qp.items = apps.map(app => ({
          label: app.name,
          description: app.spaceId ? `Space: ${app.spaceId.slice(0, 8)}…` : 'Personal',
          detail: app.updatedAt ? `Updated: ${new Date(app.updatedAt).toLocaleDateString()}` : undefined,
          app,
        }));
      } catch (err) {
        vscode.window.showErrorMessage(`Failed to search apps: ${(err as Error).message}`);
      } finally {
        qp.busy = false;
      }
    }, 400);
  };

  // Initial load
  loadApps('');

  qp.onDidChangeValue(v => loadApps(v));

  const selected = await new Promise<(vscode.QuickPickItem & { app?: QlikApp }) | undefined>(resolve => {
    qp.onDidAccept(() => { resolve(qp.selectedItems[0]); qp.hide(); });
    qp.onDidHide(() => resolve(undefined));
    qp.show();
  });

  if (!selected?.app) return;

  await loadAppScript(selected.app);
}

async function loadAppScript(app: QlikApp): Promise<void> {
  let rawScript = '';

  await vscode.window.withProgress(
    { location: vscode.ProgressLocation.Notification, title: `Loading script: ${app.name}…`, cancellable: false },
    async () => {
      rawScript = await currentClient!.getCurrentScript(app.id);
    },
  );

  const sections = parseScript(rawScript);
  currentAppId = app.id;

  // Populate virtual filesystem
  scriptFS.populateSections(app.id, sections);

  // Update tree
  treeProvider.setApp(app.id, app.name, sections);
  validateSections(sections, scriptFS, app.id, diagCollection);
  vscode.commands.executeCommand('setContext', 'qlikcloud.appLoaded', true);

  // Load history in background
  historyProvider.loadHistory(app.id, currentClient!);

  const refresh = (globalThis as Record<string, unknown>).__qlikRefreshStatus as (() => void) | undefined;
  refresh?.();

  vscode.window.showInformationMessage(
    `Loaded "${app.name}" — ${sections.length} section${sections.length !== 1 ? 's' : ''}`,
  );

  // Open the first section automatically
  if (sections.length > 0) {
    const uri = QlikScriptFS.uri(app.id, sections[0]);
    const doc = await vscode.workspace.openTextDocument(uri);
    await vscode.languages.setTextDocumentLanguage(doc, 'qlikscript');
    await vscode.window.showTextDocument(doc, { preview: false });
  }
}

async function cmdSaveScript(): Promise<void> {
  if (!currentClient || !currentAppId) {
    vscode.window.showErrorMessage('No app loaded.');
    return;
  }

  const sections = treeProvider.getSections();
  // Sync any open editors back to section store before serializing
  for (const s of sections) {
    const uri = QlikScriptFS.uri(currentAppId, s);
    const doc = vscode.workspace.textDocuments.find(d => d.uri.toString() === uri.toString());
    if (doc && !doc.isUntitled) {
      s.body = doc.getText();
    }
  }

  const script = serializeScript(sections);
  const appName = treeProvider.getAppName() ?? currentAppId;

  const msg = await vscode.window.showInputBox({
    title: 'Save Script to Qlik Cloud',
    prompt: 'Version message (optional)',
    value: `Updated via VS Code — ${new Date().toISOString().slice(0, 16).replace('T', ' ')}`,
  });

  if (msg === undefined) return; // cancelled

  await vscode.window.withProgress(
    { location: vscode.ProgressLocation.Notification, title: `Saving "${appName}"…`, cancellable: false },
    async () => {
      await currentClient!.saveScript(currentAppId!, script, msg);
    },
  );

  treeProvider.markClean();

  const refresh = (globalThis as Record<string, unknown>).__qlikRefreshStatus as (() => void) | undefined;
  refresh?.();

  vscode.window.showInformationMessage(`Script saved to Qlik Cloud: "${appName}"`);
}

async function cmdOpenSection(item: SectionItem): Promise<void> {
  const uri = QlikScriptFS.uri(item.appId, item.section);
  const doc = await vscode.workspace.openTextDocument(uri);
  await vscode.languages.setTextDocumentLanguage(doc, 'qlikscript');
  await vscode.window.showTextDocument(doc, { preview: false, preserveFocus: false });
}

async function cmdAddSection(): Promise<void> {
  if (!currentAppId) return;

  const name = await vscode.window.showInputBox({
    title: 'New Script Section',
    prompt: 'Section name',
    validateInput: v => (v.trim() ? undefined : 'Name cannot be empty'),
  });

  if (!name) return;

  const section = treeProvider.addSection(name.trim());
  scriptFS.writeSectionBody(currentAppId, section.id, '');

  const uri = QlikScriptFS.uri(currentAppId, section);
  const doc = await vscode.workspace.openTextDocument(uri);
  await vscode.languages.setTextDocumentLanguage(doc, 'qlikscript');
  await vscode.window.showTextDocument(doc, { preview: false });

  treeProvider.markDirty(section.id);
  const refresh = (globalThis as Record<string, unknown>).__qlikRefreshStatus as (() => void) | undefined;
  refresh?.();
}

async function cmdDeleteSection(item: SectionItem): Promise<void> {
  if (!currentAppId) return;

  const sections = treeProvider.getSections();
  if (sections.length <= 1) {
    vscode.window.showErrorMessage('Cannot delete the last section.');
    return;
  }

  const confirm = await vscode.window.showWarningMessage(
    `Delete section "${item.section.name}"?`,
    { modal: true },
    'Delete',
  );

  if (confirm !== 'Delete') return;

  // Close any open editor for this section
  const uri = QlikScriptFS.uri(item.appId, item.section);
  const doc = vscode.workspace.textDocuments.find(d => d.uri.toString() === uri.toString());
  if (doc) {
    await vscode.window.showTextDocument(doc, { preview: true, preserveFocus: false });
    await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
  }

  scriptFS.removeSections(item.appId, [item.section.id]);
  treeProvider.deleteSection(item.section.id);
  treeProvider.markDirty('__deleted__'); // mark dirty without a real id

  const refresh = (globalThis as Record<string, unknown>).__qlikRefreshStatus as (() => void) | undefined;
  refresh?.();
}

function cmdMoveUp(item: SectionItem): void {
  treeProvider.moveUp(item.section.id);
  treeProvider.markDirty(item.section.id);
  const refresh = (globalThis as Record<string, unknown>).__qlikRefreshStatus as (() => void) | undefined;
  refresh?.();
}

function cmdMoveDown(item: SectionItem): void {
  treeProvider.moveDown(item.section.id);
  treeProvider.markDirty(item.section.id);
  const refresh = (globalThis as Record<string, unknown>).__qlikRefreshStatus as (() => void) | undefined;
  refresh?.();
}

async function cmdRenameSection(item: SectionItem): Promise<void> {
  if (!currentAppId) return;

  const newName = await vscode.window.showInputBox({
    title: 'Rename Section',
    value: item.section.name,
    validateInput: v => (v.trim() ? undefined : 'Name cannot be empty'),
  });

  if (!newName || newName.trim() === item.section.name) return;

  treeProvider.renameSection(item.section.id, newName.trim());
  treeProvider.markDirty(item.section.id);

  const refresh = (globalThis as Record<string, unknown>).__qlikRefreshStatus as (() => void) | undefined;
  refresh?.();
}

async function cmdReloadApp(): Promise<void> {
  if (!currentClient || !currentAppId) {
    vscode.window.showErrorMessage('No app loaded.');
    return;
  }

  const appName = treeProvider.getAppName() ?? currentAppId;

  reloadOutput.clear();
  reloadOutput.show(/* preserveFocus */ true);
  reloadOutput.appendLine(`Reloading "${appName}" …`);
  reloadOutput.appendLine(`Started: ${new Date().toLocaleString()}`);
  reloadOutput.appendLine('');

  let cancelled = false;

  await vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: `Reloading "${appName}"`,
      cancellable: true,
    },
    async (_progress, token) => {
      token.onCancellationRequested(() => {
        cancelled = true;
        reloadOutput.appendLine('\n[Cancelling…]');
      });

      let result: 'succeeded' | 'failed' | 'cancelled';
      try {
        result = await currentClient!.reloadApp(
          currentAppId!,
          chunk => reloadOutput.append(chunk),
          () => cancelled,
        );
      } catch (err) {
        reloadOutput.appendLine(`\nERROR: ${(err as Error).message}`);
        vscode.window.showErrorMessage(`Reload error: ${(err as Error).message}`);
        return;
      }

      reloadOutput.appendLine('');
      reloadOutput.appendLine(`Finished: ${new Date().toLocaleString()}`);
      reloadOutput.appendLine(`Status: ${result.toUpperCase()}`);

      if (result === 'succeeded') {
        vscode.window.showInformationMessage(`Reload succeeded: "${appName}"`);
        historyProvider.loadHistory(currentAppId!, currentClient!);
      } else if (result === 'cancelled') {
        vscode.window.showWarningMessage(`Reload cancelled: "${appName}"`);
      } else {
        vscode.window.showErrorMessage(`Reload failed: "${appName}"`);
      }
    },
  );
}

async function cmdRefreshHistory(): Promise<void> {
  if (!currentClient || !currentAppId) {
    vscode.window.showErrorMessage('No app loaded.');
    return;
  }
  historyProvider.loadHistory(currentAppId, currentClient);
}

async function cmdOpenHistoryDiff(item: HistorySectionItem): Promise<void> {
  const title = `${item.sectionName} (${item.versionLabel})`;
  await vscode.commands.executeCommand('vscode.diff', item.histUri, item.currentUri, title);
}

async function cmdRevertToVersion(item: HistoryVersionItem): Promise<void> {
  if (!currentClient || !currentAppId) return;

  const label = typeof item.label === 'string' ? item.label : '(no message)';
  const confirm = await vscode.window.showWarningMessage(
    `Revert to "${label}"? Current unsaved changes will be lost.`,
    { modal: true },
    'Revert',
  );
  if (confirm !== 'Revert') return;

  // Use cached sections if available, otherwise fetch
  let rawScript: string | undefined;
  const cached = historyProvider.getCachedSections(item.scriptId);
  if (cached) {
    // Re-serialize from cached sections to get raw script
    const { serializeScript } = await import('./scriptModel');
    rawScript = serializeScript(cached);
  }

  if (!rawScript) {
    await vscode.window.withProgress(
      { location: vscode.ProgressLocation.Notification, title: 'Loading version…', cancellable: false },
      async () => {
        rawScript = await currentClient!.getScriptVersion(currentAppId!, item.scriptId);
      },
    );
  }

  const sections = parseScript(rawScript!);
  const appName = treeProvider.getAppName() ?? currentAppId;

  // Close any open qlikscript:// editors
  for (const doc of vscode.workspace.textDocuments) {
    if (doc.uri.scheme === QlikScriptFS.SCHEME) {
      await vscode.window.showTextDocument(doc, { preview: true, preserveFocus: false });
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
    }
  }

  scriptFS.populateSections(currentAppId!, sections);
  treeProvider.setApp(currentAppId!, appName!, sections);
  treeProvider.markDirty('__reverted__');
  validateSections(sections, scriptFS, currentAppId!, diagCollection);
  historyProvider.loadHistory(currentAppId!, currentClient!);

  vscode.commands.executeCommand('setContext', 'qlikcloud.appLoaded', true);

  const refresh = (globalThis as Record<string, unknown>).__qlikRefreshStatus as (() => void) | undefined;
  refresh?.();

  // Open first section
  if (sections.length > 0) {
    const uri = QlikScriptFS.uri(currentAppId!, sections[0]);
    const doc = await vscode.workspace.openTextDocument(uri);
    await vscode.languages.setTextDocumentLanguage(doc, 'qlikscript');
    await vscode.window.showTextDocument(doc, { preview: false });
  }

  vscode.window.showInformationMessage(`Reverted to "${label}" — save to push to Qlik Cloud.`);
}
