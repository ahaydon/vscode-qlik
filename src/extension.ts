import * as vscode from 'vscode';
import { loadContexts, QlikContext, SavedOauthClient, OAUTH_CLIENTS_KEY } from './contexts';
import { buildOauthHostConfig, clearOauthTokens, OAUTH_REDIRECT_URI } from './oauthFlow';
import { createClient, QlikClient, QlikApp } from './qlikClient';
import type { Space } from '@qlik/api/spaces';
import { parseScript, serializeScript } from './scriptModel';
import { QlikScriptFS } from './scriptFS';
import { QlikScriptTreeProvider, SectionItem } from './treeProvider';
import { QlikHistoryContentProvider } from './historyContentProvider';
import { QlikHistoryProvider, HistoryVersionItem, HistorySectionItem } from './historyProvider';
import { QlikCompletionProvider, QlikHoverProvider, validateSections, validateDocument } from './languageProvider';
import { QlikReloadLogProvider, ReloadLogItem } from './reloadLogProvider';
import { QlikReloadLogContentProvider } from './reloadLogContentProvider';

// ── Module-level state ─────────────────────────────────────────────────────

let extCtx: vscode.ExtensionContext | undefined;
let currentContext: QlikContext | undefined;
let currentClient: QlikClient | undefined;
let currentAppId: string | undefined;

const scriptFS = new QlikScriptFS();
const treeProvider = new QlikScriptTreeProvider();
const historyContentProvider = new QlikHistoryContentProvider();
const historyProvider = new QlikHistoryProvider(historyContentProvider);
const reloadLogContentProvider = new QlikReloadLogContentProvider();
const reloadLogProvider = new QlikReloadLogProvider();
const reloadOutput = vscode.window.createOutputChannel('Qlik Cloud Reload', 'log');
const diagCollection = vscode.languages.createDiagnosticCollection('qlikscript');

// ── Activation ─────────────────────────────────────────────────────────────

export function activate(context: vscode.ExtensionContext): void {
  extCtx = context;

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

  // Register reload log content provider (read-only, for opening log text)
  context.subscriptions.push(
    vscode.workspace.registerTextDocumentContentProvider(
      QlikReloadLogContentProvider.SCHEME,
      reloadLogContentProvider,
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

  const reloadLogView = vscode.window.createTreeView('qlikReloadLogs', {
    treeDataProvider: reloadLogProvider,
    showCollapseAll: false,
  });
  context.subscriptions.push(reloadLogView);

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
    vscode.commands.registerCommand('qlikcloud.refreshReloadLogs', cmdRefreshReloadLogs),
    vscode.commands.registerCommand('qlikcloud.showReloadSummary', cmdShowReloadSummary),
    vscode.commands.registerCommand('qlikcloud.openReloadLog', cmdOpenReloadLog),
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

  context.subscriptions.push(
    vscode.commands.registerCommand('_qlikcloud.refreshStatus', refreshStatus),
  );

  // Initial state
  vscode.commands.executeCommand('setContext', 'qlikcloud.appLoaded', false);

  // Expose refreshStatus so commands can call it
  (globalThis as Record<string, unknown>).__qlikRefreshStatus = refreshStatus;
}

export function deactivate(): void {}

// ── Command implementations ────────────────────────────────────────────────

// ── OAuth client helpers ─────────────────────────────────────────────────

function getSavedOauthClients(): SavedOauthClient[] {
  if (!extCtx) return [];
  return extCtx.globalState.get<SavedOauthClient[]>(OAUTH_CLIENTS_KEY, []);
}

async function saveOauthClients(clients: SavedOauthClient[]): Promise<void> {
  if (!extCtx) return;
  await extCtx.globalState.update(OAUTH_CLIENTS_KEY, clients);
}

function normalizeHost(input: string): string {
  return input.trim().replace(/^https?:\/\//i, '').replace(/\/+$/, '');
}

async function promptForOauthClient(existingNames: Set<string>): Promise<SavedOauthClient | undefined> {
  const host = await vscode.window.showInputBox({
    title: 'Add OAuth Client (1/3) — Tenant host',
    prompt: `Tenant hostname (e.g. my-tenant.region.qlikcloud.com). Register redirect URI ${OAUTH_REDIRECT_URI} in your Qlik OAuth client.`,
    placeHolder: 'my-tenant.region.qlikcloud.com',
    validateInput: v => (normalizeHost(v) ? undefined : 'Host cannot be empty'),
  });
  if (!host) return undefined;

  const clientId = await vscode.window.showInputBox({
    title: 'Add OAuth Client (2/3) — Client ID',
    prompt: 'OAuth client ID (from the Qlik tenant admin)',
    validateInput: v => (v.trim() ? undefined : 'Client ID cannot be empty'),
  });
  if (!clientId) return undefined;

  const normalizedHost = normalizeHost(host);
  const defaultName = normalizedHost.split('.')[0] || normalizedHost;
  const name = await vscode.window.showInputBox({
    title: 'Add OAuth Client (3/3) — Display name',
    prompt: 'A friendly name for this OAuth context',
    value: defaultName,
    validateInput: v => {
      const t = v.trim();
      if (!t) return 'Name cannot be empty';
      if (existingNames.has(t)) return `A context named "${t}" already exists`;
      return undefined;
    },
  });
  if (!name) return undefined;

  return { name: name.trim(), host: normalizedHost, clientId: clientId.trim() };
}

const REMOVE_BUTTON: vscode.QuickInputButton = {
  iconPath: new vscode.ThemeIcon('trash'),
  tooltip: 'Remove this OAuth client',
};

interface ContextPickItem extends vscode.QuickPickItem {
  ctx?: QlikContext;
  action?: 'add';
}

async function cmdSelectContext(): Promise<void> {
  if (!extCtx) return;

  const buildItems = (): ContextPickItem[] => {
    const saved = getSavedOauthClients();
    const { contexts, currentContext: defaultCtx } = loadContexts(saved);

    const yaml = contexts.filter(c => c.source === 'yaml');
    const oauth = contexts.filter(c => c.source === 'oauth');

    const items: ContextPickItem[] = [];

    if (yaml.length > 0) {
      items.push({ label: 'From ~/.qlik/contexts.yml', kind: vscode.QuickPickItemKind.Separator });
      for (const c of yaml) {
        items.push({
          label: c.name,
          description: c.server,
          detail: c.serverType === 'cloud' ? '$(cloud) Cloud' : `$(server) ${c.serverType}`,
          picked: c.name === defaultCtx,
          ctx: c,
        });
      }
    }

    if (oauth.length > 0) {
      items.push({ label: 'OAuth clients', kind: vscode.QuickPickItemKind.Separator });
      for (const c of oauth) {
        items.push({
          label: c.name,
          description: c.server,
          detail: '$(key) OAuth interactive',
          ctx: c,
          buttons: [REMOVE_BUTTON],
        });
      }
    }

    items.push({ label: '', kind: vscode.QuickPickItemKind.Separator });
    items.push({
      label: '$(add) Add OAuth client…',
      detail: 'Authenticate with a Qlik tenant via browser consent',
      action: 'add',
    });

    return items;
  };

  const qp = vscode.window.createQuickPick<ContextPickItem>();
  qp.title = 'Select Qlik Context';
  qp.placeholder = 'Choose a context, or add a new OAuth client';
  qp.items = buildItems();

  const selected = await new Promise<ContextPickItem | undefined>(resolve => {
    qp.onDidTriggerItemButton(async e => {
      if (e.button !== REMOVE_BUTTON || !e.item.ctx || e.item.ctx.source !== 'oauth') return;
      const confirm = await vscode.window.showWarningMessage(
        `Remove OAuth client "${e.item.ctx.name}"? Cached tokens will be deleted.`,
        { modal: true },
        'Remove',
      );
      if (confirm !== 'Remove') return;
      const saved = getSavedOauthClients();
      const removed = saved.find(c => c.name === e.item.ctx!.name);
      const next = saved.filter(c => c.name !== e.item.ctx!.name);
      await saveOauthClients(next);
      if (removed && extCtx) await clearOauthTokens(removed, extCtx.secrets);
      qp.items = buildItems();
    });
    qp.onDidAccept(() => {
      resolve(qp.selectedItems[0]);
      qp.hide();
    });
    qp.onDidHide(() => resolve(undefined));
    qp.show();
  });

  if (!selected) return;

  if (selected.action === 'add') {
    const existingNames = new Set([
      ...loadContexts(getSavedOauthClients()).contexts.map(c => c.name),
    ]);
    const newClient = await promptForOauthClient(existingNames);
    if (!newClient) return;
    const next = [...getSavedOauthClients(), newClient];
    await saveOauthClients(next);
    await connectToOauthClient(newClient);
    return;
  }

  if (!selected.ctx) return;

  if (selected.ctx.source === 'oauth') {
    const saved = getSavedOauthClients().find(c => c.name === selected.ctx!.name);
    if (!saved) return;
    await connectToOauthClient(saved);
    return;
  }

  currentContext = selected.ctx;
  currentClient = createClient(selected.ctx.hostConfig);
  triggerStatusRefresh();
  vscode.window.showInformationMessage(`Connected to: ${selected.ctx.server}`);
}

async function connectToOauthClient(client: SavedOauthClient): Promise<void> {
  if (!extCtx) return;
  const hostConfig = buildOauthHostConfig(client, extCtx.secrets);
  const ctx: QlikContext = {
    name: client.name,
    server: client.host,
    serverType: 'cloud',
    source: 'oauth',
    clientId: client.clientId,
    hostConfig,
  };
  const probe = createClient(hostConfig);
  try {
    await probe.getSpaces();
  } catch (err) {
    vscode.window.showErrorMessage(`OAuth login failed: ${(err as Error).message}`);
    return;
  }
  currentContext = ctx;
  currentClient = probe;
  triggerStatusRefresh();
  vscode.window.showInformationMessage(`Connected to: ${client.host}`);
}

function triggerStatusRefresh(): void {
  const fn = (globalThis as Record<string, unknown>).__qlikRefreshStatus as (() => void) | undefined;
  fn?.();
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

  // Load history and reload logs in background
  historyProvider.loadHistory(app.id, currentClient!);
  reloadLogProvider.loadLogs(app.id, currentClient!);

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
    await vscode.commands.executeCommand('vscode.open', uri, { preview: true }, sections[0].name);
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
  await vscode.commands.executeCommand('vscode.open', uri, { preview: true, preserveFocus: false }, item.section.name);
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
  await vscode.commands.executeCommand('vscode.open', uri, { preview: false }, section.name);

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
        reloadLogProvider.loadLogs(currentAppId!, currentClient!);
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

async function cmdRefreshReloadLogs(): Promise<void> {
  if (!currentClient || !currentAppId) {
    vscode.window.showErrorMessage('No app loaded.');
    return;
  }
  reloadLogProvider.loadLogs(currentAppId, currentClient);
}

function cmdShowReloadSummary(item: ReloadLogItem): void {
  reloadOutput.clear();
  reloadOutput.show(/* preserveFocus */ true);
  if (item.summary) {
    reloadOutput.append(item.summary);
  } else {
    reloadOutput.appendLine('No summary available for this reload.');
  }
}

async function cmdOpenReloadLog(item: ReloadLogItem): Promise<void> {
  if (!currentClient || !currentAppId) {
    vscode.window.showErrorMessage('No app loaded.');
    return;
  }

  if (!item.reloadId) {
    vscode.window.showErrorMessage('No reload ID available for this log entry.');
    return;
  }

  let content = reloadLogContentProvider.get(item.reloadId);
  if (!content) {
    await vscode.window.withProgress(
      { location: vscode.ProgressLocation.Notification, title: 'Downloading reload log…', cancellable: false },
      async () => {
        content = await currentClient!.getReloadLog(currentAppId!, item.reloadId);
        reloadLogContentProvider.store(item.reloadId, content!);
      },
    );
  }

  const tabLabel = typeof item.label === 'string' ? item.label : 'Reload Log';
  const uri = QlikReloadLogContentProvider.uri(item.reloadId);
  await vscode.workspace.openTextDocument(uri);
  await vscode.commands.executeCommand('vscode.open', uri, { preview: false }, tabLabel);
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
    await vscode.commands.executeCommand('vscode.open', uri, { preview: true }, sections[0].name);
  }

  vscode.window.showInformationMessage(`Reverted to "${label}" — save to push to Qlik Cloud.`);
}
