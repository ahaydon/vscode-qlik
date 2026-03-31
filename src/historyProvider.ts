import * as vscode from 'vscode';
import { parseScript, ScriptSection } from './scriptModel';
import { QlikHistoryContentProvider } from './historyContentProvider';
import type { QlikClient, ScriptMeta } from './qlikClient';

// ── Tree items ─────────────────────────────────────────────────────────────

export class HistoryVersionItem extends vscode.TreeItem {
  readonly contextValue = 'historyVersion';

  constructor(
    public readonly scriptId: string,
    public readonly appId: string,
    versionMessage: string,
    modifiedTime: string,
    size: number,
  ) {
    const label = versionMessage || '(no message)';
    super(label, vscode.TreeItemCollapsibleState.Collapsed);

    this.description = new Date(modifiedTime).toLocaleString();
    this.iconPath = new vscode.ThemeIcon('git-commit');
    this.tooltip = new vscode.MarkdownString(
      `**${label}**\n\n${this.description}\n\n${(size / 1024).toFixed(1)} KB`,
    );
  }
}

export class HistorySectionItem extends vscode.TreeItem {
  readonly contextValue = 'historySection';

  constructor(
    public readonly sectionName: string,
    public readonly status: 'modified' | 'added' | 'removed',
    /** Left side of diff — the older (before) state */
    public readonly histUri: vscode.Uri,
    /** Right side of diff — the newer (after) state */
    public readonly currentUri: vscode.Uri,
    public readonly versionLabel: string,
  ) {
    super(sectionName, vscode.TreeItemCollapsibleState.None);

    this.description = status;

    const icons: Record<typeof status, string> = {
      modified: 'diff-modified',
      added: 'diff-added',
      removed: 'diff-removed',
    };
    const colors: Record<typeof status, string> = {
      modified: 'gitDecoration.modifiedResourceForeground',
      added: 'gitDecoration.addedResourceForeground',
      removed: 'gitDecoration.deletedResourceForeground',
    };
    this.iconPath = new vscode.ThemeIcon(icons[status], new vscode.ThemeColor(colors[status]));

    this.command = {
      command: 'qlikcloud.openHistoryDiff',
      title: 'Show Diff',
      arguments: [this],
    };
  }
}

// ── Provider ───────────────────────────────────────────────────────────────

export class QlikHistoryProvider
  implements vscode.TreeDataProvider<HistoryVersionItem | HistorySectionItem>
{
  private readonly _onDidChangeTreeData = new vscode.EventEmitter<
    HistoryVersionItem | HistorySectionItem | undefined | null
  >();
  readonly onDidChangeTreeData = this._onDidChangeTreeData.event;

  private _versions: HistoryVersionItem[] = [];
  /** scriptId → parsed sections (lazy-populated on expansion) */
  private _scriptCache = new Map<string, ScriptSection[]>();
  private _appId: string | undefined;
  private _client: QlikClient | undefined;

  constructor(private readonly _contentProvider: QlikHistoryContentProvider) {}

  // ── Public API ─────────────────────────────────────────────────────────

  /** Called after an app is loaded or reverted */
  loadHistory(appId: string, client: QlikClient): void {
    this._appId = appId;
    this._client = client;
    this._scriptCache.clear();
    this._versions = [];
    this._onDidChangeTreeData.fire(null);

    client.getScriptHistory(appId).then(metas => {
      this._versions = metas.map(
        m => new HistoryVersionItem(m.scriptId!, appId, m.versionMessage ?? '', m.modifiedTime ?? '', m.size ?? 0),
      );
      this._onDidChangeTreeData.fire(null);
    }).catch(() => {});
  }

  clear(): void {
    this._versions = [];
    this._scriptCache.clear();
    this._appId = undefined;
    this._client = undefined;
    this._onDidChangeTreeData.fire(null);
  }

  /** Get cached parsed sections for a scriptId (used by revert) */
  getCachedSections(scriptId: string): ScriptSection[] | undefined {
    return this._scriptCache.get(scriptId);
  }

  // ── TreeDataProvider ───────────────────────────────────────────────────

  getTreeItem(element: HistoryVersionItem | HistorySectionItem): vscode.TreeItem {
    return element;
  }

  async getChildren(
    element?: HistoryVersionItem | HistorySectionItem,
  ): Promise<(HistoryVersionItem | HistorySectionItem)[]> {
    if (!element) return this._versions;
    if (!(element instanceof HistoryVersionItem)) return [];

    const idx = this._versions.findIndex(v => v.scriptId === element.scriptId);
    if (idx === -1 || !this._client || !this._appId) return [];

    // Fetch this version and the one before it (older) in parallel
    const prevVersion = this._versions[idx + 1]; // undefined if oldest
    const [thisSections, prevSections] = await Promise.all([
      this._fetchSections(element.scriptId),
      prevVersion ? this._fetchSections(prevVersion.scriptId) : Promise.resolve([] as ScriptSection[]),
    ]);

    const prevScriptId = prevVersion?.scriptId ?? `${element.scriptId}-empty`;

    return this._buildSectionItems(
      element.scriptId,
      thisSections,
      prevScriptId,
      prevSections,
      typeof element.label === 'string' ? element.label : '(no message)',
    );
  }

  // ── Internal ──────────────────────────────────────────────────────────

  private async _fetchSections(scriptId: string): Promise<ScriptSection[]> {
    let sections = this._scriptCache.get(scriptId);
    if (!sections) {
      const raw = await this._client!.getScriptVersion(this._appId!, scriptId);
      sections = parseScript(raw);
      this._scriptCache.set(scriptId, sections);
    }
    return sections;
  }

  private _buildSectionItems(
    scriptId: string,
    thisSections: ScriptSection[],   // newer (after) state
    prevScriptId: string,
    prevSections: ScriptSection[],   // older (before) state
    versionLabel: string,
  ): HistorySectionItem[] {
    const thisMap = new Map(thisSections.map(s => [s.name, s]));
    const prevMap = new Map(prevSections.map(s => [s.name, s]));
    const EMPTY = `${scriptId}-empty`;
    const items: HistorySectionItem[] = [];

    // Sections in this version — added or modified compared to previous
    for (const s of thisSections) {
      const prev = prevMap.get(s.name);
      if (!prev) {
        // Added in this version
        this._contentProvider.store(scriptId, s.name, s.body);
        this._contentProvider.store(EMPTY, s.name, '');
        items.push(new HistorySectionItem(
          s.name, 'added',
          QlikHistoryContentProvider.uri(EMPTY, s.name),
          QlikHistoryContentProvider.uri(scriptId, s.name),
          versionLabel,
        ));
      } else if (s.body.trim() !== prev.body.trim()) {
        // Modified in this version
        this._contentProvider.store(prevScriptId, s.name, prev.body);
        this._contentProvider.store(scriptId, s.name, s.body);
        items.push(new HistorySectionItem(
          s.name, 'modified',
          QlikHistoryContentProvider.uri(prevScriptId, s.name),
          QlikHistoryContentProvider.uri(scriptId, s.name),
          versionLabel,
        ));
      }
      // unchanged — skip
    }

    // Sections in previous version that are no longer in this version — removed
    for (const s of prevSections) {
      if (!thisMap.has(s.name)) {
        this._contentProvider.store(prevScriptId, s.name, s.body);
        this._contentProvider.store(EMPTY, s.name, '');
        items.push(new HistorySectionItem(
          s.name, 'removed',
          QlikHistoryContentProvider.uri(prevScriptId, s.name),
          QlikHistoryContentProvider.uri(EMPTY, s.name),
          versionLabel,
        ));
      }
    }

    return items;
  }
}
