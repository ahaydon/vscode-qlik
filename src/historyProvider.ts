import * as vscode from 'vscode';
import { parseScript, ScriptSection } from './scriptModel';
import { QlikScriptFS } from './scriptFS';
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
    public readonly histUri: vscode.Uri,
    public readonly currentUri: vscode.Uri,
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
  /** Current working sections — updated on every load/revert */
  private _currentSections: ScriptSection[] = [];
  private _appId: string | undefined;
  private _client: QlikClient | undefined;

  constructor(private readonly _contentProvider: QlikHistoryContentProvider) {}

  // ── Public API ─────────────────────────────────────────────────────────

  /** Called after an app is loaded or reverted */
  loadHistory(appId: string, client: QlikClient, currentSections: ScriptSection[]): void {
    this._appId = appId;
    this._client = client;
    this._currentSections = currentSections;
    this._scriptCache.clear();
    this._versions = [];
    this._onDidChangeTreeData.fire(null);

    // Fetch in background; refresh tree when done
    client.getScriptHistory(appId).then(metas => {
      this._versions = metas.map(
        m => new HistoryVersionItem(m.scriptId!, appId, m.versionMessage ?? '', m.modifiedTime ?? '', m.size ?? 0),
      );
      this._onDidChangeTreeData.fire(null);
    }).catch(() => {
      // silently ignore — user will see empty tree
    });
  }

  /** Update the current sections reference (e.g., after a section edit) */
  setCurrentSections(sections: ScriptSection[]): void {
    this._currentSections = sections;
    // Invalidate expanded children so they are re-compared on next expand
    this._onDidChangeTreeData.fire(null);
  }

  clear(): void {
    this._versions = [];
    this._scriptCache.clear();
    this._currentSections = [];
    this._appId = undefined;
    this._client = undefined;
    this._onDidChangeTreeData.fire(null);
  }

  /** Get cached parsed sections for a scriptId, or undefined if not yet fetched */
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
    if (!element) {
      return this._versions;
    }

    if (!(element instanceof HistoryVersionItem)) {
      return [];
    }

    // Lazy-fetch the historical script
    let histSections = this._scriptCache.get(element.scriptId);
    if (!histSections) {
      if (!this._client || !this._appId) return [];
      try {
        const raw = await this._client.getScriptVersion(this._appId, element.scriptId);
        histSections = parseScript(raw);
        this._scriptCache.set(element.scriptId, histSections);
      } catch {
        return [];
      }
    }

    return this._buildSectionItems(element.scriptId, histSections);
  }

  // ── Internal ──────────────────────────────────────────────────────────

  private _buildSectionItems(scriptId: string, histSections: ScriptSection[]): HistorySectionItem[] {
    const currentMap = new Map(this._currentSections.map(s => [s.name, s]));
    const histMap = new Map(histSections.map(s => [s.name, s]));
    const items: HistorySectionItem[] = [];

    for (const hist of histSections) {
      const current = currentMap.get(hist.name);
      let status: 'modified' | 'removed';

      if (!current) {
        status = 'removed';
      } else if (hist.body.trim() !== current.body.trim()) {
        status = 'modified';
      } else {
        continue; // unchanged — skip
      }

      // Prepare content for the diff editor
      this._contentProvider.store(scriptId, hist.name, hist.body);

      // Right side: current section URI (or empty historical URI if removed)
      let currentUri: vscode.Uri;
      if (current) {
        currentUri = QlikScriptFS.uri(this._appId!, current);
      } else {
        // Section was removed — show empty right side
        this._contentProvider.store(`${scriptId}-empty`, hist.name, '');
        currentUri = QlikHistoryContentProvider.uri(`${scriptId}-empty`, hist.name);
      }

      items.push(
        new HistorySectionItem(
          hist.name,
          status,
          QlikHistoryContentProvider.uri(scriptId, hist.name),
          currentUri,
        ),
      );
    }

    // Sections that exist in current but NOT in history → 'added' (shown from history perspective)
    for (const cur of this._currentSections) {
      if (!histMap.has(cur.name)) {
        // In history view: this section didn't exist yet — mark as added (it was added after)
        this._contentProvider.store(`${scriptId}-empty`, cur.name, '');
        const emptyUri = QlikHistoryContentProvider.uri(`${scriptId}-empty`, cur.name);
        items.push(
          new HistorySectionItem(
            cur.name,
            'added',
            emptyUri,
            QlikScriptFS.uri(this._appId!, cur),
          ),
        );
      }
    }

    return items;
  }
}
