import * as vscode from 'vscode';
import type { QlikClient, Reload } from './qlikClient';

// ── Helpers ────────────────────────────────────────────────────────────────

function statusLabel(status: Reload['status']): string {
  switch (status) {
    case 'SUCCEEDED': return 'Succeeded';
    case 'FAILED': return 'Failed';
    case 'QUEUED': return 'Queued';
    case 'RELOADING': return 'Reloading\u2026';
    case 'CANCELING': return 'Canceling\u2026';
    case 'CANCELED': return 'Canceled';
    case 'EXCEEDED_LIMIT': return 'Limit exceeded';
    default: return status;
  }
}

function statusIcon(status: Reload['status']): vscode.ThemeIcon {
  switch (status) {
    case 'SUCCEEDED':
      return new vscode.ThemeIcon('pass-filled', new vscode.ThemeColor('testing.iconPassed'));
    case 'FAILED':
      return new vscode.ThemeIcon('error', new vscode.ThemeColor('testing.iconFailed'));
    case 'RELOADING':
    case 'CANCELING':
      return new vscode.ThemeIcon('sync~spin');
    case 'QUEUED':
      return new vscode.ThemeIcon('clock');
    case 'CANCELED':
      return new vscode.ThemeIcon('circle-slash', new vscode.ThemeColor('disabledForeground'));
    case 'EXCEEDED_LIMIT':
      return new vscode.ThemeIcon('warning', new vscode.ThemeColor('list.warningForeground'));
    default:
      return new vscode.ThemeIcon('circle-outline');
  }
}

function formatDuration(startTime: string, endTime: string): string {
  const ms = new Date(endTime).getTime() - new Date(startTime).getTime();
  if (ms < 1000) return `${ms}ms`;
  if (ms < 60_000) return `${(ms / 1000).toFixed(1)}s`;
  const m = Math.floor(ms / 60_000);
  const s = Math.round((ms % 60_000) / 1000);
  return `${m}m ${s}s`;
}

// ── Tree item ──────────────────────────────────────────────────────────────

export class ReloadLogItem extends vscode.TreeItem {
  readonly reloadId: string;
  readonly summary: string | undefined;

  constructor(reload: Reload, hasLog: boolean) {
    const timeStr = reload.endTime ?? reload.creationTime;
    const label = timeStr ? new Date(timeStr).toLocaleString() : reload.id;

    super(label, vscode.TreeItemCollapsibleState.None);

    this.reloadId = reload.id;
    this.summary = reload.log;
    this.contextValue = hasLog ? 'reloadLogWithLog' : 'reloadLog';
    this.description = statusLabel(reload.status);
    this.iconPath = statusIcon(reload.status);

    const tooltipLines = [
      `**Status:** ${statusLabel(reload.status)}`,
      reload.endTime ? `**End time:** ${new Date(reload.endTime).toLocaleString()}` : '',
      reload.startTime && reload.endTime
        ? `**Duration:** ${formatDuration(reload.startTime, reload.endTime)}`
        : '',
      reload.type ? `**Type:** ${reload.type}` : '',
      reload.errorMessage ? `**Error:** ${reload.errorMessage}` : '',
      `**Reload ID:** ${reload.id}`,
      hasLog ? '**Log:** Available' : '**Log:** Not available',
    ].filter(Boolean).join('\n\n');

    this.tooltip = new vscode.MarkdownString(tooltipLines);

    this.command = {
      command: 'qlikcloud.showReloadSummary',
      title: 'Show Reload Summary',
      arguments: [this],
    };
  }
}

// ── Provider ───────────────────────────────────────────────────────────────

export class QlikReloadLogProvider implements vscode.TreeDataProvider<ReloadLogItem> {
  private readonly _onDidChangeTreeData = new vscode.EventEmitter<ReloadLogItem | undefined | null>();
  readonly onDidChangeTreeData = this._onDidChangeTreeData.event;

  private _items: ReloadLogItem[] = [];

  loadLogs(appId: string, client: QlikClient): void {
    this._items = [];
    this._onDidChangeTreeData.fire(null);

    Promise.all([
      client.getReloads(appId),
      client.getReloadLogs(appId),
    ]).then(([reloads, logMetas]) => {
      const logSet = new Set(logMetas.map(m => m.reloadId).filter(Boolean) as string[]);
      this._items = reloads.map(r => new ReloadLogItem(r, logSet.has(r.id)));
      this._onDidChangeTreeData.fire(null);
    }).catch(() => {});
  }

  clear(): void {
    this._items = [];
    this._onDidChangeTreeData.fire(null);
  }

  getTreeItem(element: ReloadLogItem): vscode.TreeItem {
    return element;
  }

  getChildren(element?: ReloadLogItem): ReloadLogItem[] {
    if (element) return [];
    return this._items;
  }
}
