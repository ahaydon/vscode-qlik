import * as vscode from 'vscode';

/**
 * Read-only TextDocumentContentProvider for reload log content.
 *
 * URI scheme:  qlikreloadlog://<reloadId>.log
 */
export class QlikReloadLogContentProvider implements vscode.TextDocumentContentProvider {
  static readonly SCHEME = 'qlikreloadlog';

  private readonly _cache = new Map<string, string>(); // reloadId → log text
  private readonly _emitter = new vscode.EventEmitter<vscode.Uri>();
  readonly onDidChange = this._emitter.event;

  static uri(reloadId: string): vscode.Uri {
    return vscode.Uri.from({
      scheme: QlikReloadLogContentProvider.SCHEME,
      path: `/${reloadId}.log`,
    });
  }

  get(reloadId: string): string | undefined {
    return this._cache.get(reloadId);
  }

  store(reloadId: string, content: string): void {
    this._cache.set(reloadId, content);
    this._emitter.fire(QlikReloadLogContentProvider.uri(reloadId));
  }

  provideTextDocumentContent(uri: vscode.Uri): string {
    const reloadId = uri.path.replace(/^\//, '').replace(/\.log$/, '');
    return this._cache.get(reloadId) ?? '';
  }
}
