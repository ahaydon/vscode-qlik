import * as vscode from 'vscode';

/**
 * Read-only TextDocumentContentProvider for historical script section content.
 *
 * URI scheme:  qlikhist://<scriptId>/<encodedSectionName>.qvs
 *
 * The provider is populated by QlikHistoryProvider before opening a diff editor.
 */
export class QlikHistoryContentProvider implements vscode.TextDocumentContentProvider {
  static readonly SCHEME = 'qlikhist';

  private readonly _cache = new Map<string, string>(); // uri.toString() → body
  private readonly _emitter = new vscode.EventEmitter<vscode.Uri>();
  readonly onDidChange = this._emitter.event;

  static uri(scriptId: string, sectionName: string): vscode.Uri {
    return vscode.Uri.from({
      scheme: QlikHistoryContentProvider.SCHEME,
      authority: scriptId,
      path: `/${encodeURIComponent(sectionName)}.qvs`,
    });
  }

  /** Called by QlikHistoryProvider when section content is available */
  store(scriptId: string, sectionName: string, body: string): void {
    const uri = QlikHistoryContentProvider.uri(scriptId, sectionName);
    this._cache.set(uri.toString(), body);
    this._emitter.fire(uri);
  }

  provideTextDocumentContent(uri: vscode.Uri): string {
    return this._cache.get(uri.toString()) ?? '';
  }
}
