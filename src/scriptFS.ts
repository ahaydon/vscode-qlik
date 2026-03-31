import * as vscode from 'vscode';
import { ScriptSection } from './scriptModel';

/**
 * In-memory virtual filesystem for Qlik script sections.
 *
 * URI scheme:  qlikscript://<sectionId>/<appId>/<encodedName>.qvs
 *
 * The authority holds the internal sectionId (not visible in tab titles).
 * The first path segment is the appId, visible in the tab title/breadcrumb.
 *
 * Sections are read/written as UTF-8 text.  The ordering lives in the
 * QlikScriptTreeProvider (owned by the extension); only the content lives here.
 */
export class QlikScriptFS implements vscode.FileSystemProvider {
  static readonly SCHEME = 'qlikscript';

  private readonly _files = new Map<string, Uint8Array>();
  private readonly _emitter = new vscode.EventEmitter<vscode.FileChangeEvent[]>();
  readonly onDidChangeFile = this._emitter.event;

  /** Called when sections are loaded from Qlik Cloud */
  populateSections(appId: string, sections: ScriptSection[]): void {
    // Remove stale files for this app
    const prefix = `/${appId}/`;
    for (const key of this._files.keys()) {
      if (key.startsWith(prefix)) this._files.delete(key);
    }

    for (const s of sections) {
      const key = this._key(appId, s.id);
      this._files.set(key, Buffer.from(s.body, 'utf8'));
    }
  }

  /** Returns updated body for a section, or undefined if not modified */
  readSectionBody(appId: string, sectionId: string): string | undefined {
    const data = this._files.get(this._key(appId, sectionId));
    return data ? Buffer.from(data).toString('utf8') : undefined;
  }

  /** Force-write section body (e.g., after rename) without triggering a FS event */
  writeSectionBody(appId: string, sectionId: string, body: string): void {
    this._files.set(this._key(appId, sectionId), Buffer.from(body, 'utf8'));
  }

  removeSections(appId: string, sectionIds: string[]): void {
    for (const id of sectionIds) this._files.delete(this._key(appId, id));
  }

  /** Build a URI for a section */
  static uri(appId: string, section: ScriptSection): vscode.Uri {
    const safeName = encodeURIComponent(section.name);
    return vscode.Uri.from({
      scheme: QlikScriptFS.SCHEME,
      authority: section.id,
      path: `/${appId}/${safeName}.qvs`,
    });
  }

  // ── FileSystemProvider ─────────────────────────────────────────────────────

  watch(): vscode.Disposable {
    return new vscode.Disposable(() => {});
  }

  stat(uri: vscode.Uri): vscode.FileStat {
    const key = this._uriToKey(uri);
    if (key && this._files.has(key)) {
      return { type: vscode.FileType.File, ctime: 0, mtime: 0, size: this._files.get(key)!.length };
    }
    // Directory stat for the app root
    return { type: vscode.FileType.Directory, ctime: 0, mtime: 0, size: 0 };
  }

  readDirectory(uri: vscode.Uri): [string, vscode.FileType][] {
    const appId = uri.path.split('/').filter(Boolean)[0] ?? uri.authority;
    const prefix = `/${appId}/`;
    const result: [string, vscode.FileType][] = [];
    for (const key of this._files.keys()) {
      if (key.startsWith(prefix)) {
        result.push([key.slice(prefix.length), vscode.FileType.File]);
      }
    }
    return result;
  }

  createDirectory(): void { /* not used */ }

  readFile(uri: vscode.Uri): Uint8Array {
    const key = this._uriToKey(uri);
    if (key) {
      const data = this._files.get(key);
      if (data) return data;
    }
    throw vscode.FileSystemError.FileNotFound(uri);
  }

  writeFile(uri: vscode.Uri, content: Uint8Array, _options: { create: boolean; overwrite: boolean }): void {
    const key = this._uriToKey(uri);
    if (!key) throw vscode.FileSystemError.FileNotFound(uri);
    this._files.set(key, content);
    this._emitter.fire([{ type: vscode.FileChangeType.Changed, uri }]);
    // Notify the extension that a section was modified
    this._onSectionWritten.fire(uri);
  }

  delete(uri: vscode.Uri): void {
    const key = this._uriToKey(uri);
    if (key) this._files.delete(key);
  }

  rename(oldUri: vscode.Uri, newUri: vscode.Uri): void {
    const oldKey = this._uriToKey(oldUri);
    const newKey = this._uriToKey(newUri);
    if (oldKey && newKey) {
      const data = this._files.get(oldKey);
      if (data) {
        this._files.set(newKey, data);
        this._files.delete(oldKey);
      }
    }
  }

  // ── Internal ───────────────────────────────────────────────────────────────

  readonly _onSectionWritten = new vscode.EventEmitter<vscode.Uri>();
  readonly onSectionWritten = this._onSectionWritten.event;

  private _key(appId: string, sectionId: string): string {
    return `/${appId}/${sectionId}`;
  }

  private _uriToKey(uri: vscode.Uri): string | undefined {
    // URI: qlikscript://<sectionId>/<appId>/<name>.qvs
    // authority = sectionId, parts[0] = appId
    const parts = uri.path.split('/').filter(Boolean);
    if (parts.length < 1) return undefined;
    const sectionId = uri.authority;
    const appId = parts[0];
    return `/${appId}/${sectionId}`;
  }
}

