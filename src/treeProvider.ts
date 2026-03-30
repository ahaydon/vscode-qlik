import * as vscode from 'vscode';
import { ScriptSection } from './scriptModel';
import { QlikScriptFS } from './scriptFS';

export class SectionItem extends vscode.TreeItem {
  constructor(
    public readonly section: ScriptSection,
    public readonly index: number,
    public readonly total: number,
    public readonly appId: string,
    public isDirty: boolean,
  ) {
    const label = `${index + 1}. ${section.name}${isDirty ? ' ●' : ''}`;
    super(label, vscode.TreeItemCollapsibleState.None);

    this.contextValue = 'section';
    this.tooltip = section.name;
    this.description = isDirty ? 'unsaved' : undefined;

    this.command = {
      command: 'qlikcloud.openSection',
      title: 'Open Section',
      arguments: [this],
    };

    this.iconPath = new vscode.ThemeIcon(
      isDirty ? 'circle-filled' : 'symbol-file',
      isDirty ? new vscode.ThemeColor('gitDecoration.modifiedResourceForeground') : undefined,
    );
  }
}

export class QlikScriptTreeProvider implements vscode.TreeDataProvider<SectionItem> {
  private _onDidChangeTreeData = new vscode.EventEmitter<SectionItem | undefined | null>();
  readonly onDidChangeTreeData = this._onDidChangeTreeData.event;

  private _sections: ScriptSection[] = [];
  private _appId: string | undefined;
  private _appName: string | undefined;
  private _dirtyIds = new Set<string>();

  setApp(appId: string, appName: string, sections: ScriptSection[]): void {
    this._appId = appId;
    this._appName = appName;
    this._sections = sections;
    this._dirtyIds.clear();
    this._onDidChangeTreeData.fire(null);
  }

  clear(): void {
    this._appId = undefined;
    this._appName = undefined;
    this._sections = [];
    this._dirtyIds.clear();
    this._onDidChangeTreeData.fire(null);
  }

  markDirty(sectionId: string): void {
    this._dirtyIds.add(sectionId);
    this._onDidChangeTreeData.fire(null);
  }

  markClean(): void {
    this._dirtyIds.clear();
    this._onDidChangeTreeData.fire(null);
  }

  getSections(): ScriptSection[] {
    return this._sections;
  }

  getAppId(): string | undefined {
    return this._appId;
  }

  getAppName(): string | undefined {
    return this._appName;
  }

  isDirty(): boolean {
    return this._dirtyIds.size > 0;
  }

  addSection(name: string): ScriptSection {
    const { randomUUID } = require('crypto') as typeof import('crypto');
    const section: ScriptSection = { id: randomUUID(), name, body: '' };
    this._sections.push(section);
    this._onDidChangeTreeData.fire(null);
    return section;
  }

  deleteSection(sectionId: string): void {
    this._sections = this._sections.filter(s => s.id !== sectionId);
    this._dirtyIds.delete(sectionId);
    this._onDidChangeTreeData.fire(null);
  }

  moveUp(sectionId: string): void {
    const idx = this._sections.findIndex(s => s.id === sectionId);
    if (idx > 0) {
      [this._sections[idx - 1], this._sections[idx]] = [this._sections[idx], this._sections[idx - 1]];
      this._onDidChangeTreeData.fire(null);
    }
  }

  moveDown(sectionId: string): void {
    const idx = this._sections.findIndex(s => s.id === sectionId);
    if (idx >= 0 && idx < this._sections.length - 1) {
      [this._sections[idx], this._sections[idx + 1]] = [this._sections[idx + 1], this._sections[idx]];
      this._onDidChangeTreeData.fire(null);
    }
  }

  renameSection(sectionId: string, newName: string): void {
    const s = this._sections.find(s => s.id === sectionId);
    if (s) {
      s.name = newName;
      this._onDidChangeTreeData.fire(null);
    }
  }

  // ── TreeDataProvider ───────────────────────────────────────────────────────

  getTreeItem(element: SectionItem): vscode.TreeItem {
    return element;
  }

  getChildren(element?: SectionItem): SectionItem[] {
    if (element) return [];
    if (!this._appId) return [];

    return this._sections.map(
      (s, i) => new SectionItem(s, i, this._sections.length, this._appId!, this._dirtyIds.has(s.id)),
    );
  }
}
