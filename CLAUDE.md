# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Development Commands

```bash
# Install dependencies
npm install

# Compile (one-shot, outputs to dist/extension.js)
npm run compile

# Watch mode for development
npm run watch

# Package as .vsix
npm run package
```

The build uses esbuild (not tsc directly) — see `esbuild.js`. TypeScript is type-checked separately via `tsconfig.json`.

To run the extension, open the project in VS Code and press **F5** to launch the Extension Development Host.

There are no automated tests in this project.

## Architecture

This is a VS Code extension that lets users edit Qlik Cloud app load scripts directly in VS Code. The main flow is: select a Qlik context → open an app → edit sections → save back to Qlik Cloud.

### Key architectural concepts

**Script model** (`src/scriptModel.ts`): Qlik load scripts are delimited by `///$tab <name>` comment lines. `parseScript` splits the raw string into ordered `ScriptSection[]` (each with a stable UUID `id`, `name`, and `body`). `serializeScript` rejoins them. This is the source of truth for the section data model.

**Virtual filesystem** (`src/scriptFS.ts`): Sections are exposed as in-memory files under the `qlikscript://` URI scheme. URI format: `qlikscript://<sectionId>/<appId>/<name>.qvs`. The authority holds the `sectionId`; the first path segment holds the `appId`. Section content lives here; section ordering lives in the tree provider.

**Tree provider** (`src/treeProvider.ts`): Drives the "Script Sections" sidebar view. Owns section ordering and dirty state. When sections are reordered or renamed, the tree is the source of truth; the FS is updated to match.

**History** (`src/historyProvider.ts`, `src/historyContentProvider.ts`): Loads script version history via the Qlik API and presents it in the "Script History" view. `historyContentProvider` registers the `qlikhist://` scheme used as the left side of diff editors. Each version is diffed against its direct predecessor (not the working copy).

**Language support** (`src/languageProvider.ts`): Provides completions (keywords, functions, variables, subs, tables parsed from all open sections), hover docs, and diagnostics for `.qvs` files. Completions are cross-section — symbols from all sections are available everywhere.

**Qlik API client** (`src/qlikClient.ts`): Thin wrapper around `@qlik/api` (ESM-only package, bundled by esbuild). Handles spaces, app search, script CRUD, and app reload via QIX engine WebSocket (`openAppSession`). Reload streams progress via `getProgress(0)` polling.

**Context loading** (`src/contexts.ts`): Reads `~/.qlik/contexts.yml`. Supports API key (Bearer token in headers) and OAuth2 client credentials auth. Contexts without usable auth are silently skipped.

**Reload logs** (`src/reloadLogProvider.ts`, `src/reloadLogContentProvider.ts`): The "Reloads" panel lists recent reloads (status, timestamp, duration) via `QlikReloadLogProvider`. `reloadLogContentProvider` registers the `qlikreloadlog://` scheme; reload log text is fetched on demand and cached.

**Extension entry** (`src/extension.ts`): Module-level state (`currentContext`, `currentClient`, `currentAppId`) is the only global state. All commands are registered in `activate()`. The `__qlikRefreshStatus` globalThis hack lets command implementations trigger status bar updates without circular imports.

**Tab titles**: Script sections and reload logs are opened via `vscode.commands.executeCommand('vscode.open', uri, options, label)` rather than `showTextDocument`, so a custom label (section name or reload timestamp) is set explicitly. History diffs use `vscode.diff` with an explicit title for the same reason. Without a custom label, VS Code shows the full URI path with leading backslashes on Windows.

### URI schemes
- `qlikscript://` — writable virtual FS for editing sections
- `qlikhist://` — read-only content provider for history diff views (left side)
- `qlikreloadlog://` — read-only content provider for reload log text
