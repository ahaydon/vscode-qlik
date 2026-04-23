# Changelog

## [0.1.0] - 2026-04-23

### Added

- **Language intelligence** — completions (keywords, functions, variables, subs, tables), hover documentation, and diagnostics for `.qvs` files; completions are cross-section, drawing symbols from all open sections
- **Reload history panel** — the sidebar now shows all recent reloads with status icons, timestamps, and duration; reloads that have a log available show a download icon to open the full log in the editor

### Fixed

- Tab titles for script sections and reload logs no longer begin with a leading backslash on Windows; sections now show just the section name and reload logs show the reload timestamp
- History diff no longer compares a version against the current working copy; it now diffs each version against its direct predecessor
- URL-encoded characters in section names are decoded correctly in tab titles

## [0.0.1] - 2026-03-31

### Added

- **Script Editor** — open and edit Qlik Cloud app load script sections as virtual documents in VS Code with full syntax highlighting (`.qvs`)
- **Script History panel** — browse previous versions of the load script in a tree view, expand a version to see which sections were added, modified, or removed, and click a section to open a diff editor showing the changes made in that version
- **Revert to version** — hover over any history entry and click the discard icon to restore the script to that version; a confirmation prompt is shown before applying
- **App reload** — trigger a Qlik Cloud app reload from VS Code; live log output streams to a dedicated Output Channel as the reload progresses, with a cancellable progress notification
- **Qlik script syntax highlighting** — grammar-based highlighting for `.qvs` files
