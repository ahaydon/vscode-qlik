# Qlik Cloud Script Editor

A VS Code extension for editing and managing Qlik Cloud app load scripts directly from your editor.

## Features

### Script Editing
- Browse and open Qlik Cloud apps via the **Qlik Cloud** activity bar panel
- View and edit individual script **sections** (tabs) as separate files with full syntax highlighting
- Reorder, rename, add, and delete sections from the tree view
- Save changes back to Qlik Cloud with an optional version message

### Script History
- Browse the full version history of an app's load script
- Expand any version to see which sections changed compared to your current working copy
- Click a changed section to open a **diff editor** showing exactly what differs
- Revert to any previous version with a single click

### App Reload & Logs
- Trigger a Qlik Cloud app reload from VS Code using the engine API
- Live log output streams to a dedicated **Qlik Cloud Reload** output channel
- Cancel an in-progress reload via the notification
- Browse recent reloads in the **Reloads** panel with status icons, timestamps, and duration
- Click any reload with an available log to open the full log text in the editor

### Language Intelligence
- Completions for keywords, built-in functions, variables, subs, and table names — drawn from all open sections
- Hover documentation for built-in functions
- Diagnostics flag unknown variables and common mistakes

### Syntax Highlighting
- Full syntax highlighting for Qlik Script (`.qvs`) files
- Covers keywords, functions, comments, strings, dollar-sign expansions, and more

## Requirements

- A Qlik Cloud tenant
- [qlik-cli](https://qlik.dev/toolkits/qlik-cli/) installed and at least one context configured in `~/.qlik/contexts.yml`

## Getting Started

1. Install the extension
2. Open the **Qlik Cloud** panel in the activity bar
3. Click **Select Qlik Context** (person icon) and choose your tenant
4. Click **Open Qlik App** (search icon) to browse and open an app
5. Script sections appear in the **Script Sections** tree — click any section to open it in the editor

## Commands

| Icon | Command | Description |
|------|---------|-------------|
| $(account) | Select Qlik Context | Choose which tenant to connect to |
| $(search) | Open Qlik App | Browse and open an app from your tenant |
| $(cloud-upload) | Save Script | Save all sections back to Qlik Cloud |
| $(add) | Add Section | Add a new script section |
| $(run) | Reload App | Execute the load script with live log output |
| $(refresh) | Refresh History | Reload the Script History panel |

## Authentication

The extension reads authentication from `~/.qlik/contexts.yml` (created by qlik-cli). Both **API key** and **OAuth2 client credentials** auth types are supported.

## License

[MIT](LICENSE)
