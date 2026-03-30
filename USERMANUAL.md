# User Manual - Managed Bookmarks Creator

**Version:** 2.12.0.0
**Author:** Decoster Hans
**Script:** `ManagedBookmarksCreator.ps1`

---

## Table of Contents

1. [Overview](#overview)
2. [Credits and AI assistance](#credits-and-ai-assistance)
3. [License](#license)
4. [Public repository](#public-repository)
5. [Requirements](#requirements)
6. [Configuration file](#configuration-file)
7. [Starting the tool](#starting-the-tool)
8. [Recovery and autosave](#recovery-and-autosave)
9. [Toolbar reference](#toolbar-reference)
10. [Building a bookmark structure](#building-a-bookmark-structure)
11. [Importing existing JSON](#importing-existing-json)
12. [Reading from the registry](#reading-from-the-registry)
13. [Writing directly to the registry](#writing-directly-to-the-registry)
14. [Saving and loading JSON files](#saving-and-loading-json-files)
15. [Exporting a deployment script](#exporting-a-deployment-script)
16. [Deploying the generated script](#deploying-the-generated-script)
17. [Logging](#logging)
18. [Strong points](#strong-points)
19. [Weak points](#weak-points)

---

## Overview

Managed Bookmarks Creator is a PowerShell WinForms GUI that lets you build and maintain a
`ManagedBookmarks` JSON policy for **Google Chrome** and **Microsoft Edge**.

The policy controls the browser bookmarks bar visible to every managed user. Administrators
can create folders, sub-folders and links in the GUI, preview the resulting JSON, and either:

- save the JSON to a file for manual use in a GPO or Intune policy,
- export a ready-to-deploy PowerShell script that writes the policy keys directly to the
  Windows registry, or
- write directly to the registry from within the GUI when running with Administrator rights.

When launched with elevated Administrator privileges the title bar is prefixed with **`Admin: `**
and additional admin-only features become available: the **Registry** tab and the
**Write to Registry** toolbar button.

---

## Credits and AI assistance

- The images for this folder are created with Copilot.
- The changelog and this user manual were created with ChatGPT and Claude.ai.
- The GUI was adjusted with help from ChatGPT and Claude.ai.

---

## License

This application may be used, copied, modified, merged, published, distributed, sublicensed,
and sold under the terms of the **MIT License**.

In practical terms, the MIT License means:

- You are allowed to use the software privately and commercially.
- You are allowed to adapt the code and redistribute your own modified version.
- You must keep the original copyright notice and the MIT license text with the software.
- The software is provided **"as is"**, without warranty of any kind.
- The author cannot be held liable for damages, data loss, or other issues resulting from use of the software.

The MIT License is a short and permissive open-source license. It gives broad freedom to use
the application, but it also makes clear that the software comes without guarantees or support obligations.

---

## Public repository

The source code for this tool is publicly available on GitHub:

**[https://github.com/hansdeco/Managed_Bookmarks_manager](https://github.com/hansdeco/Managed_Bookmarks_manager)**

The repository contains:

- The main script `ManagedBookmarksCreator.ps1`
- This user manual
- The changelog
- Assets (toolbar icons, application images)

Bug reports, feature requests, and contributions are welcome via the repository's Issues and Pull Requests.

---

## Requirements

| Item | Minimum |
|---|---|
| PowerShell | 5.1 or higher |
| Windows | Windows 10 / Server 2016 or later |
| .NET | 4.7.2 (WinForms) |
| Output module *(optional)* | `Decoster.tech.output.psm1` in `..\..\Classes\` |
| Config file *(optional)* | `Decoster.tech.config.base.psd1` in `..\..\ConfigFiles\` |
| Administrator rights *(optional)* | Required for the Registry tab and Write to Registry button |

The script runs without the output module and config file. When they are absent, logging is
disabled and all file dialogs fall back to standard Windows behaviour.

---

## Configuration file

Located at `..\..\ConfigFiles\Decoster.tech.config.base.psd1` relative to the script.

```powershell
@{
    Logging = @{
        BasePath         = "C:\ProgramData\Decoster.tech\Scripting\Logs"
        LogRetentionDays = 400
    }
    Paths = @{
        ScriptBasePath = "C:\ProgramData\Decoster.tech\Scripting\PowerShell"
        ExeBasePath    = "C:\ProgramData\Decoster.tech\Scripting\Executables"
        JsonBasePath   = "C:\ProgramData\Decoster.tech\Scripting\Json"
    }
    Company = @{
        Name         = "Decoster.tech"
        Author       = "Decoster Hans"
        RegistryBase = "HKLM:\Software\Decoster.tech\Scripting"
    }
}
```

| Key | Effect when set |
|---|---|
| `Logging.BasePath` | Session log written to `<BasePath>\ManagedBookmarksCreator\` |
| `Logging.LogRetentionDays` | Log files older than this are deleted on startup |
| `Paths.JsonBasePath` | Save / Load dialogs open here; folder created automatically |
| `Paths.ScriptBasePath` | Generated `.ps1` files saved here; enables the Export Script button |
| `Company.Name` | Inserted in the generated script header and `$companyName` variable |
| `Company.Author` | Inserted in the generated script's origin-story block |
| `Company.RegistryBase` | Registry tracking path used by the generated script |

> **Note:** A missing or incomplete config file does not prevent the tool from running. Each
> missing key degrades only the feature that depends on it.

---

## Starting the tool

Run the script from PowerShell:

```powershell
.\ManagedBookmarksCreator.ps1
```

Or right-click and choose **Run with PowerShell** in Explorer.

To unlock admin-only features such as the Registry tab and the Write to Registry button,
start an elevated PowerShell session first:

```powershell
# Right-click PowerShell and choose "Run as Administrator", then:
.\ManagedBookmarksCreator.ps1
```

On startup the script:
1. Detects whether it is running with Administrator rights.
2. Loads the config file silently if present.
3. Loads the output module silently if present.
4. Validates all required config keys and shows a warning in the status bar if any are missing.
5. Cleans up log files older than `LogRetentionDays`.
6. Checks for an unclean previous session and offers to restore the last autosave snapshot.
7. Opens the GUI, with `Admin: ` prefixed to the title bar when elevated.

---

## Recovery and autosave

The editor now includes a built-in session recovery flow to reduce data loss after an unexpected close.

### How it works

- Recovery data is stored per user in `%LOCALAPPDATA%\ManagedBookmarksCreator\Recovery`.
- On startup the tool creates a `session.lock` marker and removes it on clean shutdown.
- The current editor state is marked as **dirty** whenever content changes.
- A debounce timer (`15` seconds) writes `latest.state.json` only when there are unsaved changes.
- On clean close, a final snapshot is written and the lock file is removed.

### Recovery prompt

When the previous session ended unexpectedly (for example a crash or forced close), the lock marker is still present.
At next startup, the tool prompts whether to restore the latest snapshot.

- **Yes**: restore top-level label and full bookmark tree from the snapshot.
- **No**: keep current empty/new state and continue normally.

If recovery data is missing or invalid, startup continues safely without blocking the GUI.

---

## Toolbar reference

| Button | Description | Admin only |
|---|---|---|
| `Folder` | Add a new folder at the root level | No |
| `Subfolder` | Add a subfolder inside the selected folder | No |
| `Link` | Add a bookmark link inside the selected folder | No |
| `Root link` | Add a bookmark link at the root level | No |
| `Edit` | Edit the selected item's name / URL | No |
| `Delete` | Delete the selected item with confirmation | No |
| `Up` | Move the selected item one position up; crosses folder boundaries | No |
| `Down` | Move the selected item one position down; crosses folder boundaries | No |
| `Undo` | Revert the previous editor change | No |
| `Redo` | Re-apply the most recently undone change | No |
| `Load` | Load an existing `managed_bookmarks.json` from disk | No |
| `Save` | Save the current JSON to `JsonBasePath` or a file dialog | No |
| `Copy` | Copy the compact JSON to the clipboard | No |
| `Validate` | Run JSON roundtrip contract test (build -> import -> build) | No |
| `Export Script` | Generate a registry deployment `.ps1` when config is complete | No |
| `Write to Registry` | Write bookmarks directly to HKLM for Chrome and Edge | Yes |

The **Export Script** button is hidden when the config file is absent or `Paths.ScriptBasePath`
is not set.

The **Write to Registry** button and the **Registry** tab are hidden when the script is not
running with Administrator rights.

The **Undo** and **Redo** buttons are enabled only when the corresponding history stack contains items.

---

## Building a bookmark structure

### Adding a folder

1. Click **Folder**.
2. Enter a name.
3. Click **OK**.

The folder appears at the root level of the tree.

### Adding a subfolder

1. Select an existing folder in the tree.
2. Click **Subfolder**.
3. Enter a name.
4. Click **OK**.

### Adding a link

1. Select the folder in which the link should appear.
2. Click **Link**.
3. Enter a **Name** and **URL**.
4. Click **OK**.

### Adding a root-level link

Click **Root link** to add a link outside any folder, directly in the bookmarks bar.

### Editing an item

Select the item and click **Edit** or double-click it. Modify the name and/or URL and click **OK**.

### Deleting an item

Select the item and click **Delete**. Confirm the prompt. Deleting a folder removes all its
children recursively.

### Reordering items

**Using the toolbar buttons:**

Select an item and use **▲ Up** / **▼ Down** to move it one position at a time through the entire tree.

- Within a folder, the item moves up or down among its siblings normally.
- When the next sibling is a folder, pressing **Down** moves the item **into** that folder as its first child (and vice versa for **Up** and the last child).
- When the item reaches the top of a folder, pressing **Up** moves it out to just before the parent folder. When it reaches the bottom, pressing **Down** moves it out to just after the parent folder.
- This way every visual position in the tree is reachable one step at a time.

**Using drag & drop:**

Click and hold an item, drag it to the desired location, and release the mouse button.

- The target node is highlighted in blue while you drag over it.
- The **top half** of a node inserts the dragged item **before** that node; the **bottom half** inserts it **after**.
- Dropping onto a descendant of the dragged item is not allowed.
- The tree scrolls automatically when you drag near the top or bottom edge.
- The move is undoable with **↶ Undo**.

---

## Importing existing JSON

Switch to the **JSON Import** tab and paste an existing `ManagedBookmarks` JSON string. Two
formats are accepted:

**Direct policy array** (paste straight from a GPO or Intune policy field):
```json
[{"name":"Google","url":"google.com"},{"name":"IT","children":[...]}]
```

**Full policy object** (exported from an admin tool or ADMX template):
```json
{"BookmarkBarEnabled":true,"ManagedBookmarks":[{"name":"Google","url":"google.com"},...]}
```

Click **Import JSON** to parse the string and rebuild the tree. Wrapper objects using either
`ManagedBookmarks` or `ManagedFavorites` are accepted. The **JSON Preview** tab is selected
automatically after a successful import.

Click **Clear** to empty the input field.

---

## Reading from the registry

> **Requires Administrator rights.** This tab is hidden when the script runs without elevation.

The **Registry** tab lets you inspect the currently active policy value stored in the Windows
registry, without needing the original JSON file.

### Steps

1. Open the **Registry** tab.
2. Select the browser source from the dropdown:
   - `Chrome (HKLM\...\Google\Chrome)` reads `ManagedBookmarks`
   - `Edge   (HKLM\...\Microsoft\Edge)` reads `ManagedFavorites`
3. Click **Read**.
4. Optionally click **Load in editor** to import the displayed JSON directly into the tree.

If the registry key or value does not exist a warning is shown in the status bar and the text
area is cleared.

---

## Writing directly to the registry

> **Requires Administrator rights.** This button is hidden when the script runs without elevation.

The **Write to Registry** toolbar button performs the same registry writes as the exported
deployment script, but immediately and interactively from within the GUI.

### Steps

1. Build or import the desired bookmark structure.
2. Click **Write to Registry**.
3. Review the confirmation dialog and click **Yes**.

### What gets written

| Registry path | Value name | Type | Value |
|---|---|---|---|
| `HKLM:\SOFTWARE\Policies\Google\Chrome` | `ManagedBookmarks` | String | Current JSON |
| `HKLM:\SOFTWARE\Policies\Google\Chrome` | `BookmarkBarEnabled` | DWORD | `1` |
| `HKLM:\SOFTWARE\Policies\Microsoft\Edge` | `ManagedFavorites` | String | Current JSON |
| `HKLM:\SOFTWARE\Policies\Microsoft\Edge` | `FavoritesBarEnabled` | DWORD | `1` |

Registry keys that do not exist are created automatically.

### After writing

A dialog confirms success and reminds you to fully close and reopen Chrome and Edge,
including the system tray icon, for the new policy to take effect.

If one browser fails, for example due to access denied, an error dialog lists the affected
browser with the exception message. The other browser is still written if its write succeeded.

---

## Saving and loading JSON files

### Save

Click **Save**.

- When `JsonBasePath` is configured, a small dialog asks for a file name only.
  The `.json` extension is added automatically. The folder is created if it does not exist.
  An overwrite confirmation is shown when the file already exists.
- When `JsonBasePath` is not configured, the standard Windows Save File dialog is shown.

### Load

Click **Load**. Select a `.json` file. The tool accepts the same two formats described under
[Importing existing JSON](#importing-existing-json).

When `JsonBasePath` is configured and exists, the file dialog opens in that folder by default.

### Copy

Click **Copy** to copy the compact single-line JSON string to the clipboard, ready to paste
into a GPO, Intune configuration profile, or ADMX template field.

---

## Exporting a deployment script

> Requires a config file with `Paths.ScriptBasePath`, `Company.Name`, `Company.Author`,
> and `Company.RegistryBase` set.

1. Build the bookmark structure.
2. Click **Export Script**.
3. Enter a **script name**.
4. Optionally edit the **description** that appears in the script's `.DESCRIPTION` block.
5. Click **Export**.

The generated script is saved to `Paths.ScriptBasePath`. If the folder does not exist it is
created automatically.

### What the generated script does

When deployed with local Administrator rights it:

1. Creates a transcript log in the company scripting folder.
2. Checks that it runs as Administrator.
3. Detects the OS version.
4. Writes registry tracking keys under `Company.RegistryBase\<ScriptName>`.
5. Sets `ManagedBookmarks` and `BookmarkBarEnabled` under `HKLM:\SOFTWARE\Policies\Google\Chrome`.
6. Sets `ManagedFavorites` and `FavoritesBarEnabled` under `HKLM:\SOFTWARE\Policies\Microsoft\Edge`.
7. Writes an on-screen reminder that Chrome and Edge must be fully closed and reopened.
8. Writes the end time to the tracking registry key.
9. Stops the transcript.

The script is self-contained: the JSON is embedded as a literal string so it requires no
external files at deploy time.

---

## Deploying the generated script

Common deployment options:

| Method | Notes |
|---|---|
| SCCM / ConfigMgr | Deploy as a script package; run as SYSTEM |
| Intune | Use as part of a PowerShell deployment or remediation package |
| GPO Startup script | Machine-level GPO; runs as SYSTEM at boot |
| Manual | Run as Administrator in a PowerShell console |

---

## Logging

When the output module and config are present, a formatted session log is written to:

```text
<Logging.BasePath>\ManagedBookmarksCreator\<username>_<yyyyMMdd_HHmmss>_ManagedBookmarksCreator.txt
```

In addition, the application now keeps a plain-text fallback log in:

```text
%LOCALAPPDATA%\Decoster.tech\ManagedBookmarksCreator\Logs\<username>_<yyyyMMdd_HHmmss>_ManagedBookmarksCreator.log
```

This fallback log is intended for diagnostics and remains available even when the optional output module is not installed.

The startup log includes:

| Key | Value |
|---|---|
| Started | Timestamp |
| User | Windows username |
| Admin | Elevated or not |
| Author | Script author |
| Module | Loaded or not found |
| Config | JsonBasePath value or not found |
| ScriptPath | ScriptBasePath value or not configured |

All actions performed during the session are logged: folder and link additions, edits,
deletions, file loads, saves, imports, registry reads, registry writes, and exports.

The logging layer now also captures additional runtime failures:

- explicitly handled action errors from `try/catch` blocks,
- new PowerShell runtime errors detected in the global `$Error` buffer,
- unhandled WinForms UI exceptions, and
- unhandled AppDomain exceptions.

This means the log now covers a much larger part of the errors that would normally be visible in a PowerShell session.
Log files older than `LogRetentionDays` are deleted automatically on the next startup.

---

## Strong points

- **No dependencies for basic use** - runs standalone without config or output module.
- **Admin-mode detection** - elevated privileges are detected at startup and unlock admin-only features automatically.
- **Session recovery and autosave** - unexpected shutdowns can be restored from the last snapshot.
- **Direct registry read** - the Registry tab lets admins inspect the currently active policy directly from `HKLM`.
- **Direct registry write** - the Write to Registry button applies the policy immediately on the target machine.
- **Dual JSON format support** - imports both direct arrays and wrapper objects such as `ManagedBookmarks` and `ManagedFavorites`.
- **Undo/redo history** - most editor operations can be reverted without reloading from disk.
- **Contract validation** - built-in roundtrip test helps detect import/export drift early.
- **URL validation** - link dialogs validate format and limit entries to `http`/`https` URLs.
- **PS 5.1 compatible** - explicit wrapping guards against single-element array unwrapping.
- **Folder auto-creation** - `JsonBasePath` and `ScriptBasePath` are created on the fly.
- **Config validation on startup** - missing or empty config keys are reported immediately.
- **Ready-to-deploy script generation** - the exported `.ps1` follows the company scripting template.
- **Fully embedded JSON** - the generated deployment script contains the JSON as a literal string.
- **Session log with retention** - older logs are cleaned up automatically.
- **Fallback diagnostics log** - a plain-text LocalAppData log is kept even without the optional output module.
- **Drag-and-drop reordering** - items can be dragged to any position in the tree with immediate visual feedback.
- **Full flat-list Up/Down navigation** - arrow buttons traverse every visual position including across folder boundaries.
- **MIT licensed usage model** - broad reuse and adaptation are allowed with minimal obligations.

---

## Weak points

- **No multi-select** - only one item can be selected, moved, or deleted at a time.
- **Recovery is single-snapshot** - only the latest state is retained; there is no multi-version recovery history.
- **Flat file picker for Load** - the standard Windows file picker is still used.
- **Script version in generated output is static** - there is no automatic increment mechanism.
- **Admin rights required for registry features** - reading from and writing to `HKLM:\SOFTWARE\Policies` requires elevation.
- **No uninstall / removal script** - no companion script is generated to remove the registry keys.

---

## Signature

```text
╔════════════════════════════════════════════╗
║  ██ ██  WRITTEN DESIGNED BY DECOSTER.TECH  ║
║  █████                                      ║
╚══██ ██══════════════════════════════════════╝
    ██ ██ans
Email: scripting@decoster.tech
```
