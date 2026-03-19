# Changelog

## 2.9.2.0

### EXE startup and path handling

- Added packaged-app detection so startup and shutdown `Write-Host` messages are suppressed when the tool runs as a converted `.exe`.
- Added exe-friendly asset resolution for the config file and output module by searching the script folder, startup folder and parent folders before falling back.
- Added save/export fallbacks so JSON and generated scripts can still be written through a normal save dialog when config-based base paths are unavailable.

## 2.9.1.0

### Documentation and licensing

- Added a dedicated MIT license section to `USERMANUAL.md`, including a short practical explanation of what the license allows and requires.
- Added a standalone `LICENSE` file with the full MIT License text in the `ManagedBookmarksCreator` folder.
- Added a one-line MIT license reference in the first comment block of `ManagedBookmarksCreator.ps1`.
- Cleaned up `USERMANUAL.md` by removing merge-conflict remnants and refreshing the document structure.

## 2.9.0.0

### Admin mode — title bar, Registry tab, Write to Registry

#### Admin detection moved to startup
- `$isAdmin` and `$adminPrefix` are now evaluated immediately after the `Add-LogEntry` helper is
  defined, before the startup log is written — so admin status can be logged at launch.
- Startup log now includes `Admin: yes — elevated` or `Admin: no` as a key-value entry.

#### Title bar prefix
- When the script runs elevated the form title is prefixed with `"Admin: "` so it is immediately
  clear from the taskbar and window title that elevated privileges are active.

#### New: Registry tab (admin only)
- A third tab **"Registry"** is added to the tab control, but only when `$isAdmin` is `$true`.
  Non-admin users never see the tab.
- Contains a dropdown to select the registry source:
  - `Chrome  (HKLM\...\Google\Chrome)` → reads `ManagedBookmarks`
  - `Edge    (HKLM\...\Microsoft\Edge)` → reads `ManagedFavorites`
- **Read** button reads the selected registry value and displays the raw JSON in the tab's text box.
  Logs the path and value name as a `step` entry; logs success as `ok` or failure as `fail`.
- **Load in editor** button imports the displayed JSON into the tree (same logic as the Import tab)
  and switches to the JSON Preview tab. Logs import result including root item count.

#### New: 🖊️ Write to Registry toolbar button (admin only)
- Visible only when `$isAdmin` is `$true`; rendered in orange (`$clrWarn`) as a visual hint that
  this is a destructive/write operation.
- Shows a Yes/No confirmation dialog before writing.
- Writes the current bookmarks JSON directly to HKLM for both browsers:
  - **Chrome**: `ManagedBookmarks` (String) + `BookmarkBarEnabled` (DWord = 1)
    at `HKLM:\SOFTWARE\Policies\Google\Chrome`
  - **Edge**: `ManagedFavorites` (String) + `FavoritesBarEnabled` (DWord = 1)
    at `HKLM:\SOFTWARE\Policies\Microsoft\Edge`
  - Registry key is created automatically if it does not exist.
- Each write is logged individually (`ok` on success, `fail` with error message on failure).
- On full success shows an information dialog reminding the user to fully restart Chrome and Edge.
- On partial or full failure shows an error dialog listing the failed browser(s) with the exception message.

## 2.8.0.0

### Review fixes

- Fixed generated deployment scripts so embedded JSON is now escaped safely when bookmark names, labels or URLs contain apostrophes.
- Import now also accepts wrapper objects that use `ManagedFavorites` as the root property, aligning the GUI with the generated Edge deployment script.
- Replaced direct `PlaceholderText` usage with a compatibility helper so the GUI also starts correctly in Windows PowerShell 5.1.
- Synced the user manual with the current script version and removed the outdated `gpupdate /force` deployment step.

## 2.7.0.0

### Generated script — gpupdate verwijderd

- `gpupdate /force` verwijderd uit het gegenereerde script. Chrome en Edge lezen de
  policy-sleutels rechtstreeks uit de registry bij opstart; `gpupdate` ververst
  GPO-policies vanuit Active Directory en heeft geen effect op directe registry-schrijfacties.
- De ACTION REQUIRED-melding blijft behouden zodat de gebruiker weet dat beide browsers
  volledig gesloten en heropend moeten worden.

## 2.6.0.0

### Bar label field moved into the tree pane

- Removed the bar label `ToolStripLabel` + `ToolStripTextBox` from the toolbar.
- Added a `Panel` (height 36 px, `Dock=Top`) above the `TreeView` in the upper split pane.
  It contains the static label "Bookmarks bar label:" on the left and the text input filling
  the rest — making the field's purpose immediately clear in context of the tree structure.
- The JSON preview still updates live as you type.

## 2.5.0.0

### toplevel_name support — Edge fix + user-settable bar label

#### Root cause of Edge not showing managed favorites
- Edge's `ManagedFavorites` policy requires a `{"toplevel_name":"..."}` entry as the first
  element of the JSON array. Without it Edge silently ignores the policy.
  Chrome is more lenient and shows bookmarks even without this entry (falling back to the
  label "Managed Bookmarks"), which is why Chrome worked but Edge did not.

#### New: Bar label field in the toolbar
- A `Bar label:` text box has been added to the right end of the toolbar.
- The value entered here is written as `{"toplevel_name":"<value>"}` at the start of the
  JSON array, which sets the folder name visible in both Chrome and Edge.
- The field is optional: when left empty the `toplevel_name` entry is omitted (Chrome
  shows its default label; Edge may not display the managed favorites).
- The JSON preview updates live as you type in the field.

#### Build-Json updated
- Now prepends `{"toplevel_name":"<value>"}` as the first array element when the bar label
  field is non-empty. The remaining bookmarks follow as before.

#### Import-JsonString updated
- When loading or importing a JSON that starts with a bare `toplevel_name` entry, the value
  is extracted and placed back into the bar label field instead of being skipped silently.

## 2.4.0.0

### Bug fix — single root item produces object instead of array

#### Build-Json: pipe replaced by -InputObject
- When `$objects` contained only one element, piping it to `ConvertTo-Json` caused PowerShell
  to unroll the array and serialize a plain object `{"name":...,"children":[...]}` instead of
  the required array `[{"name":...,"children":[...]}]`.
- Both Chrome and Edge require the policy value to be a JSON **array** at the root; a plain
  object is silently ignored and no bookmarks appear.
- Fix: replaced `$objects | ConvertTo-Json` with `ConvertTo-Json -InputObject $objects` so the
  array wrapper is always preserved regardless of how many root items there are.

## 2.3.0.0

### Generated script — bug fix Edge key name + gpupdate

#### Bug fix: Edge uses ManagedFavorites, not ManagedBookmarks
- The working reference script uses `ManagedFavorites` as the registry value name under
  `HKLM:\SOFTWARE\Policies\Microsoft\Edge`. The generated script was incorrectly using
  `ManagedBookmarks` for Edge, which Edge does not recognise — causing the bar to appear
  empty. Fixed: Edge now receives `ManagedFavorites`; Chrome continues to use `ManagedBookmarks`.

#### Bug fix: Set-ItemProperty instead of New-ItemProperty
- Switched from `New-ItemProperty -Force` to `Set-ItemProperty` for all policy values,
  matching the pattern used in the verified working reference script.

#### Added: gpupdate /force after registry writes
- The generated script now runs `gpupdate /force` after all registry values are written so
  Chrome and Edge pick up the new policy in the same session.
- A yellow `ACTION REQUIRED` message reminds the user to close and reopen both browsers.

## 2.2.0.0

### Generated script — enable bookmarks/favorites bar

- Chrome: `BookmarkBarEnabled = 1` (DWORD) written to `HKLM:\SOFTWARE\Policies\Google\Chrome`
  so the bookmarks bar is shown by default.
- Edge: `FavoritesBarEnabled = 1` (DWORD) written to `HKLM:\SOFTWARE\Policies\Microsoft\Edge`
  so the favorites bar is shown by default.
- Both values are set in the same registry path as `ManagedBookmarks`, requiring no extra path creation.

## 2.1.0.0

### Export Script feature

#### New: 📜 Export Script toolbar button
- New toolbar button `📜 Export Script`, hidden by default.
- Becomes visible on form load only when the config is present **and** `Paths.ScriptBasePath` is set.
- When the config file is absent the button remains invisible and the feature is completely inaccessible.

#### New: Show-ExportScriptDialog
- Styled dark-theme dialog asking for a script name (spaces replaced with `_`, no extension)
  and an optional description for the generated script's `.DESCRIPTION` block.

#### New: New-ManagedBookmarksScript
- Generates a `.ps1` deployment script that follows the company scripting template convention
  (help block, origin-story table, transcript/log setup, admin check, OS detection, registry
  tracking, and the actual policy write).
- Sets `ManagedBookmarks` as a `String` registry value under both:
  - `HKLM:\SOFTWARE\Policies\Google\Chrome`
  - `HKLM:\SOFTWARE\Policies\Microsoft\Edge`
- The compact JSON produced by `Build-Json` is embedded as a single-quoted literal string
  so no variable expansion occurs at deploy time.
- All company-specific values (`$companyName`, `$registryPath`, transcript/log paths) come
  from the config file (`Company.Name`, `Company.RegistryBase`, `Logging.BasePath`).

#### Config additions (Decoster.tech.config.base.psd1)
- New `Company` section added: `Name`, `Author`, `RegistryBase`.
- Config validation updated: `Paths.ScriptBasePath`, `Company.Name`, `Company.Author` and
  `Company.RegistryBase` added to the required-key check on startup.

## 2.0.0.0

### Full English translation + Save dialog fix

#### Full English translation
- All Dutch UI text translated to English: toolbar buttons and tooltips, dialog titles and labels,
  status messages, log entries, MessageBox prompts, placeholder text, and comments.

#### Save dialog — bug fix and simplified flow
- **Root cause**: `Test-Path $GLB_jsonBasePath` returned `$false` when the folder did not exist yet
  (e.g. `c:\programdata\decoster.tech\scripting\Json` had never been created), causing the script
  to fall through to the standard Windows SaveFileDialog even though `JsonBasePath` was configured.
- **Fix**: removed the `Test-Path` guard from the save condition. When `JsonBasePath` is set in the
  config the custom `Show-SaveNameDialog` is always used; the folder is created automatically with
  `New-Item -Force` if it does not exist yet.
- **Fallback**: when `JsonBasePath` is not configured (no config file), the standard SaveFileDialog
  is shown as before.

## 1.8.1.0

### Config validation + Save dialog (name only)

#### Test-DecosterOutputConfig integrated
- After loading config and module, `Test-DecosterOutputConfig` is called with required keys
  `Logging.BasePath`, `Logging.LogRetentionDays` and `Paths.JsonBasePath`.
- Missing or empty keys are stored in `$script:configWarnings`.
- On form load the warning is shown in the status label (`actionLabel`) when the config is
  incomplete: `"Config incomplete — missing keys: ..."` in orange warning colour.
- No validation when the config file is absent (normal graceful degradation).

#### Show-SaveNameDialog
- New helper `Show-SaveNameDialog($savePath)`: shows a styled dark-theme dialog that asks only
  for a file name.
- Displays the target path at the top so the user knows where the file will be saved.
- Automatically appends `.json` if the entered name does not already have that extension.
- Asks for overwrite confirmation when the file already exists (Yes/No dialog).

## 1.8.0.0

### Config-validatie

#### Test-DecosterOutputConfig geïntegreerd
- Na het laden van de config én de module wordt `Test-DecosterOutputConfig` aangeroepen met de vereiste sleutels `Logging.BasePath`, `Logging.LogRetentionDays` en `Paths.JsonBasePath`.
- Ontbrekende of lege sleutels worden opgeslagen in `$script:configWarnings`.
- Bij form-load wordt de waarschuwing getoond in de status-label (`actionLabel`) als de config onvolledig is: `"Config onvolledig — ontbrekende sleutels: ..."` in oranje waarschuwingskleur.
- Geen validatie wanneer de config-file ontbreekt (normaal gedrag: graceful degradation).

## 1.7.0.0

### Config-integratie en PSScriptAnalyzer-fixes

#### JsonBasePath uit config geladen
- `$GLB_jsonBasePath` wordt uitgelezen uit `Paths.JsonBasePath` in `Decoster.tech.config.base.psd1`.
- `Laden`-dialog (`OpenFileDialog`) en `Opslaan`-dialog (`SaveFileDialog`) openen nu automatisch in `JsonBasePath` wanneer de config geladen is én de map bestaat; bij ontbrekende config of map valt de dialog terug op het laatste gebruikte pad (standaard Windows-gedrag).
- `JsonBasePath` gelogd bij opstarten via `Add-LogKeyValue "Config"`.
- `JsonBasePath` (`c:\programdata\decoster.tech\Json`) toegevoegd aan `Decoster.tech.config.base.psd1` met toelichting in de commentaar-header.

#### PSUseApprovedVerbs — hernoemen
- `Rebuild-Tree` → `Update-Tree` (overal: definitie + alle aanroepplaatsen).
- `Refresh-Preview` → `Update-Preview` (overal: definitie + alle aanroepplaatsen).

## 1.6.0.0

### Bug fix — Format-DecosterOutputStep not recognized / Lines null

#### Import-Module -Force toegevoegd
- De module werd gecached bij de eerste import in een PowerShell-sessie; nieuwe geëxporteerde functies (`Format-DecosterOutputStep`, `Format-DecosterOutputSubItem`) werden daardoor niet opgepikt zonder een sessie-herstart.
- `-Force` toegevoegd aan `Import-Module` zodat de module altijd opnieuw van schijf wordt geladen bij elke scriptrun.

#### Add-LogEntry defensief gemaakt
- De `switch`-body en `Add-DecosterOutputOutput`-aanroep omgeven met `try/catch`.
- Bij een onbekende of ontbrekende module-functie wordt de fout afgevangen en als gele `Write-Host`-lijn naar de console geschreven in plaats van de uitvoering te stoppen.
- Null-check op `$lines` vóór de aanroep van `Add-DecosterOutputOutput` om de `Cannot bind argument to parameter 'Lines' because it is null`-fout te voorkomen.

## 1.5.0.0

### Bug fix — JSON invoer doet niets bij wrapper-formaat

#### Import-JsonString herkent nu twee root-formaten
- **Policy-value array** (directe invoer voor de policy-sleutel):
  `[{"name":"Google","url":"google.com"},...]`
- **Volledig policy-object** (export uit ADMX-template of admin-tool):
  `{"BookmarkBarEnabled":true,"ManagedBookmarks":[...]}`

  Wanneer het root-object een `ManagedBookmarks`-eigenschap bevat, worden de bookmarks uit die eigenschap geëxtraheerd; `BookmarkBarEnabled` en eventuele andere sleutels worden genegeerd. De `toplevel_name`-entry bovenaan de `ManagedBookmarks`-array wordt (zoals in v1.4) automatisch overgeslagen.

- De JSON-uitvoer (Opslaan / Kopieer) blijft altijd de **directe array** — de waarde die je rechtstreeks in de `ManagedBookmarks`-policy plakt.

## 1.4.0.0

### Bug fixes — JSON importeren geeft name/url: null

#### Add-LogEntry parameter volgorde hersteld
- `$State` stond als vierde positieparameter; alle aanroepen van de vorm `Add-LogEntry "tekst" 'ok'` plaatsten de state-string per ongeluk in `$Detail` waardoor iedere log-lijn als `info` werd gelogd.
- `$State` teruggezet als tweede positieparameter (direct na `$Text`); `$Detail` en `$ErrorMessage` worden uitsluitend via named parameters meegegeven en staan nu als derde en vierde param.

#### Import-Node robuust gemaakt
- **`toplevel_name` ondersteuning**: Chrome's ManagedBookmarks JSON gebruikt `toplevel_name` (i.p.v. `name`) voor het label van de beheerde bladwijzerbalk. `Import-Node` gebruikt nu `name` bij voorkeur en valt terug op `toplevel_name`; entries met alleen `toplevel_name` en geen `children` of `url` worden overgeslagen (i.p.v. omgezet naar een URL-node met null-waarden).
- **`children`-detectie via PSObject.Properties**: de controle op een folder-node gebruikt nu `$jsonNode.PSObject.Properties.Name -contains 'children'` i.p.v. `$null -ne $jsonNode.children`. Dit onderscheidt correct een expliciete `"children": []` (lege folder) van een ontbrekende `children`-eigenschap (link).
- **PS 5.1 array-unwrapping**: `@($jsonNode.children)` toegevoegd in de `foreach` om te voorkomen dat PowerShell 5.1 een single-element children-array uitpakt naar een enkel object.
- **Null-guard**: als `$jsonNode` zelf `$null` is, geeft `Import-Node` direct `$null` terug.
- **Type-cast**: naam en url worden expliciet gecast naar `[string]` zodat `$null`-waarden als lege string doorgegeven worden i.p.v. als `$null`.

#### Import-JsonString robuust gemaakt
- `@($raw | ConvertFrom-Json)` toegevoegd zodat een single-element root-array in PS 5.1 als array behandeld wordt en niet als enkel PSCustomObject.
- Null-resultaten van `Import-Node` (bijv. overgeslagen `toplevel_name`-entries) worden gefilterd voor ze aan `$script:bookmarks` worden toegevoegd.

## 1.3.0.0

### Volledige integratie van Decoster.tech.output.psm1

#### Add-LogEntry uitgebreid
- Toegevoegd parameter `$Detail` — doorgegeven als `$Detail` aan `Format-DecosterOutputWarn` en als waarde bij `'keyvalue'`.
- Toegevoegd parameter `$ErrorMessage` — doorgegeven als `$ErrorMessage` aan `Format-DecosterOutputFail` en `Format-DecosterOutputWarn`; foutmeldingen verschijnen nu als tweede insprong-lijn in het logbestand in plaats van verloren te gaan.
- Uitgebreid `[ValidateSet]` met drie nieuwe states: `'keyvalue'`, `'step'`, `'subitem'`; elke state delegeert naar de overeenkomstige `Format-DecosterOutput*` functie uit de module.

#### Add-LogKeyValue helper
- Nieuwe helper `Add-LogKeyValue($Key, $Value)` als kortschrift voor `Add-LogEntry -Text $Key -Detail $Value -State 'keyvalue'`; maakt gestructureerde sleutel-waarde logging op één leesbare regel mogelijk.

#### Gestructureerde logging in alle handlers
- Startup-logging omgezet van `Format-DecosterOutputInfo`-strings naar `Add-LogKeyValue` voor Started, User, Author en Module; gevolgd door een divider.
- **Folder / Subfolder toevoegen**: naam en (voor subfolder) containerfolder gelogd als aparte key-value regels.
- **Link toevoegen**: naam, URL en containerfolder gelogd als afzonderlijke key-value regels.
- **Root-link toevoegen**: naam en URL gelogd als key-value regels.
- **Item bewerken**: naam en type gelogd als key-value regels.
- **Item verwijderen**: naam gelogd als key-value regel.
- **Bestand laden**: step-indicator vóór het laden, pad gelogd als key-value, root-itemtelling na succes; foutmelding via `$ErrorMessage` parameter (verschijnt als insprong-lijn).
- **Bestand opslaan**: step-indicator vóór het opslaan, pad en root-itemtelling als key-value; foutmelding via `$ErrorMessage`.
- **JSON-string importeren**: step-indicator vóór het parsen, root-itemtelling na succes; foutmelding via `$ErrorMessage`.
- **Sessie-einde**: tijdstip en root-itemtelling gelogd als key-value regels.

#### Nieuwe module-functies (zie ook module-changelog)
- Gebruikt `Format-DecosterOutputStep` voor bestand laden, opslaan en JSON-string importeren.
- Gebruikt `Format-DecosterOutputKeyValue` overal waar sleutel-waarde data gelogd wordt.

## 1.2.0.0

### Structural rewrite — aligned with ConvertToExe_GUI.PS1
- Added comment-based help block (`.SYNOPSIS`, `.DESCRIPTION`, `.INPUTS`, `.OUTPUTS`, `.EXAMPLE`, `.LINK`, `.NOTES`).
- Added global metadata variables: `$GLB_scriptVersion`, `$GLB_ScriptUpdateDate`, `$GLB_scriptcontributer`, `$GLB_ScriptTitel`.
- Added `#region`/`#endregion` sections throughout: load assemblies, Config, Output module, Logging setup, Colors & Fonts, Data model, Data helpers, TreeView helpers, Helper factories, Status helper, Dialogs, Form, Toolbar, Content panel, Status section, Signature, Layout, Button events, File operations, Write session log.
- Added config loading from `..\..\ConfigFiles\Decoster.tech.config.base.psd1`; graceful fallback when file is not found.
- Added output module loading from `..\..\Classes\Decoster.tech.output.psm1`; graceful fallback to built-in styling when module is not found.
- Added hidden `ListBox` log collector and `Add-LogEntry` helper that routes through `Format-DecosterOutput*` + `Add-DecosterOutputOutput`; no-op when module is not loaded.
- Added session log write after `$form.Dispose()` via `Write-DecosterOutputLog`; silently skipped when module or config are not available.
- Added log retention at startup via `Remove-DecosterOutputOldLogs`.
- All toolbar actions now log their result via `Add-LogEntry` and update the status label via `Set-StatusLabel`.
- Replaced original `Node-ToJson` with `ConvertTo-NodeObject` + `ConvertTo-Json -Compress -Depth 20` for compact, whitespace-free JSON output.
- Moved `Import-Node` to a top-level function so it can be shared between file load and string import.
- Applied dark theme (`#1E1E2E` background, `#2A2A3D` card, `#313145` input, `#0078D4` accent) to all controls including TreeView, TabControl and dialogs.
- Folder nodes now render in `#89B4FA` (blue); link nodes in `#CDD6F4` (light text) for visual contrast.
- Added `Decoster.tech` signature label in Consolas monospace font at the bottom of the form (green `#A6E3A1`).
- Added dynamic `Invoke-Layout` function wired to `$form.add_Load` and `$form.add_Resize` so all anchored controls (content panel, status section, signature) reflow correctly on resize.
- Replaced manual dialog construction with `New-StyledDialog`, `New-StyledTextBox` and `New-DialogLabel` factory helpers to eliminate repetition in `Show-FolderDialog` and `Show-LinkDialog`.
- Added `New-StyledButton` factory (flat style, hand cursor, configurable back/fore/font) used in dialogs and the import panel.
- Added `Set-StatusLabel` wrapper (delegates to `Set-DecosterOutputLabelStatus` when module is loaded; falls back to theme colors).
- Replaced `$form.ShowDialog()` return discard with `$null = $form.ShowDialog()` followed by `$form.Dispose()`.
- Toolbar `ToolStrip` now uses `System` render mode and hidden grip, with matching dark background and text colors.

## 1.1.0.0

### Compact JSON preview
- Replaced the indented `Node-ToJson` formatter with `ConvertTo-NodeObject` + `ConvertTo-Json -Compress -Depth 20` so the preview shows a single-line, whitespace-free JSON string — ready to paste directly into a GPO or Intune policy field.

### JSON Invoeren tab
- Replaced the single `TextBox`+`Label` JSON preview panel with a `TabControl` containing two tabs.
- **JSON Preview** tab: readonly compact output, dark-styled (existing behavior, now compact).
- **JSON Invoeren** tab: editable multiline `Consolas` textbox where an existing ManagedBookmarks JSON string can be pasted; **Importeer JSON** button parses the string via the shared `Import-JsonString` function, rebuilds the tree, and switches back to the Preview tab; **Wis** button clears the input field.
- Extracted common import logic into a top-level `Import-JsonString($raw)` function used by both the file-load dialog and the string import button.
- The "Laden" (file) button was simplified to delegate to `Import-JsonString` instead of containing its own inline `Import-Node` definition.

## 1.0.0.0

### Initial release
- PowerShell WinForms GUI for creating and editing a ManagedBookmarks JSON array compatible with Chrome and Edge managed browser policies.
- `ToolStrip` toolbar with actions: add root folder, add subfolder, add link (in folder), add root-level link, edit, delete, move up, move down, load from file, save to file, copy to clipboard.
- `TreeView` showing the bookmark hierarchy; folders displayed in bold blue, links in black.
- Folder nodes store `type`, `name` and a `List[object]` of children; link nodes store `type`, `name` and `url`.
- `Show-FolderDialog` and `Show-LinkDialog` modal dialogs for entering/editing node properties.
- `Find-ParentList` / `Search-Children` helpers to locate a node's owning list for delete and reorder operations.
- Move up / move down operations rebuild the full tree and re-select the moved node via `Select-ByData`.
- `Rebuild-Tree` recursively repopulates the `TreeView` from `$script:bookmarks`.
- Double-clicking a tree node triggers the edit dialog.
- `OpenFileDialog` for loading an existing `managed_bookmarks.json`; `SaveFileDialog` for saving output.
- JSON generated via indented `Node-ToJson` (replaced in v1.1 with compact formatter).
- Copy to clipboard button.
- `SplitContainer` (horizontal) divides the tree (top) from the JSON preview (bottom).


