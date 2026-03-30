<#
    .SYNOPSIS
        GUI tool to create and edit a ManagedBookmarks JSON for Chrome / Edge policies.

    .DESCRIPTION
        Provides a tree view to build folders, sub-folders and bookmark links.
        The resulting JSON can be saved to a file, copied to the clipboard,
        or pasted back in to continue editing an existing policy value.

    .INPUTS
        None — all input is handled through the GUI.

    .OUTPUTS
        A JSON array compatible with the Chrome/Edge ManagedBookmarks policy.

    .EXAMPLE
        Run the script, add folders and links, then click Save or Copy.

    .LINK
        CHANGELOG.md

    .LINK
        LICENSE

    .NOTES
        Licensed under the MIT License; see LICENSE in this folder.
#>

$GLB_scriptVersion     = "2.12.0.0"
$GLB_ScriptUpdateDate  = "30/03/2026"
$GLB_scriptcontributer = "Decoster Hans"
$GLB_ScriptTitel       = "Managed Bookmarks Creator"

#region load assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#endregion

function Test-HasText {
    param([string]$Value)
    # Central null/empty/whitespace check used throughout the script to keep
    # guard clauses short and consistent.
    return -not [string]::IsNullOrWhiteSpace($Value)
}

function ConvertTo-FriendlyErrorMessage {
    param([string]$RawMessage)

    # Some PowerShell/UI failures arrive as CLIXML payloads instead of plain
    # text. This helper extracts the readable error lines for message boxes.

    if (-not (Test-HasText $RawMessage)) { return 'Unknown error.' }

    $msg = $RawMessage
    if ($msg -match '^#<\s*CLIXML') {
        $errors = [System.Text.RegularExpressions.Regex]::Matches($msg, '<S\s+S="Error">(.*?)</S>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        if ($errors.Count -gt 0) {
            $parts = @()
            foreach ($m in $errors) {
                $part = $m.Groups[1].Value
                $part = $part -replace '_x000D__x000A_', [System.Environment]::NewLine
                $part = $part.Trim()
                if (Test-HasText $part) { $parts += $part }
            }
            if ($parts.Count -gt 0) {
                return ($parts -join [System.Environment]::NewLine)
            }
        }
    }

    return $msg.Trim()
}

function Test-IsValidHttpUrl {
    param([string]$Url)
    # Managed bookmarks should contain absolute web URLs only; reject custom
    # schemes and malformed strings early in the dialog.
    $uri = $null
    if (-not [System.Uri]::TryCreate($Url, [System.UriKind]::Absolute, [ref]$uri)) { return $false }
    return $uri.Scheme -in @('http', 'https')
}

function ConvertTo-SafeFileBaseName {
    param([string]$InputName)

    # Normalize free-form user input into a Windows-safe file base name so the
    # generated JSON/script names do not fail on invalid path characters.

    if (-not (Test-HasText $InputName)) { return '' }

    $name = $InputName.Trim() -replace '\s+', '_'
    $name = $name -replace '[\\/:*?"<>|]', ''
    $name = $name -replace '[^A-Za-z0-9_.-]', ''
    $name = $name.Trim('. ')

    $reserved = @('CON','PRN','AUX','NUL','COM1','COM2','COM3','COM4','COM5','COM6','COM7','COM8','COM9','LPT1','LPT2','LPT3','LPT4','LPT5','LPT6','LPT7','LPT8','LPT9')
    if ($reserved -contains $name.ToUpperInvariant()) {
        $name = "${name}_file"
    }

    if ($name.Length -gt 120) {
        $name = $name.Substring(0, 120).Trim('. ')
    }

    return $name
}

$script:isPackagedExecutable = $false
try {
    $currentProcessPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $processName = [System.IO.Path]::GetFileNameWithoutExtension($currentProcessPath)
    $script:isPackagedExecutable = @('powershell', 'pwsh', 'powershell_ise') -notcontains $processName
}
catch {
    $script:isPackagedExecutable = $false
}

function Write-AppHostMessage {
    param(
        [string]$Message,
        [string]$Color = 'Gray'
    )

    # Avoid console chatter when the script is packaged as a GUI executable.
    # During normal .ps1 runs, information messages remain useful for startup diagnostics.
    if ($script:isPackagedExecutable) { return }
    Write-Information "[$Color] $Message" -InformationAction Continue
}

function Get-AppSearchRoots {
    # Build an ordered list of directories where assets may live. This allows
    # the same script to find config/module files whether it runs as source,
    # from a packaged executable, or from a different working directory.
    $roots = [System.Collections.Generic.List[string]]::new()
    foreach ($candidate in @(
            $PSScriptRoot,
            [System.Windows.Forms.Application]::StartupPath,
            (Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)),
            (Get-Location).Path
        )) {
        if ((Test-HasText $candidate) -and (Test-Path -LiteralPath $candidate) -and -not $roots.Contains($candidate)) {
            $roots.Add($candidate) | Out-Null
        }
    }
    return $roots
}

function Resolve-AppAssetPath {
    param(
        [string[]]$RelativeCandidates
    )

    # Try each candidate relative path against each approved root and return
    # the first real file that exists.
    foreach ($root in (Get-AppSearchRoots)) {
        foreach ($relativePath in $RelativeCandidates) {
            if (-not (Test-HasText $relativePath)) { continue }
            $candidate = [System.IO.Path]::GetFullPath((Join-Path $root $relativePath))
            if (Test-Path -LiteralPath $candidate) {
                return $candidate
            }
        }
    }

    return $null
}
#region ── Config ────────────────────────────────────────────────────────────
# Supports both:
# - running the .ps1 from GUI\ManagedBookmarksCreator
# - running a converted .exe from the same folder
$GLB_jsonBasePath      = $null
$GLB_scriptBasePath    = $null
$GLB_companyName       = $null
$GLB_companyAuthor     = $null
$GLB_companyRegBase    = $null
$GLB_config            = $null
$GLB_configPath        = Resolve-AppAssetPath @(
    'Decoster.tech.config.base.psd1',
    'ConfigFiles\Decoster.tech.config.base.psd1',
    '..\ConfigFiles\Decoster.tech.config.base.psd1',
    '..\..\ConfigFiles\Decoster.tech.config.base.psd1'
)
try {
    if (-not (Test-HasText $GLB_configPath)) {
        throw "Config file not found."
    }
    $GLB_config        = Import-PowerShellDataFile -Path $GLB_configPath -ErrorAction Stop
    $GLB_jsonBasePath  = $GLB_config.Paths.JsonBasePath
    $GLB_scriptBasePath = $GLB_config.Paths.ScriptBasePath
    $GLB_companyName   = $GLB_config.Company.Name
    $GLB_companyAuthor = $GLB_config.Company.Author
    $GLB_companyRegBase = $GLB_config.Company.RegistryBase
    Write-AppHostMessage "Config loaded." "Cyan"
}
catch {
    Write-AppHostMessage "Config not found — JSON will be saved to the last used folder." "Yellow"
}
#endregion

#region ── Output module ─────────────────────────────────────────────────────
$GLB_outputModulePath      = Resolve-AppAssetPath @(
    'Decoster.tech.output.psm1',
    'Classes\Decoster.tech.output.psm1',
    '..\Classes\Decoster.tech.output.psm1',
    '..\..\Classes\Decoster.tech.output.psm1'
)
$script:outputModuleLoaded = $false
try {
    # -Force ensures an already-loaded module is always reloaded from disk,
    # picking up any new exported functions added since the session started.
    if (-not (Test-HasText $GLB_outputModulePath)) {
        throw "Output module not found."
    }
    Import-Module $GLB_outputModulePath -Force -ErrorAction Stop
    $script:outputModuleLoaded = $true
    Write-AppHostMessage "Output module loaded." "Cyan"
}
catch {
    Write-AppHostMessage "Output module not found — using built-in styling." "Yellow"
}
#endregion

#region ── Config validation ─────────────────────────────────────────────────
# Only runs when both the config AND the module are available.
# $script:configWarnings is shown in the status label on form load.
$script:configWarnings = @()
if ($null -ne $GLB_config -and $script:outputModuleLoaded) {
    $script:configWarnings = Test-DecosterOutputConfig -Config $GLB_config `
        -RequiredKeys @('Logging.BasePath','Logging.LogRetentionDays','Paths.JsonBasePath',
                        'Paths.ScriptBasePath','Company.Name','Company.Author','Company.RegistryBase')
    if ($script:configWarnings.Count -gt 0) {
        Write-AppHostMessage "Config incomplete — missing keys: $($script:configWarnings -join ', ')" "Yellow"
    }
}
#endregion

#region ── Logging setup ─────────────────────────────────────────────────────
$logListBox  = New-Object System.Windows.Forms.ListBox
$GLB_logPath = $null

$fallbackLogRoot = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Decoster.tech\ManagedBookmarksCreator\Logs'
$fallbackLogName = '{0}_{1}_ManagedBookmarksCreator.log' -f $env:USERNAME, (Get-Date -Format 'yyyyMMdd_HHmmss')
$script:fallbackLogPath = Join-Path $fallbackLogRoot $fallbackLogName

if ($script:outputModuleLoaded -and $GLB_config -and $GLB_config.Logging.BasePath) {
    $GLB_logPath = Join-Path $GLB_config.Logging.BasePath "ManagedBookmarksCreator"
    Remove-DecosterOutputOldLogs -LogPath $GLB_logPath -RetentionDays $GLB_config.Logging.LogRetentionDays
}
#endregion

function Add-FallbackLogLine {
    param(
        [string]$Text         = '',
        [string]$State        = 'info',
        [string]$Detail       = '',
        [string]$ErrorMessage = ''
    )

    if (-not (Test-HasText $script:fallbackLogPath)) { return }

    try {
        $directory = Split-Path -Parent $script:fallbackLogPath
        if ((Test-HasText $directory) -and -not (Test-Path -LiteralPath $directory)) {
            New-Item -ItemType Directory -Path $directory -Force -ErrorAction Stop | Out-Null
        }

        $parts = [System.Collections.Generic.List[string]]::new()
        if (Test-HasText $Text)         { $parts.Add($Text) | Out-Null }
        if (Test-HasText $Detail)       { $parts.Add("Detail: $Detail") | Out-Null }
        if (Test-HasText $ErrorMessage) { $parts.Add("Error: $ErrorMessage") | Out-Null }
        if ($parts.Count -eq 0) { $parts.Add('') | Out-Null }

        $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $State.ToUpperInvariant(), ($parts -join ' | ')
        Add-Content -LiteralPath $script:fallbackLogPath -Value $line -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        # Do not let diagnostics logging interrupt the main GUI flow.
    }
}

function Mark-ErrorRecordSeen {
    param($Record)

    if ($null -eq $Record) { return }

    try {
        $recordId = [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($Record)
        $null = $script:seenErrorRecordIds.Add($recordId)
    }
    catch {
        # Best-effort deduplication only.
    }
}

function ConvertTo-ErrorRecordLogMessage {
    param([System.Management.Automation.ErrorRecord]$Record)

    if ($null -eq $Record) { return 'Unknown PowerShell error.' }

    $parts = [System.Collections.Generic.List[string]]::new()
    $friendly = ConvertTo-FriendlyErrorMessage $Record.Exception.Message
    if (Test-HasText $friendly) { $parts.Add($friendly) | Out-Null }

    if ($null -ne $Record.CategoryInfo -and (Test-HasText ([string]$Record.CategoryInfo.Category))) {
        $parts.Add("Category: $($Record.CategoryInfo.Category)") | Out-Null
    }
    if ($null -ne $Record.InvocationInfo -and (Test-HasText $Record.InvocationInfo.MyCommand.Name)) {
        $parts.Add("Command: $($Record.InvocationInfo.MyCommand.Name)") | Out-Null
    }
    if ($null -ne $Record.TargetObject -and (Test-HasText ([string]$Record.TargetObject))) {
        $parts.Add("Target: $($Record.TargetObject)") | Out-Null
    }
    if ($null -ne $Record.InvocationInfo -and (Test-HasText $Record.InvocationInfo.PositionMessage)) {
        $position = ($Record.InvocationInfo.PositionMessage -replace '\r?\n', ' | ').Trim()
        if (Test-HasText $position) {
            $parts.Add("Position: $position") | Out-Null
        }
    }

    return ($parts -join ' || ')
}

function Sync-GlobalErrorLog {
    $records = @($global:Error)
    if ($records.Count -eq 0) { return }

    [array]::Reverse($records)
    foreach ($record in $records) {
        if ($null -eq $record) { continue }

        try {
            $recordId = [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($record)
        }
        catch {
            continue
        }

        if (-not $script:seenErrorRecordIds.Add($recordId)) { continue }
        Add-LogEntry 'PowerShell runtime error' 'fail' -ErrorMessage (ConvertTo-ErrorRecordLogMessage $record)
    }
}

function Start-ErrorMonitoring {
    foreach ($existingError in @($global:Error)) {
        Mark-ErrorRecordSeen $existingError
    }

    if ($null -eq $script:errorMonitorTimer) {
        $script:errorMonitorTimer = New-Object System.Windows.Forms.Timer
        $script:errorMonitorTimer.Interval = 1000
        $script:errorMonitorTimer.Add_Tick({ Sync-GlobalErrorLog })
    }
    $script:errorMonitorTimer.Start()

    if ($null -eq $script:threadExceptionHandler) {
        # Do not call SetUnhandledExceptionMode here. In packaged hosts this can emit
        # a MethodInvocation error popup even when wrapped, depending on host behavior.
        # We still attach thread/appdomain exception handlers for best-effort diagnostics.
        $script:threadExceptionHandler = [System.Threading.ThreadExceptionEventHandler]{
            param($_sender, $eventArgs)

            if ($null -eq $eventArgs -or $null -eq $eventArgs.Exception) { return }

            Add-LogEntry 'Unhandled UI exception' 'fail' -ErrorMessage $eventArgs.Exception.ToString()
            if ($null -ne $actionLabel) {
                Set-StatusLabel $actionLabel 'Unhandled UI exception logged. Check the session log for details.' 'fail'
            }
        }
        [System.Windows.Forms.Application]::add_ThreadException($script:threadExceptionHandler)
    }

    if ($null -eq $script:unhandledExceptionHandler) {
        $script:unhandledExceptionHandler = [System.UnhandledExceptionEventHandler]{
            param($_sender, $eventArgs)

            $exceptionText = if ($null -ne $eventArgs -and $eventArgs.ExceptionObject -is [System.Exception]) {
                $eventArgs.ExceptionObject.ToString()
            }
            elseif ($null -ne $eventArgs) {
                [string]$eventArgs.ExceptionObject
            }
            else {
                'Unknown unhandled application exception.'
            }

            Add-LogEntry 'Unhandled application exception' 'fail' -ErrorMessage $exceptionText
        }
        [AppDomain]::CurrentDomain.add_UnhandledException($script:unhandledExceptionHandler)
    }
}

function Stop-ErrorMonitoring {
    if ($null -ne $script:errorMonitorTimer) {
        $script:errorMonitorTimer.Stop()
        $script:errorMonitorTimer.Dispose()
        $script:errorMonitorTimer = $null
    }

    if ($null -ne $script:threadExceptionHandler) {
        [System.Windows.Forms.Application]::remove_ThreadException($script:threadExceptionHandler)
        $script:threadExceptionHandler = $null
    }

    if ($null -ne $script:unhandledExceptionHandler) {
        [AppDomain]::CurrentDomain.remove_UnhandledException($script:unhandledExceptionHandler)
        $script:unhandledExceptionHandler = $null
    }
}

function Add-LogEntry {
    <#
    .SYNOPSIS  Routes a log line through the output module (no-op when not loaded).
    .PARAMETER Text
        Primary message or — for 'keyvalue' — the key name.
    .PARAMETER Detail
        For 'warn'     : passed as $Detail to Format-DecosterOutputWarn.
        For 'keyvalue' : the value (paired with $Text as the key).
        For 'subitem'  : ignored ($Text is used directly).
        Ignored for all other states.
    .PARAMETER ErrorMessage
        For 'fail' : passed as $ErrorMessage to Format-DecosterOutputFail.
        For 'warn' : passed as $ErrorMessage to Format-DecosterOutputWarn.
        Ignored for all other states.
    .PARAMETER State
        Log level / format selector.
    #>
    param(
        [string]$Text         = '',
        [ValidateSet('ok','fail','warn','error','info','header',
                     'divider','separator','keyvalue','step','subitem')][string]$State = 'info',
        [string]$Detail       = '',
        [string]$ErrorMessage = ''
    )
    Add-FallbackLogLine -Text $Text -State $State -Detail $Detail -ErrorMessage $ErrorMessage
    if (-not $script:outputModuleLoaded) { return }
    try {
        $lines = switch ($State) {
            'ok'        { Format-DecosterOutputOk       $Text }
            'fail'      { Format-DecosterOutputFail      $Text  $ErrorMessage }
            'warn'      { Format-DecosterOutputWarn      $Text  $Detail  $ErrorMessage }
            'error'     { Format-DecosterOutputError     $Text }
            'header'    { Format-DecosterOutputHeader    $Text }
            'divider'   { Format-DecosterOutputDivider        }
            'separator' { Format-DecosterOutputSeparator      }
            'keyvalue'  { Format-DecosterOutputKeyValue  $Text  $Detail }
            'step'      { Format-DecosterOutputStep      $Text }
            'subitem'   { Format-DecosterOutputSubItem   $Text }
            default     { Format-DecosterOutputInfo      $Text }
        }
        if ($null -ne $lines) {
            Add-DecosterOutputOutput -Listbox $logListBox -Lines $lines
        }
    }
    catch {
        Mark-ErrorRecordSeen $_
        Write-AppHostMessage "[Log] $State : $Text $(if($Detail){"| $Detail"}) — $($_.Exception.Message)" "DarkYellow"
        Add-FallbackLogLine -Text 'Fallback logger pipeline error' -State 'warn' -Detail $Text -ErrorMessage $_.Exception.Message
    }
}

# Shorthand: Add-LogKeyValue "Name" "Intranet"
function Add-LogKeyValue($Key, $Value) {
    Add-LogEntry -Text $Key -Detail $Value -State 'keyvalue'
}

$isAdmin     = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$adminPrefix = if ($isAdmin) { "Admin: " } else { "" }

Add-LogEntry "$GLB_ScriptTitel  v$GLB_scriptVersion" 'header'
Add-LogEntry '' 'divider'
Add-LogKeyValue "Started" "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
Add-LogKeyValue "User"    "$env:USERNAME"
Add-LogKeyValue "Admin"   "$(if ($isAdmin) { 'yes — elevated' } else { 'no' })"
Add-LogKeyValue "Author"  "$GLB_scriptcontributer"
Add-LogKeyValue "Module"  "$(if ($script:outputModuleLoaded) { 'loaded' } else { 'not found — built-in styling' })"
Add-LogKeyValue "Config"       "$(if ($GLB_jsonBasePath) { "loaded — JsonBasePath: $GLB_jsonBasePath" } else { 'not found — no default path' })"
Add-LogKeyValue "ScriptPath"   "$(if ($GLB_scriptBasePath) { $GLB_scriptBasePath } else { 'not configured' })"
Add-LogEntry '' 'divider'

#region Colors & Fonts
$clrBg      = [System.Drawing.ColorTranslator]::FromHtml("#1E1E2E")
$clrCard    = [System.Drawing.ColorTranslator]::FromHtml("#2A2A3D")
$clrInput   = [System.Drawing.ColorTranslator]::FromHtml("#313145")
$clrBorder  = [System.Drawing.ColorTranslator]::FromHtml("#3D3D5C")
$clrText    = [System.Drawing.ColorTranslator]::FromHtml("#CDD6F4")
$clrMuted   = [System.Drawing.ColorTranslator]::FromHtml("#6C7086")
$clrAccent  = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
$clrSuccess = [System.Drawing.ColorTranslator]::FromHtml("#A6E3A1")
$clrError   = [System.Drawing.ColorTranslator]::FromHtml("#F38BA8")
$clrWarn    = [System.Drawing.ColorTranslator]::FromHtml("#FAB387")
$clrSig     = [System.Drawing.ColorTranslator]::FromHtml("#A6E3A1")
$clrFolder  = [System.Drawing.ColorTranslator]::FromHtml("#89B4FA")

$fontMain = New-Object System.Drawing.Font("Segoe UI", 9)
$fontBold = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$fontMono = New-Object System.Drawing.Font("Consolas", 8)
#endregion

#region ── Data model ────────────────────────────────────────────────────────
# Folder : @{ type="folder"; name="..."; children=List[object] }
# Link   : @{ type="url";    name="..."; url="..."             }
$script:bookmarks  = [System.Collections.Generic.List[object]]::new()
$script:undoStack  = [System.Collections.Generic.List[string]]::new()
$script:redoStack  = [System.Collections.Generic.List[string]]::new()
$script:dragNode   = $null   # TreeNode being dragged
$script:dropTarget = $null   # TreeNode currently highlighted as drop zone
$script:dropBefore = $true   # insert before ($true) or after ($false) drop target

# -----------------------------------------------------------------------------
# Autosave / Session-Recovery runtime state
# -----------------------------------------------------------------------------
# The recovery system is intentionally file-based so it still works when the
# PowerShell host is terminated unexpectedly. We persist a compact editor-state
# JSON snapshot and use a lock file as a crash indicator.
#
# Key concepts:
#   1) latest.state.json  -> the most recent editor snapshot
#   2) session.lock       -> created while app is running; removed on clean exit
#   3) dirty flag         -> saves only when state changed (debounced writes)
#
# This design keeps IO low during normal editing and maximizes recoverability
# after a crash/kill/restart.
$script:recoveryRootPath          = $null
$script:recoveryStateFilePath     = $null
$script:recoveryLockFilePath      = $null
$script:recoveryActive            = $false
$script:autosaveIsDirty           = $false
$script:isRestoringEditorState    = $false
$script:autosaveTimer             = $null
$script:autosaveIntervalMs        = 15000

# Runtime diagnostics state.
# - fallback log path keeps a plain-text audit trail even when the optional
#   output module is unavailable.
# - seen error ids prevent the periodic $Error scan from logging the same
#   ErrorRecord object multiple times.
$script:fallbackLogPath           = $null
$script:seenErrorRecordIds        = [System.Collections.Generic.HashSet[int]]::new()
$script:errorMonitorTimer         = $null
$script:threadExceptionHandler    = $null
$script:unhandledExceptionHandler = $null
#endregion

#region ── Data helpers ──────────────────────────────────────────────────────
function New-FolderNode($name) {
    # Data-model constructor for folder items. The TreeView stores these hashtables
    # in Node.Tag so UI operations and JSON serialization work on the same object.
    return @{ type = "folder"; name = $name; children = [System.Collections.Generic.List[object]]::new() }
}

function New-UrlNode($name, $url) {
    # Data-model constructor for bookmark link items.
    return @{ type = "url"; name = $name; url = $url }
}

function ConvertTo-NodeObject($node) {
    # Convert the internal hashtable/tree model into plain PSCustomObjects that
    # serialize cleanly to JSON without leaking WinForms-specific state.
    if ($node.type -eq "folder") {
        return [PSCustomObject]@{
            name     = $node.name
            children = @($node.children | ForEach-Object { ConvertTo-NodeObject $_ })
        }
    }
    return [PSCustomObject]@{ name = $node.name; url = $node.url }
}

function Build-Json {
    # Build the final policy payload. The optional toplevel_name entry must be
    # the first array element because Edge expects that specific structure.
    $bookmarkObjects = @($script:bookmarks | ForEach-Object { ConvertTo-NodeObject $_ })
    $name = if ($null -ne $script:txtTopLevelName) { $script:txtTopLevelName.Text.Trim() } else { '' }
    if ($name -ne '') {
        $all = @([PSCustomObject]@{ toplevel_name = $name }) + $bookmarkObjects
    } else {
        $all = $bookmarkObjects
    }
    return (ConvertTo-Json -InputObject $all -Compress -Depth 20)
}

function Get-EditorStateJson {
    # Capture the full editor state in one compact payload so undo/redo and
    # session recovery can restore both the tree and the bar label together.
    $topName = if ($null -ne $script:txtTopLevelName) { $script:txtTopLevelName.Text } else { '' }
    $state = [PSCustomObject]@{
        toplevel_name = $topName
        bookmarks     = @($script:bookmarks | ForEach-Object { ConvertTo-NodeObject $_ })
    }
    return ($state | ConvertTo-Json -Compress -Depth 25)
}

function Restore-EditorStateJson {
    param([string]$StateJson)

    # Rebuild the editor from a saved snapshot. The restore flag prevents the
    # recovery/undo plumbing from treating programmatic changes as user edits.
    try {
        $script:isRestoringEditorState = $true

        $state = $StateJson | ConvertFrom-Json
        $script:bookmarks.Clear()
        $script:tree.Nodes.Clear()

        if ($null -ne $script:txtTopLevelName) {
            $script:txtTopLevelName.Text = [string]$state.toplevel_name
        }

        foreach ($item in @($state.bookmarks)) {
            $node = Import-Node $item
            if ($null -ne $node) {
                $script:bookmarks.Add($node) | Out-Null
            }
        }

        Update-Tree
        Update-Preview
    }
    finally {
        $script:isRestoringEditorState = $false
    }
}

function Push-UndoState {
    # Save the current state before a mutating action. Redo history is cleared
    # because a new edit creates a new branch of history.
    $script:undoStack.Add((Get-EditorStateJson)) | Out-Null
    if ($script:undoStack.Count -gt 100) {
        $script:undoStack.RemoveAt(0)
    }
    $script:redoStack.Clear()
    if (-not $script:isRestoringEditorState) {
        $script:autosaveIsDirty = $true
    }
}

function Invoke-Undo {
    # Undo swaps the current state into redo and restores the latest undo snapshot.
    if ($script:undoStack.Count -eq 0) { return $false }
    $script:redoStack.Add((Get-EditorStateJson)) | Out-Null
    $last = $script:undoStack[$script:undoStack.Count - 1]
    $script:undoStack.RemoveAt($script:undoStack.Count - 1)
    Restore-EditorStateJson $last
    return $true
}

function Invoke-Redo {
    # Redo mirrors undo: move current state back to undo, then restore the most
    # recently undone snapshot.
    if ($script:redoStack.Count -eq 0) { return $false }
    $script:undoStack.Add((Get-EditorStateJson)) | Out-Null
    $next = $script:redoStack[$script:redoStack.Count - 1]
    $script:redoStack.RemoveAt($script:redoStack.Count - 1)
    Restore-EditorStateJson $next
    return $true
}

function Test-JsonRoundtripContract {
    # Safety test: export the current state to JSON, import it again, then verify
    # the normalized JSON is unchanged. This catches drift between build/import logic.
    $snapshot = Get-EditorStateJson
    try {
        $before = Build-Json
        Import-JsonString $before
        $after = Build-Json

        $beforeNorm = (($before | ConvertFrom-Json) | ConvertTo-Json -Compress -Depth 25)
        $afterNorm  = (($after  | ConvertFrom-Json) | ConvertTo-Json -Compress -Depth 25)
        return ($beforeNorm -eq $afterNorm)
    }
    finally {
        Restore-EditorStateJson $snapshot
    }
}

function Import-Node($jsonNode) {
    if ($null -eq $jsonNode) { return $null }

    # Chrome ManagedBookmarks uses 'toplevel_name' for the bookmarks-bar label entry.
    # Prefer 'name'; fall back to 'toplevel_name'; default to empty string.
    $nodeName = if ($null -ne $jsonNode.name)              { [string]$jsonNode.name }
                elseif ($null -ne $jsonNode.toplevel_name) { [string]$jsonNode.toplevel_name }
                else                                       { '' }

    # Folder: has a 'children' property (even when empty)
    if ($jsonNode.PSObject.Properties.Name -contains 'children') {
        $folder = New-FolderNode $nodeName
        # @() ensures PS 5.1 single-element arrays are not unwrapped
        foreach ($child in @($jsonNode.children)) {
            $childNode = Import-Node $child
            if ($null -ne $childNode) {
                $folder.children.Add($childNode) | Out-Null
            }
        }
        return $folder
    }

    # Skip entries that have no usable name and no url
    # (e.g. a bare toplevel_name-only header with no children)
    if ($nodeName -eq '' -and $null -eq $jsonNode.url) { return $null }

    return New-UrlNode $nodeName ([string]$jsonNode.url)
}

function Import-JsonString($raw) {
    $parsed = $raw | ConvertFrom-Json

    # Two supported root formats:
    #   1. Policy-value array   : [{...}, {...}]
    #   2. Full policy object   : {"ManagedBookmarks":[...],"BookmarkBarEnabled":true}
    #                           : {"ManagedFavorites":[...],"FavoritesBarEnabled":true}
    #      (as exported by some admin tools / ADMX templates)
    if ($parsed.PSObject.Properties.Name -contains 'ManagedBookmarks') {
        $items = @($parsed.ManagedBookmarks)
    } elseif ($parsed.PSObject.Properties.Name -contains 'ManagedFavorites') {
        $items = @($parsed.ManagedFavorites)
    } else {
        # @() keeps PS 5.1 from unwrapping a single-element root array into a bare object
        $items = @($parsed)
    }

    # Extract toplevel_name from the first element if it is a bare label entry
    $extractedName = ''
    if ($items.Count -gt 0) {
        $firstProps = $items[0].PSObject.Properties.Name
        if ($firstProps -contains 'toplevel_name' -and
            $firstProps -notcontains 'url' -and
            $firstProps -notcontains 'children') {
            $extractedName = [string]$items[0].toplevel_name
            $items = if ($items.Count -gt 1) { @($items[1..($items.Count - 1)]) } else { @() }
        }
    }
    if ($null -ne $script:txtTopLevelName) { $script:txtTopLevelName.Text = $extractedName }

    $script:bookmarks.Clear()
    $script:tree.Nodes.Clear()
    foreach ($item in $items) {
        $node = Import-Node $item
        if ($null -ne $node) {
            $script:bookmarks.Add($node) | Out-Null
        }
    }
    Update-Tree
    Update-Preview
}
#endregion

#region ── TreeView helpers ──────────────────────────────────────────────────
function Add-TreeNode($parentTvNode, $dataNode) {
    # Create the visible TreeView node and keep a reference to the underlying
    # data object in Tag so later UI actions can mutate the real model directly.
    $tv = New-Object System.Windows.Forms.TreeNode
    $tv.Tag = $dataNode
    Update-TreeNodeText $tv
    if ($null -eq $parentTvNode) {
        $script:tree.Nodes.Add($tv) | Out-Null
    } else {
        $parentTvNode.Nodes.Add($tv) | Out-Null
    }
    return $tv
}

function Update-TreeNodeText($tvNode) {
    # Visual formatting is derived from the data type so folder/link styling
    # always stays in sync after edits.
    $d = $tvNode.Tag
    if ($d.type -eq "folder") {
        $tvNode.Text      = "[Folder] $($d.name)"
        $tvNode.ForeColor = $clrFolder
        $tvNode.NodeFont  = $fontBold
    } else {
        $tvNode.Text      = "$($d.name)  ($($d.url))"
        $tvNode.ForeColor = $clrText
        $tvNode.NodeFont  = $fontMain
    }
}

function Update-Tree {
    # Full tree rebuild from the in-memory model. This is simpler and safer than
    # trying to patch every UI branch manually after complex operations.
    $script:tree.Nodes.Clear()
    foreach ($node in $script:bookmarks) {
        $tv = Add-TreeNode $null $node
        if ($node.type -eq "folder") { Add-ChildrenToTree $tv $node }
    }
    $script:tree.ExpandAll()
}

function Add-ChildrenToTree($tvParent, $dataParent) {
    # Recursive renderer for nested folders.
    foreach ($child in $dataParent.children) {
        $tv = Add-TreeNode $tvParent $child
        if ($child.type -eq "folder") { Add-ChildrenToTree $tv $child }
    }
}

function Find-ParentList($dataNode, [ref]$outList, [ref]$outParent) {
    # Find which collection currently owns a node. This is needed for delete and
    # move operations because nodes can live either at root level or inside folders.
    foreach ($n in $script:bookmarks) {
        if ($n -eq $dataNode) { $outList.Value = $script:bookmarks; $outParent.Value = $null; return $true }
        if ($n.type -eq "folder" -and (Search-Children $n $dataNode $outList $outParent)) { return $true }
    }
    return $false
}

function Search-Children($folder, $target, [ref]$outList, [ref]$outParent) {
    # Depth-first search through nested folder children until the target item is found.
    foreach ($c in $folder.children) {
        if ($c -eq $target) { $outList.Value = $folder.children; $outParent.Value = $folder; return $true }
        if ($c.type -eq "folder" -and (Search-Children $c $target $outList $outParent)) { return $true }
    }
    return $false
}

function Select-ByData($tvNode, $data) {
    # After rebuilding the tree, reselect the node that still points to the same
    # underlying data object so keyboard/move workflows feel continuous.
    if ($tvNode.Tag -eq $data) { $script:tree.SelectedNode = $tvNode; return }
    foreach ($child in $tvNode.Nodes) { Select-ByData $child $data }
}

function Clear-DropHighlight {
    if ($null -ne $script:dropTarget) {
        $script:dropTarget.BackColor = [System.Drawing.Color]::Empty
        $script:dropTarget.ForeColor = $clrText
        $script:dropTarget = $null
    }
}
#endregion

#region ── Script generator ──────────────────────────────────────────────────
function New-ManagedBookmarksScript {
    param(
        [string]$ScriptName,
        [string]$Description,
        [string]$JsonValue,
        [string]$CompanyName,
        [string]$Author,
        [string]$RegistryBase,
        [string]$ScriptingRoot,
        [string]$Date,
        [string]$Version
    )

    # Build the content with explicit string concatenation to avoid here-string escaping issues.
    # Each $-variable in the GENERATED script is written as a literal string.
    $nl = [System.Environment]::NewLine

    $escapedJsonValue = $JsonValue -replace "'", "''"

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add('<# ')
    $lines.Add('.SYNOPSIS ')
    $lines.Add(" Set ManagedBookmarks policy for Chrome and Edge")
    $lines.Add(' ')
    $lines.Add('.DESCRIPTION ')
    $lines.Add(" $Description")
    $lines.Add(' ')
    $lines.Add('.NOTES')
    $lines.Add('')
    $lines.Add("$ScriptName v$Version")
    $lines.Add('# ----------------------------------------------------------------------------')
    $lines.Add('# ORIGIN STORY')
    $lines.Add('# ----------------------------------------------------------------------------')
    $lines.Add("#   DATE        : $Date")
    $lines.Add("#   AUTHOR      : $Author")
    $lines.Add("#   DESCRIPTION : v$Version")
    $lines.Add('# ----------------------------------------------------------------------------')
    $lines.Add(' ')
    $lines.Add('.EXAMPLE')
    $lines.Add('')
    $lines.Add("1. $ScriptName.ps1    ")
    $lines.Add('')
    $lines.Add('#> ')
    $lines.Add('')
    $lines.Add("`$nameOfScript       = `"$ScriptName`"")
    $lines.Add("`$scriptVersionValue = `"$Version`"")
    $lines.Add("`$companyName        = `"$CompanyName`"")
    $lines.Add('')
    $lines.Add("`$ErrorActionPreference = `"Continue`"")
    $lines.Add('')
    $lines.Add("`$transcriptPath = `"$ScriptingRoot`"")
    $lines.Add('$transcriptName = $nameOfScript + "_transcript.txt"')
    $lines.Add("`$logExportPath  = `"$ScriptingRoot`"")
    $lines.Add('$LogExportName  = $nameOfScript + "_log.txt"')
    $lines.Add('')
    $lines.Add("`$registryPath = `"$RegistryBase\`" + `$nameOfScript")
    $lines.Add('')
    $lines.Add('$startDay      = "Installation Start"')
    $lines.Add('$StartDayValue = (Get-Date -Format "yyyyMMdd HH:mm:ss")')
    $lines.Add('$EndDay        = "Installation End"')
    $lines.Add('$endDayValue   = ""')
    $lines.Add('$execPSKey     = "ExecutionPowershell"')
    $lines.Add('$execPSValue   = $PSCommandPath')
    $lines.Add('')
    $lines.Add('# ---------------------------------------------------------------------------')
    $lines.Add('# Managed Bookmarks JSON (generated by ManagedBookmarksCreator)')
    $lines.Add('# Single-quoted string — no variable expansion at runtime.')
    $lines.Add('# ---------------------------------------------------------------------------')
    $lines.Add("`$managedBookmarksJson = '$escapedJsonValue'")
    $lines.Add('')
    $lines.Add('$chromePolicyPath = "HKLM:\SOFTWARE\Policies\Google\Chrome"')
    $lines.Add('$edgePolicyPath   = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"')
    $lines.Add('')
    $lines.Add('# ---------------------------------------------------------------------------')
    $lines.Add('# Ensure transcript directory exists')
    $lines.Add('# ---------------------------------------------------------------------------')
    $lines.Add('if (-not (Test-Path $transcriptPath)) {')
    $lines.Add('    Write-Host (Get-Date -Format "yyyyMMdd HH:mm:ss") "- Creating" $transcriptPath -ForegroundColor Green')
    $lines.Add('    New-Item -Path $transcriptPath -ItemType Directory -Force | Out-Null')
    $lines.Add('}')
    $lines.Add('')
    $lines.Add('try { Stop-Transcript -ErrorAction Stop | Out-Null } catch { }')
    $lines.Add('Start-Transcript -Path "$transcriptPath$transcriptName" -Append')
    $lines.Add('')
    $lines.Add('# ---------------------------------------------------------------------------')
    $lines.Add('# Check admin rights')
    $lines.Add('# ---------------------------------------------------------------------------')
    $lines.Add('if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {')
    $lines.Add('    Write-Host "Not running as Administrator. Please restart with admin privileges." -ForegroundColor Red')
    $lines.Add('    try { Stop-Transcript -ErrorAction Stop | Out-Null } catch { }')
    $lines.Add('    exit 1')
    $lines.Add('}')
    $lines.Add('')
    $lines.Add('# Get OS')
    $lines.Add('$os = (Get-CimInstance -ClassName Win32_OperatingSystem).Caption')
    $lines.Add('Write-Host (Get-Date -Format "yyyyMMdd HH:mm:ss") "- OS: $os" -ForegroundColor Green')
    $lines.Add('')
    $lines.Add('# ---------------------------------------------------------------------------')
    $lines.Add('# Registry tracking')
    $lines.Add('# ---------------------------------------------------------------------------')
    $lines.Add('if (-not (Test-Path -Path $registryPath)) { New-Item -Path $registryPath -Force | Out-Null }')
    $lines.Add('New-ItemProperty -Path $registryPath -Name "Name"      -Value $nameOfScript       -PropertyType String -Force | Out-Null')
    $lines.Add('New-ItemProperty -Path $registryPath -Name "Version"   -Value $scriptVersionValue -PropertyType String -Force | Out-Null')
    $lines.Add('New-ItemProperty -Path $registryPath -Name $execPSKey  -Value $execPSValue        -PropertyType String -Force | Out-Null')
    $lines.Add('New-ItemProperty -Path $registryPath -Name $startDay   -Value $StartDayValue      -PropertyType String -Force | Out-Null')
    $lines.Add('')
    $lines.Add('# ---------------------------------------------------------------------------')
    $lines.Add('# Set ManagedBookmarks — Chrome')
    $lines.Add('# Chrome policy key : ManagedBookmarks')
    $lines.Add('# ---------------------------------------------------------------------------')
    $lines.Add('Write-Host (Get-Date -Format "yyyyMMdd HH:mm:ss") "- Setting ManagedBookmarks for Chrome" -ForegroundColor Green')
    $lines.Add('if (-not (Test-Path $chromePolicyPath)) { New-Item -Path $chromePolicyPath -Force | Out-Null }')
    $lines.Add('Set-ItemProperty -Path $chromePolicyPath -Name "ManagedBookmarks"  -Value $managedBookmarksJson -Type String')
    $lines.Add('Set-ItemProperty -Path $chromePolicyPath -Name "BookmarkBarEnabled" -Value 1                   -Type DWord')
    $lines.Add('Write-Host (Get-Date -Format "yyyyMMdd HH:mm:ss") "- Chrome: ManagedBookmarks + BookmarkBarEnabled set." -ForegroundColor Green')
    $lines.Add('')
    $lines.Add('# ---------------------------------------------------------------------------')
    $lines.Add('# Set ManagedFavorites — Edge')
    $lines.Add('# Edge policy key   : ManagedFavorites  (NOT ManagedBookmarks)')
    $lines.Add('# ---------------------------------------------------------------------------')
    $lines.Add('Write-Host (Get-Date -Format "yyyyMMdd HH:mm:ss") "- Setting ManagedFavorites for Edge" -ForegroundColor Green')
    $lines.Add('if (-not (Test-Path $edgePolicyPath)) { New-Item -Path $edgePolicyPath -Force | Out-Null }')
    $lines.Add('Set-ItemProperty -Path $edgePolicyPath -Name "ManagedFavorites"    -Value $managedBookmarksJson -Type String')
    $lines.Add('Set-ItemProperty -Path $edgePolicyPath -Name "FavoritesBarEnabled" -Value 1                    -Type DWord')
    $lines.Add('Write-Host (Get-Date -Format "yyyyMMdd HH:mm:ss") "- Edge: ManagedFavorites + FavoritesBarEnabled set." -ForegroundColor Green')
    $lines.Add('')
    $lines.Add('Write-Host ""')
    $lines.Add('Write-Host "ACTION REQUIRED:" -ForegroundColor Yellow')
    $lines.Add('Write-Host "  Chrome : close ALL Chrome windows (including system tray), then reopen." -ForegroundColor Yellow')
    $lines.Add('Write-Host "  Edge   : close ALL Edge windows (including system tray), then reopen." -ForegroundColor Yellow')
    $lines.Add('Write-Host "  The bookmarks bar label / folder name updates only after a full browser restart." -ForegroundColor Yellow')
    $lines.Add('Write-Host ""')
    $lines.Add('')
    $lines.Add('# ---------------------------------------------------------------------------')
    $lines.Add('# Finish')
    $lines.Add('# ---------------------------------------------------------------------------')
    $lines.Add('$endDayValue = (Get-Date -Format "yyyyMMdd HH:mm:ss")')
    $lines.Add('New-ItemProperty -Path $registryPath -Name $EndDay -Value $endDayValue -PropertyType String -Force | Out-Null')
    $lines.Add('Write-Host (Get-Date -Format "yyyyMMdd HH:mm:ss") "- Script completed." -ForegroundColor Green')
    $lines.Add('')
    $lines.Add('try { Stop-Transcript -ErrorAction Stop | Out-Null } catch { }')

    return $lines -join $nl
}
#endregion

#region ── Helper factories ───────────────────────────────────────────────────
function New-StyledButton {
    param(
        [string]$Text, [int]$X, [int]$Y, [int]$W, [int]$H,
        [System.Drawing.Color]$Back, [System.Drawing.Color]$Fore,
        [System.Drawing.Font]$Font = $null
    )
    # Small UI factory to keep button styling consistent across dialogs.
    $btn           = New-Object System.Windows.Forms.Button
    $btn.Text      = $Text
    $btn.Location  = New-Object System.Drawing.Point($X, $Y)
    $btn.Size      = New-Object System.Drawing.Size($W, $H)
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn.BackColor = $Back
    $btn.ForeColor = $Fore
    $btn.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btn.FlatAppearance.BorderColor = $Back
    $btn.FlatAppearance.BorderSize  = 0
    $btn.Font = if ($Font) { $Font } else { $fontMain }
    return $btn
}

function New-StyledDialog {
    param([string]$Title, [int]$W, [int]$H)
    # Shared modal-dialog shell so every custom prompt inherits the same size,
    # theme, and startup behavior.
    $dlg                 = New-Object System.Windows.Forms.Form
    $dlg.Text            = $Title
    $dlg.Size            = New-Object System.Drawing.Size($W, $H)
    $dlg.StartPosition   = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox     = $false
    $dlg.BackColor       = $clrCard
    $dlg.Font            = $fontMain
    return $dlg
}

function New-StyledTextBox {
    param([int]$X, [int]$Y, [int]$W, [string]$DefaultText = "")
    # Shared textbox factory for consistent colors, borders, and fonts.
    $tb              = New-Object System.Windows.Forms.TextBox
    $tb.Location     = New-Object System.Drawing.Point($X, $Y)
    $tb.Size         = New-Object System.Drawing.Size($W, 26)
    $tb.Text         = $DefaultText
    $tb.BackColor    = $clrInput
    $tb.ForeColor    = $clrText
    $tb.BorderStyle  = [System.Windows.Forms.BorderStyle]::FixedSingle
    $tb.Font         = $fontMain
    return $tb
}

function Set-TextBoxPlaceholder {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.TextBox]$TextBox,
        [string]$Placeholder = ''
    )

    # Newer .NET/PowerShell environments expose PlaceholderText directly. Older
    # WinForms versions need the Win32 cue-banner message instead.
    if ($TextBox.PSObject.Properties.Name -contains 'PlaceholderText') {
        $TextBox.PlaceholderText = $Placeholder
        return
    }

    if (-not ('Win32CueBanner' -as [type])) { return }

    $TextBox.AccessibleDescription = $Placeholder
    $applyCueBanner = {
        param($_sender, $_e)
        if ($_sender.IsHandleCreated) {
            [void][Win32CueBanner]::SendMessage($_sender.Handle, 0x1501, [IntPtr]1, [string]$_sender.AccessibleDescription)
        }
    }

    if ($TextBox.IsHandleCreated) {
        [void][Win32CueBanner]::SendMessage($TextBox.Handle, 0x1501, [IntPtr]1, $Placeholder)
    } else {
        $TextBox.add_HandleCreated($applyCueBanner)
    }
}

function New-DialogLabel {
    param([string]$Text, [int]$X, [int]$Y)
    # Lightweight label factory used by the custom dialogs.
    $lbl           = New-Object System.Windows.Forms.Label
    $lbl.Text      = $Text
    $lbl.Location  = New-Object System.Drawing.Point($X, $Y)
    $lbl.AutoSize  = $true
    $lbl.ForeColor = $clrMuted
    $lbl.Font      = $fontMain
    return $lbl
}
#endregion

#region ── Status helper ─────────────────────────────────────────────────────
function Set-StatusLabel {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Label]$Label,
        [string]$Text  = '',
        [ValidateSet('ok','fail','warn','info')][string]$State = 'info'
    )
    if ($script:outputModuleLoaded -and $Text -ne '') {
        Set-DecosterOutputLabelStatus -Label $Label -Text $Text -State $State
        if ($State -eq 'info') { $Label.ForeColor = $clrMuted }
    } else {
        $Label.Text      = $Text
        $Label.ForeColor = switch ($State) {
            'ok'    { $clrSuccess }
            'fail'  { $clrError   }
            'warn'  { $clrWarn    }
            default { $clrMuted   }
        }
    }
}

function Update-Preview { $txtJson.Text = Build-Json }

# -----------------------------------------------------------------------------
# Autosave + Session-Recovery helpers
# -----------------------------------------------------------------------------
# These functions intentionally sit close to UI/status helpers because they are
# UI-aware (status messages, prompts) and editor-state aware (Build/Restore).

function Initialize-RecoveryStorage {
    # Choose a stable per-user path that works with both .ps1 and packaged .exe.
    $baseLocalAppData = [Environment]::GetFolderPath('LocalApplicationData')
    $script:recoveryRootPath      = Join-Path $baseLocalAppData 'Decoster.tech\ManagedBookmarksCreator\Recovery'
    $script:recoveryStateFilePath = Join-Path $script:recoveryRootPath 'latest.state.json'
    $script:recoveryLockFilePath  = Join-Path $script:recoveryRootPath 'session.lock'

    try {
        if (-not (Test-Path -LiteralPath $script:recoveryRootPath)) {
            New-Item -ItemType Directory -Path $script:recoveryRootPath -Force -ErrorAction Stop | Out-Null
        }
        $script:recoveryActive = $true
        Write-AppHostMessage "Recovery storage initialized: $($script:recoveryRootPath)" 'Cyan'
    }
    catch {
        Mark-ErrorRecordSeen $_
        # Recovery is optional; the editor must keep working even if path setup fails.
        $script:recoveryActive = $false
        Write-AppHostMessage "Recovery disabled (path init failed): $($_.Exception.Message)" 'Yellow'
        Add-LogEntry 'Recovery storage initialization failed' 'warn' -ErrorMessage $_.Exception.Message
    }
}

function Save-RecoverySnapshot {
    param([string]$Reason = 'autosave')

    if (-not $script:recoveryActive) { return $false }

    try {
        # Wrap state with metadata so future schema changes stay manageable.
        $payload = [PSCustomObject]@{
            schema_version = 1
            saved_at       = (Get-Date).ToString('o')
            reason         = $Reason
            editor_state   = (Get-EditorStateJson)
        }

        $json = $payload | ConvertTo-Json -Compress -Depth 10
        Set-Content -LiteralPath $script:recoveryStateFilePath -Value $json -Encoding UTF8 -ErrorAction Stop
        $script:autosaveIsDirty = $false
        return $true
    }
    catch {
        Mark-ErrorRecordSeen $_
        Write-AppHostMessage "Recovery save failed: $($_.Exception.Message)" 'Yellow'
        Add-LogEntry 'Recovery snapshot save failed' 'warn' -ErrorMessage $_.Exception.Message
        return $false
    }
}

function Try-RestoreRecoverySnapshot {
    if (-not $script:recoveryActive) { return $false }
    if (-not (Test-Path -LiteralPath $script:recoveryStateFilePath)) { return $false }

    # Recovery is offered when a previous lock file still exists. That indicates
    # the prior session likely ended unexpectedly.
    if (-not (Test-Path -LiteralPath $script:recoveryLockFilePath)) { return $false }

    $choice = [System.Windows.Forms.MessageBox]::Show(
        "An unfinished previous session was detected. Do you want to restore the last autosaved state?",
        "Session recovery",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) { return $false }

    try {
        $payload = Get-Content -LiteralPath $script:recoveryStateFilePath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($null -eq $payload.editor_state) {
            throw 'Recovery payload does not contain editor_state.'
        }
        Restore-EditorStateJson $payload.editor_state
        $script:autosaveIsDirty = $false
        Add-LogEntry 'Recovered previous session from autosave.' 'ok'
        Set-StatusLabel -Label $actionLabel -Text 'Recovered previous session from autosave.' -State 'ok'
        return $true
    }
    catch {
        Mark-ErrorRecordSeen $_
        Add-LogEntry 'Recovery load failed' 'fail' -ErrorMessage $_.Exception.Message
        $friendly = ConvertTo-FriendlyErrorMessage $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Could not restore autosave:`n$friendly", "Recovery error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return $false
    }
}

function Start-RecoverySession {
    if (-not $script:recoveryActive) { return }

    # Touch the lock file. If this process exits cleanly, we remove it again.
    # If not, next startup sees the lock and offers recovery.
    $lockInfo = "started=$(Get-Date -Format o); pid=$PID"
    Set-Content -LiteralPath $script:recoveryLockFilePath -Value $lockInfo -Encoding UTF8 -ErrorAction SilentlyContinue

    $script:autosaveTimer = New-Object System.Windows.Forms.Timer
    $script:autosaveTimer.Interval = $script:autosaveIntervalMs
    $script:autosaveTimer.Add_Tick({
        if ($script:autosaveIsDirty -and -not $script:isRestoringEditorState) {
            Save-RecoverySnapshot -Reason 'autosave-timer' | Out-Null
        }
    })
    $script:autosaveTimer.Start()
}

function Stop-RecoverySession {
    # Final autosave on clean shutdown, then remove lock marker.
    if ($script:recoveryActive) {
        Save-RecoverySnapshot -Reason 'clean-exit' | Out-Null
        if (Test-Path -LiteralPath $script:recoveryLockFilePath) {
            Remove-Item -LiteralPath $script:recoveryLockFilePath -Force -ErrorAction SilentlyContinue
        }
    }

    if ($null -ne $script:autosaveTimer) {
        $script:autosaveTimer.Stop()
        $script:autosaveTimer.Dispose()
        $script:autosaveTimer = $null
    }
}
#endregion

#region ── Dialogs ───────────────────────────────────────────────────────────
function Show-FolderDialog($title, $defaultName = "") {
    # Reused for both create and edit folder actions. Returns trimmed text or null
    # when the user cancels the dialog.
    $dlg = New-StyledDialog $title 360 150
    $dlg.Controls.AddRange(@(
        (New-DialogLabel "Folder name:" 10 14),
        ($txt = New-StyledTextBox 10 34 320 $defaultName),
        ($ok  = New-StyledButton "OK"     152 72 80 28 $clrAccent ([System.Drawing.Color]::White) $fontBold),
        ($can = New-StyledButton "Cancel" 240 72 90 28 $clrBorder $clrText)
    ))
    $ok.DialogResult  = [System.Windows.Forms.DialogResult]::OK
    $can.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.AcceptButton = $ok
    $dlg.CancelButton = $can

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $txt.Text.Trim() -ne "") {
        return $txt.Text.Trim()
    }
    return $null
}

function Show-SaveNameDialog($savePath) {
    # Custom lightweight save dialog used when JsonBasePath is configured and we
    # only need a file name, not a full folder browser.
    $dlg = New-StyledDialog "File Name" 400 160
    $dlg.Controls.AddRange(@(
        (New-DialogLabel "Save in: $savePath" 10 12),
        (New-DialogLabel "File name:" 10 42),
        ($txt = New-StyledTextBox 10 62 360 "managed_bookmarks"),
        ($ok  = New-StyledButton "Save"   194 100 90 28 $clrAccent ([System.Drawing.Color]::White) $fontBold),
        ($can = New-StyledButton "Cancel" 292 100 90 28 $clrBorder $clrText)
    ))
    $ok.DialogResult  = [System.Windows.Forms.DialogResult]::OK
    $can.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.AcceptButton = $ok
    $dlg.CancelButton = $can

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $txt.Text.Trim() -ne "") {
        $name = $txt.Text.Trim()
        if (-not $name.EndsWith(".json")) { $name = "$name.json" }
        return $name
    }
    return $null
}

function Show-SaveFileDialog {
    param(
        [string]$Title,
        [string]$Filter,
        [string]$DefaultFileName,
        [string]$InitialDirectory = ""
    )

    # Fallback to the standard Windows dialog when no managed default path is available.
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title = $Title
    $dlg.Filter = $Filter
    $dlg.FileName = $DefaultFileName
    if ((Test-HasText $InitialDirectory) -and (Test-Path -LiteralPath $InitialDirectory)) {
        $dlg.InitialDirectory = $InitialDirectory
    }

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and (Test-HasText $dlg.FileName)) {
        return $dlg.FileName
    }

    return $null
}

function Show-ExportScriptDialog {
    # Gather only the inputs required for script generation and sanitize the
    # script base name before any file write happens.
    $dlg = New-StyledDialog "Export Deployment Script" 480 230
    $dlg.Controls.AddRange(@(
        (New-DialogLabel "Script name (no spaces, no extension):" 10 12),
        ($txtName = New-StyledTextBox 10 32 440 "Set-ManagedBookmarks"),
        (New-DialogLabel "Description (for the script header):" 10 72),
        ($txtDesc = New-StyledTextBox 10 92 440 "Sets the ManagedBookmarks policy for Chrome and Edge."),
        ($ok  = New-StyledButton "Export" 264 160 100 28 $clrAccent ([System.Drawing.Color]::White) $fontBold),
        ($can = New-StyledButton "Cancel" 372 160  90 28 $clrBorder $clrText)
    ))
    $ok.DialogResult  = [System.Windows.Forms.DialogResult]::OK
    $can.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.AcceptButton = $ok
    $dlg.CancelButton = $can

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $txtName.Text.Trim() -ne "") {
        $safeName = ConvertTo-SafeFileBaseName $txtName.Text
        if (-not (Test-HasText $safeName)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Invalid script name. Use letters, numbers, underscore (_), dash (-), or dot (.).",
                "Invalid name",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return $null
        }
        return @{ name = $safeName; description = $txtDesc.Text.Trim() }
    }
    return $null
}

function Show-LinkDialog($title, $defaultName = "", $defaultUrl = "") {
    # Reused for add/edit bookmark actions. Validation happens before returning
    # so callers can assume they receive a usable absolute URL.
    $dlg = New-StyledDialog $title 420 210
    $dlg.Controls.AddRange(@(
        (New-DialogLabel "Name:" 10 14),
        ($txtN = New-StyledTextBox 10 34 382 $defaultName),
        (New-DialogLabel "URL:" 10 70),
        ($txtU = New-StyledTextBox 10 90 382 $defaultUrl),
        ($ok   = New-StyledButton "OK"     220 138 80 28 $clrAccent ([System.Drawing.Color]::White) $fontBold),
        ($can  = New-StyledButton "Cancel" 308 138 90 28 $clrBorder $clrText)
    ))
    $ok.DialogResult  = [System.Windows.Forms.DialogResult]::OK
    $can.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.AcceptButton = $ok
    $dlg.CancelButton = $can

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK `
        -and $txtN.Text.Trim() -ne "" -and $txtU.Text.Trim() -ne "") {
        $urlText = $txtU.Text.Trim()
        if (-not (Test-IsValidHttpUrl $urlText)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Invalid URL. Only absolute http/https URLs are allowed.",
                "Invalid URL",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return $null
        }
        return @{ name = $txtN.Text.Trim(); url = $urlText }
    }
    return $null
}
#endregion

#region ── Form ──────────────────────────────────────────────────────────────
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "$adminPrefix$GLB_ScriptTitel  |  v$GLB_scriptVersion  |  $GLB_ScriptUpdateDate"
$form.Size            = New-Object System.Drawing.Size(1100, 860)
$form.StartPosition   = "CenterScreen"
$form.MinimumSize     = New-Object System.Drawing.Size(900, 700)
$form.BackColor       = $clrBg
$form.Font            = $fontMain
#endregion

#region ── Toolbar ───────────────────────────────────────────────────────────
$toolStrip            = New-Object System.Windows.Forms.ToolStrip
$toolStrip.Dock       = "Top"
$toolStrip.BackColor  = $clrCard
$toolStrip.ForeColor  = $clrText
$toolStrip.RenderMode = [System.Windows.Forms.ToolStripRenderMode]::System
$toolStrip.GripStyle  = [System.Windows.Forms.ToolStripGripStyle]::Hidden

function New-TsButton($text, $tip) {
    $b              = New-Object System.Windows.Forms.ToolStripButton
    $b.Text         = $text
    $b.ToolTipText  = $tip
    $b.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::Text
    $b.ForeColor    = $clrText
    $b.BackColor    = $clrCard
    return $b
}

$btnAddFolder    = New-TsButton "📁 Folder"     "Add root folder"
$btnAddSubFolder = New-TsButton "📁+ Subfolder" "Add subfolder to selected folder"
$btnAddLink      = New-TsButton "🔗 Link"       "Add link to selected folder"
$btnAddRootLink  = New-TsButton "🔗 Root link"  "Add link at root level"
$sep1            = New-Object System.Windows.Forms.ToolStripSeparator
$btnEdit         = New-TsButton "✏️ Edit"        "Edit selected item"
$btnDelete       = New-TsButton "🗑️ Delete"      "Delete selected item"
$sep2            = New-Object System.Windows.Forms.ToolStripSeparator
$btnMoveUp       = New-TsButton "▲ Up"          "Move item up"
$btnMoveDown     = New-TsButton "▼ Down"        "Move item down"
$btnUndo         = New-TsButton "↶ Undo"        "Undo last change"
$btnRedo         = New-TsButton "↷ Redo"        "Redo last undone change"
$sep3            = New-Object System.Windows.Forms.ToolStripSeparator
$btnLoad         = New-TsButton "📂 Load"          "Load existing JSON file"
$btnSave         = New-TsButton "💾 Save"          "Save JSON"
$btnCopy         = New-TsButton "📋 Copy"          "Copy JSON to clipboard"
$sep4            = New-Object System.Windows.Forms.ToolStripSeparator
$btnExportScript = New-TsButton "📜 Export Script" "Generate a registry deployment script (.ps1)"
$btnExportScript.Available = $false   # hidden until config with ScriptBasePath is confirmed

$sep5               = New-Object System.Windows.Forms.ToolStripSeparator
$btnWriteRegistry   = New-TsButton "🖊️ Write to Registry" "Write bookmarks directly to HKLM registry (Chrome + Edge)"
$btnWriteRegistry.ForeColor = $clrWarn
$btnWriteRegistry.Available = $isAdmin
$sep6               = New-Object System.Windows.Forms.ToolStripSeparator
$btnValidateContract = New-TsButton "✅ Validate" "Run JSON import/export roundtrip contract test"

$toolStrip.Items.AddRange(@(
    $btnAddFolder, $btnAddSubFolder, $btnAddLink, $btnAddRootLink,
    $sep1, $btnEdit, $btnDelete, $sep2, $btnMoveUp, $btnMoveDown, $btnUndo, $btnRedo,
    $sep3, $btnLoad, $btnSave, $btnCopy,
    $sep4, $btnExportScript,
    $sep5, $btnWriteRegistry,
    $sep6, $btnValidateContract
))
#endregion

#region ── Content panel (tree + tabs) ───────────────────────────────────────
$contentPanel           = New-Object System.Windows.Forms.Panel
$contentPanel.BackColor = $clrCard

$split                     = New-Object System.Windows.Forms.SplitContainer
$split.Dock                = "Fill"
$split.Orientation         = [System.Windows.Forms.Orientation]::Horizontal
$split.SplitterDistance    = 320
$split.BackColor           = $clrBorder
$split.Panel1.BackColor    = $clrInput
$split.Panel2.BackColor    = $clrCard

# ── Bar label panel (above the tree)
$pnlBarLabel              = New-Object System.Windows.Forms.Panel
$pnlBarLabel.Dock         = [System.Windows.Forms.DockStyle]::Top
$pnlBarLabel.Height       = 36
$pnlBarLabel.BackColor    = $clrCard
$pnlBarLabel.Padding      = New-Object System.Windows.Forms.Padding(0, 1, 0, 1)

$lblBarNameStatic           = New-Object System.Windows.Forms.Label
$lblBarNameStatic.Text      = "Bookmarks bar label:"
$lblBarNameStatic.Dock      = [System.Windows.Forms.DockStyle]::Left
$lblBarNameStatic.Width     = 145
$lblBarNameStatic.ForeColor = $clrText
$lblBarNameStatic.Font      = $fontMain
$lblBarNameStatic.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblBarNameStatic.Padding   = New-Object System.Windows.Forms.Padding(8, 0, 4, 0)

$script:txtTopLevelName             = New-Object System.Windows.Forms.TextBox
$script:txtTopLevelName.Dock        = [System.Windows.Forms.DockStyle]::Fill
$script:txtTopLevelName.BackColor   = $clrInput
$script:txtTopLevelName.ForeColor   = $clrText
$script:txtTopLevelName.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$script:txtTopLevelName.Font        = $fontMain
Set-TextBoxPlaceholder -TextBox $script:txtTopLevelName -Placeholder "e.g. Managed Bookmarks  (required for Edge)"

$txtWrapper           = New-Object System.Windows.Forms.Panel
$txtWrapper.Dock      = [System.Windows.Forms.DockStyle]::Fill
$txtWrapper.BackColor = $clrInput
$txtWrapper.Padding   = New-Object System.Windows.Forms.Padding(0, 5, 0, 0)
$txtWrapper.Controls.Add($script:txtTopLevelName)

# Dock order: Fill control added first (index 0), Left/Top controls added after (index 1+)
# WinForms processes docking from last index to first, so the Top/Left controls claim
# their edge first and the Fill control gets whatever remains.
$pnlBarLabel.Controls.Add($txtWrapper)       # index 0 — Fill (claims last)
$pnlBarLabel.Controls.Add($lblBarNameStatic) # index 1 — Left (claims first)

$script:txtTopLevelName.add_TextChanged({
    Update-Preview
    if (-not $script:isRestoringEditorState) {
        $script:autosaveIsDirty = $true
    }
})

# ── TreeView
$script:tree               = New-Object System.Windows.Forms.TreeView
$script:tree.Dock          = "Fill"
$script:tree.HideSelection = $false
$script:tree.BackColor     = $clrInput
$script:tree.ForeColor     = $clrText
$script:tree.BorderStyle   = [System.Windows.Forms.BorderStyle]::None
$script:tree.DrawMode      = [System.Windows.Forms.TreeViewDrawMode]::OwnerDrawText
$script:tree.ShowLines     = $true
$script:tree.ShowRootLines = $true
$script:tree.ShowPlusMinus = $true
$script:tree.Font          = $fontMain
$script:tree.AllowDrop     = $true
$script:tree.add_DrawNode({ param($s, $e); $e.DrawDefault = $true })

# Tree (Fill) added first (index 0), bar panel (Top) added second (index 1)
# so the bar panel claims the top 36 px and the tree fills the rest.
$split.Panel1.Controls.Add($script:tree)    # index 0 — Fill
$split.Panel1.Controls.Add($pnlBarLabel)    # index 1 — Top

# ── Tab control
$tabControl           = New-Object System.Windows.Forms.TabControl
$tabControl.Dock      = "Fill"
$tabControl.BackColor = $clrCard

# Tab 1 — Preview
$tabPreview           = New-Object System.Windows.Forms.TabPage
$tabPreview.Text      = "JSON Preview"
$tabPreview.BackColor = $clrCard

$txtJson              = New-Object System.Windows.Forms.TextBox
$txtJson.Multiline    = $true
$txtJson.ScrollBars   = "Both"
$txtJson.WordWrap     = $false
$txtJson.Font         = New-Object System.Drawing.Font("Consolas", 9)
$txtJson.Dock         = "Fill"
$txtJson.ReadOnly     = $true
$txtJson.BackColor    = $clrInput
$txtJson.ForeColor    = $clrText
$txtJson.BorderStyle  = [System.Windows.Forms.BorderStyle]::None
$tabPreview.Controls.Add($txtJson)

# Tab 2 — Import
$tabImport            = New-Object System.Windows.Forms.TabPage
$tabImport.Text       = "JSON Import"
$pnlImport            = New-Object System.Windows.Forms.Panel
$pnlImport.Dock       = "Fill"

$txtImport              = New-Object System.Windows.Forms.TextBox
$txtImport.Multiline    = $true
$txtImport.ScrollBars   = "Both"
$txtImport.WordWrap     = $false
$txtImport.Font         = New-Object System.Drawing.Font("Consolas", 9)
$txtImport.Dock         = "Fill"
$txtImport.BackColor    = $clrInput
$txtImport.ForeColor    = $clrText
$txtImport.BorderStyle  = [System.Windows.Forms.BorderStyle]::None
Set-TextBoxPlaceholder -TextBox $txtImport -Placeholder 'Paste a ManagedBookmarks/Favorites JSON string here...'

$pnlImportBtn           = New-Object System.Windows.Forms.Panel
$pnlImportBtn.Dock      = "Bottom"
$pnlImportBtn.Height    = 40
$pnlImportBtn.BackColor = $clrCard

$btnImportStr   = New-StyledButton "Import JSON" 6  6 120 28 $clrAccent ([System.Drawing.Color]::White) $fontBold
$btnClearImport = New-StyledButton "Clear"       134 6  60 28 $clrBorder $clrText

$pnlImportBtn.Controls.AddRange(@($btnImportStr, $btnClearImport))
$pnlImport.Controls.AddRange(@($txtImport, $pnlImportBtn))
$tabImport.Controls.Add($pnlImport)

# Tab 3 — Registry (admin only)
if ($isAdmin) {
    $tabRegistry            = New-Object System.Windows.Forms.TabPage
    $tabRegistry.Text       = "Registry"
    $tabRegistry.BackColor  = $clrCard

    $pnlRegistry            = New-Object System.Windows.Forms.Panel
    $pnlRegistry.Dock       = "Fill"

    # Top toolbar panel
    $pnlRegToolbar          = New-Object System.Windows.Forms.Panel
    $pnlRegToolbar.Dock     = "Top"
    $pnlRegToolbar.Height   = 44
    $pnlRegToolbar.BackColor = $clrCard

    $lblRegBrowser          = New-Object System.Windows.Forms.Label
    $lblRegBrowser.Text     = "Browser:"
    $lblRegBrowser.ForeColor = $clrText
    $lblRegBrowser.Font     = $fontMain
    $lblRegBrowser.Location = New-Object System.Drawing.Point(6, 13)
    $lblRegBrowser.AutoSize = $true

    $cmbRegBrowser          = New-Object System.Windows.Forms.ComboBox
    $cmbRegBrowser.Location = New-Object System.Drawing.Point(60, 9)
    $cmbRegBrowser.Size     = New-Object System.Drawing.Size(160, 26)
    $cmbRegBrowser.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbRegBrowser.BackColor = $clrInput
    $cmbRegBrowser.ForeColor = $clrText
    $cmbRegBrowser.Font     = $fontMain
    $cmbRegBrowser.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    [void]$cmbRegBrowser.Items.AddRange(@(
        'Chrome  (HKLM\...\Google\Chrome)',
        'Edge    (HKLM\...\Microsoft\Edge)'
    ))
    $cmbRegBrowser.SelectedIndex = 0

    $btnReadReg  = New-StyledButton "Read"         230 8 70 28 $clrAccent ([System.Drawing.Color]::White) $fontBold
    $btnLoadReg  = New-StyledButton "Load in editor" 308 8 110 28 $clrSuccess ([System.Drawing.Color]::Black) $fontBold

    $pnlRegToolbar.Controls.AddRange(@($lblRegBrowser, $cmbRegBrowser, $btnReadReg, $btnLoadReg))

    # Value display
    $txtReg              = New-Object System.Windows.Forms.TextBox
    $txtReg.Multiline    = $true
    $txtReg.ScrollBars   = "Both"
    $txtReg.WordWrap     = $false
    $txtReg.Font         = New-Object System.Drawing.Font("Consolas", 9)
    $txtReg.Dock         = "Fill"
    $txtReg.ReadOnly     = $true
    $txtReg.BackColor    = $clrInput
    $txtReg.ForeColor    = $clrText
    $txtReg.BorderStyle  = [System.Windows.Forms.BorderStyle]::None

    $pnlRegistry.Controls.AddRange(@($txtReg, $pnlRegToolbar))
    $tabRegistry.Controls.Add($pnlRegistry)
}

$tabPages = @($tabPreview, $tabImport)
if ($isAdmin) { $tabPages += $tabRegistry }
$tabControl.TabPages.AddRange($tabPages)
$split.Panel2.Controls.Add($tabControl)
$contentPanel.Controls.Add($split)
#endregion

#region ── Status section ────────────────────────────────────────────────────
$separator1           = New-Object System.Windows.Forms.Label
$separator1.BackColor = $clrBorder
$separator1.Height    = 1

$actionLabel            = New-Object System.Windows.Forms.Label
$actionLabel.Font       = $fontBold
$actionLabel.ForeColor  = $clrMuted
$actionLabel.Height     = 22
#endregion

#region ── Signature ─────────────────────────────────────────────────────────
$separator2           = New-Object System.Windows.Forms.Label
$separator2.BackColor = $clrBorder
$separator2.Height    = 1

$sigText = @"
╔════════════════════════════════════════════╗
║  ██ ██  WRITTEN DESIGNED BY DECOSTER.TECH  ║
║  █████                                      ║
╚══██ ██══════════════════════════════════════╝
    ██ ██ans            scripting@decoster.tech
"@
$sigLabel           = New-Object System.Windows.Forms.Label
$sigLabel.Text      = $sigText
$sigLabel.Font      = $fontMono
$sigLabel.ForeColor = $clrSig
$sigLabel.BackColor = $clrBg
$sigLabel.Height    = 108
#endregion

$form.Controls.AddRange(@($toolStrip, $contentPanel, $separator1, $actionLabel, $separator2, $sigLabel))

#region ── Layout ────────────────────────────────────────────────────────────
function Invoke-Layout {
    # Centralized manual layout so the form keeps predictable spacing when the
    # window is resized. Positions are computed from the bottom up because the
    # signature and status area reserve fixed vertical space.
    $w      = $form.ClientSize.Width
    $h      = $form.ClientSize.Height
    $margin = 16
    $toolH  = $toolStrip.Height

    $sigLabel.Location   = New-Object System.Drawing.Point($margin, ($h - 108))
    $sigLabel.Width      = $w - $margin * 2

    $separator2.Location = New-Object System.Drawing.Point($margin, ($h - 108 - 1 - 6))
    $separator2.Width    = $w - $margin * 2

    $actionLabel.Location = New-Object System.Drawing.Point($margin, ($h - 108 - 1 - 22 - 10))
    $actionLabel.Width    = $w - $margin * 2

    $separator1.Location = New-Object System.Drawing.Point($margin, ($h - 150))
    $separator1.Width    = $w - $margin * 2

    $top    = $toolH + $margin
    $bottom = $separator1.Top - $margin
    $contentPanel.Location = New-Object System.Drawing.Point($margin, $top)
    $contentPanel.Size     = New-Object System.Drawing.Size(($w - $margin * 2), ($bottom - $top))
}

$form.add_Load({
    Invoke-Layout

    # Initialize recovery and, if needed, offer restoring the previous session.
    Initialize-RecoveryStorage
    Start-ErrorMonitoring
    Try-RestoreRecoverySnapshot | Out-Null
    Start-RecoverySession

    if ($script:configWarnings.Count -gt 0) {
        Set-StatusLabel $actionLabel "Config incomplete — missing keys: $($script:configWarnings -join ', ')" 'warn'
    }
    $btnExportScript.Available = $true

})
$form.add_Resize({ Invoke-Layout })
$form.add_FormClosing({
    # Ensure lock file cleanup and a last snapshot on normal app shutdown.
    Stop-RecoverySession
    Sync-GlobalErrorLog
    Stop-ErrorMonitoring
})
#endregion

#region ── Button events ─────────────────────────────────────────────────────

# Add root folder
$btnAddFolder.add_Click({
    $name = Show-FolderDialog "New root folder"
    if ($null -eq $name) { return }
    Push-UndoState
    $node = New-FolderNode $name
    $script:bookmarks.Add($node)
    Add-TreeNode $null $node | Out-Null
    $script:tree.ExpandAll()
    Update-Preview
    Add-LogEntry "Folder added" 'ok'
    Add-LogKeyValue "Name" $name
    Set-StatusLabel $actionLabel "Folder '$name' added." 'ok'
})

# Add subfolder
$btnAddSubFolder.add_Click({
    $sel = $script:tree.SelectedNode
    if ($null -eq $sel -or $sel.Tag.type -ne "folder") {
        Set-StatusLabel $actionLabel "Select a folder first." 'warn'
        return
    }
    $name = Show-FolderDialog "New subfolder"
    if ($null -eq $name) { return }
    Push-UndoState
    $node = New-FolderNode $name
    $sel.Tag.children.Add($node)
    Add-TreeNode $sel $node | Out-Null
    $sel.Expand()
    Update-Preview
    Add-LogEntry "Subfolder added" 'ok'
    Add-LogKeyValue "Name"      $name
    Add-LogKeyValue "In folder" $sel.Tag.name
    Set-StatusLabel $actionLabel "Subfolder '$name' added." 'ok'
})

# Add link in selected folder
$btnAddLink.add_Click({
    $sel = $script:tree.SelectedNode
    if ($null -eq $sel -or $sel.Tag.type -ne "folder") {
        Set-StatusLabel $actionLabel "Select a folder first." 'warn'
        return
    }
    $result = Show-LinkDialog "New link"
    if ($null -eq $result) { return }
    Push-UndoState
    $node = New-UrlNode $result.name $result.url
    $sel.Tag.children.Add($node)
    Add-TreeNode $sel $node | Out-Null
    $sel.Expand()
    Update-Preview
    Add-LogEntry "Link added" 'ok'
    Add-LogKeyValue "Name"      $result.name
    Add-LogKeyValue "URL"       $result.url
    Add-LogKeyValue "In folder" $sel.Tag.name
    Set-StatusLabel $actionLabel "Link '$($result.name)' added." 'ok'
})

# Add link at root level
$btnAddRootLink.add_Click({
    $result = Show-LinkDialog "New root link"
    if ($null -eq $result) { return }
    Push-UndoState
    $node = New-UrlNode $result.name $result.url
    $script:bookmarks.Add($node)
    Add-TreeNode $null $node | Out-Null
    Update-Preview
    Add-LogEntry "Root link added" 'ok'
    Add-LogKeyValue "Name" $result.name
    Add-LogKeyValue "URL"  $result.url
    Set-StatusLabel $actionLabel "Root link '$($result.name)' added." 'ok'
})

# Edit selected item
$btnEdit.add_Click({
    $sel = $script:tree.SelectedNode
    if ($null -eq $sel) {
        Set-StatusLabel $actionLabel "Select an item to edit." 'warn'
        return
    }
    $d = $sel.Tag
    Push-UndoState
    if ($d.type -eq "folder") {
        $name = Show-FolderDialog "Edit folder" $d.name
        if ($null -eq $name) { return }
        $d.name = $name
    } else {
        $result = Show-LinkDialog "Edit link" $d.name $d.url
        if ($null -eq $result) { return }
        $d.name = $result.name
        $d.url  = $result.url
    }
    Update-TreeNodeText $sel
    Update-Preview
    Add-LogEntry "Item updated" 'ok'
    Add-LogKeyValue "Name" $d.name
    Add-LogKeyValue "Type" $d.type
    Set-StatusLabel $actionLabel "'$($d.name)' updated." 'ok'
})

# Delete selected item
$btnDelete.add_Click({
    $sel = $script:tree.SelectedNode
    if ($null -eq $sel) { return }
    $ans = [System.Windows.Forms.MessageBox]::Show(
        "Delete '$($sel.Tag.name)'?", "Confirm",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    Push-UndoState

    $outList = $null; $outParent = $null
    Find-ParentList $sel.Tag ([ref]$outList) ([ref]$outParent) | Out-Null
    $deletedName = $sel.Tag.name
    $outList.Remove($sel.Tag) | Out-Null
    $sel.Remove()
    Update-Preview
    Add-LogEntry "Item deleted" 'warn'
    Add-LogKeyValue "Name" $deletedName
    Set-StatusLabel $actionLabel "'$deletedName' deleted." 'warn'
})

# Move up
$btnMoveUp.add_Click({
    $sel = $script:tree.SelectedNode
    if ($null -eq $sel) { return }
    $outList = $null; $outParent = $null
    Find-ParentList $sel.Tag ([ref]$outList) ([ref]$outParent) | Out-Null
    $idx = $outList.IndexOf($sel.Tag)
    if ($idx -gt 0) {
        $prevSibling = $outList[$idx - 1]
        if ($prevSibling.type -eq "folder") {
            # Previous sibling is a folder — enter it as last child so the item
            # moves through every visual position instead of skipping the folder.
            Push-UndoState
            $outList.RemoveAt($idx)
            $prevSibling.children.Add($sel.Tag) | Out-Null
        } else {
            # Normal swap with the item directly above.
            Push-UndoState
            $outList.RemoveAt($idx)
            $outList.Insert($idx - 1, $sel.Tag)
        }
    } elseif ($null -ne $outParent) {
        # At top of folder — move item out, just before the parent folder.
        $grandList = $null; $grandParent = $null
        Find-ParentList $outParent ([ref]$grandList) ([ref]$grandParent) | Out-Null
        $folderIdx = $grandList.IndexOf($outParent)
        Push-UndoState
        $outList.RemoveAt($idx)
        $grandList.Insert($folderIdx, $sel.Tag)
    } else {
        return  # already at very top of root list
    }
    $data = $sel.Tag
    Update-Tree
    foreach ($tv in $script:tree.Nodes) { Select-ByData $tv $data }
    Update-Preview
})

# Move down
$btnMoveDown.add_Click({
    $sel = $script:tree.SelectedNode
    if ($null -eq $sel) { return }
    $outList = $null; $outParent = $null
    Find-ParentList $sel.Tag ([ref]$outList) ([ref]$outParent) | Out-Null
    $idx = $outList.IndexOf($sel.Tag)
    if ($idx -lt $outList.Count - 1) {
        $nextSibling = $outList[$idx + 1]
        if ($nextSibling.type -eq "folder") {
            # Next sibling is a folder — enter it as first child so the item
            # moves through every visual position instead of skipping the folder.
            Push-UndoState
            $outList.RemoveAt($idx)
            $nextSibling.children.Insert(0, $sel.Tag)
        } else {
            # Normal swap with the item directly below.
            Push-UndoState
            $outList.RemoveAt($idx)
            $outList.Insert($idx + 1, $sel.Tag)
        }
    } elseif ($null -ne $outParent) {
        # At bottom of folder — move item out, just after the parent folder.
        $grandList = $null; $grandParent = $null
        Find-ParentList $outParent ([ref]$grandList) ([ref]$grandParent) | Out-Null
        $folderIdx = $grandList.IndexOf($outParent)
        Push-UndoState
        $outList.RemoveAt($idx)
        $grandList.Insert($folderIdx + 1, $sel.Tag)
    } else {
        return  # already at very bottom of root list
    }
    $data = $sel.Tag
    Update-Tree
    foreach ($tv in $script:tree.Nodes) { Select-ByData $tv $data }
    Update-Preview
})

# ── Drag & Drop ──────────────────────────────────────────────────────────────

# Start drag when the user drags a tree node with the left mouse button.
$script:tree.add_ItemDrag({
    param($s, $e)
    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
    $script:dragNode = $e.Item
    [void]$script:tree.DoDragDrop($e.Item, [System.Windows.Forms.DragDropEffects]::Move)
})

# Accept only TreeNode data dragged from this tree itself.
$script:tree.add_DragEnter({
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.TreeNode])) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Move
    } else {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
})

# Highlight the target node and track insert-before/after as the cursor moves.
$script:tree.add_DragOver({
    param($s, $e)
    $e.Effect = [System.Windows.Forms.DragDropEffects]::Move
    $pt         = $script:tree.PointToClient([System.Drawing.Point]::new($e.X, $e.Y))
    $targetNode = $script:tree.GetNodeAt($pt)

    # No valid target: clear any existing highlight.
    if ($null -eq $targetNode -or $targetNode -eq $script:dragNode) {
        Clear-DropHighlight
        return
    }

    # Do not allow dropping onto a descendant of the dragged node.
    $check = $targetNode.Parent
    while ($null -ne $check) {
        if ($check -eq $script:dragNode) { Clear-DropHighlight; return }
        $check = $check.Parent
    }

    # Determine insert position: top half → before, bottom half → after.
    $script:dropBefore = ($pt.Y - $targetNode.Bounds.Top) -lt ($targetNode.Bounds.Height / 2)

    # Update highlight only when the target node changes.
    if ($script:dropTarget -ne $targetNode) {
        Clear-DropHighlight
        $script:dropTarget           = $targetNode
        $script:dropTarget.BackColor = $clrAccent
        $script:dropTarget.ForeColor = [System.Drawing.Color]::White
    }

    # Auto-scroll when the cursor is near the top or bottom edge.
    if ($pt.Y -lt 20 -and $null -ne $script:tree.TopNode) {
        $prev = $script:tree.TopNode.PrevVisibleNode
        if ($null -ne $prev) { $script:tree.TopNode = $prev }
    } elseif ($pt.Y -gt ($script:tree.Height - 20) -and $null -ne $script:tree.TopNode) {
        $next = $script:tree.TopNode.NextVisibleNode
        if ($null -ne $next) { $script:tree.TopNode = $next }
    }
})

# Clear highlight when the cursor leaves the tree without dropping.
$script:tree.add_DragLeave({ Clear-DropHighlight })

# Perform the actual move in the data model when the user releases the mouse.
$script:tree.add_DragDrop({
    param($s, $e)
    Clear-DropHighlight

    $pt         = $script:tree.PointToClient([System.Drawing.Point]::new($e.X, $e.Y))
    $targetNode = $script:tree.GetNodeAt($pt)
    if ($null -eq $targetNode -or $targetNode -eq $script:dragNode) { return }

    # Reject drops onto descendants of the dragged node.
    $check = $targetNode.Parent
    while ($null -ne $check) {
        if ($check -eq $script:dragNode) { return }
        $check = $check.Parent
    }

    $dragData     = $script:dragNode.Tag
    $targetData   = $targetNode.Tag
    $insertBefore = $script:dropBefore

    # Locate the dragged item in the data model and remove it.
    $dragList = $null; $dragParent = $null
    Find-ParentList $dragData ([ref]$dragList) ([ref]$dragParent) | Out-Null
    $dragIdx = $dragList.IndexOf($dragData)
    Push-UndoState
    $dragList.RemoveAt($dragIdx)

    # Re-locate the target after removal (its index may have shifted).
    $targetList = $null; $targetParent = $null
    Find-ParentList $targetData ([ref]$targetList) ([ref]$targetParent) | Out-Null
    $targetIdx = $targetList.IndexOf($targetData)

    if ($insertBefore) {
        $targetList.Insert($targetIdx, $dragData)
    } else {
        $targetList.Insert($targetIdx + 1, $dragData)
    }

    $data = $dragData
    Update-Tree
    foreach ($tv in $script:tree.Nodes) { Select-ByData $tv $data }
    Update-Preview
})

# Double-click → edit
$script:tree.add_DoubleClick({ $btnEdit.PerformClick() })
#endregion

#region ── File operations ───────────────────────────────────────────────────

# Load from file
$btnLoad.add_Click({
    $ofd        = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $ofd.Title  = "Load ManagedBookmarks JSON"
    if ((Test-HasText $GLB_jsonBasePath) -and (Test-Path -LiteralPath $GLB_jsonBasePath)) { $ofd.InitialDirectory = $GLB_jsonBasePath }
    if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    Add-LogEntry "Loading file" 'step'
    Add-LogKeyValue "Path" $ofd.FileName
    try {
        Push-UndoState
        Import-JsonString (Get-Content $ofd.FileName -Raw -Encoding UTF8)
        Add-LogEntry "File loaded" 'ok'
        Add-LogKeyValue "Root items" "$($script:bookmarks.Count)"
        Set-StatusLabel $actionLabel "Loaded: $($ofd.SafeFileName)" 'ok'
    } catch {
        Mark-ErrorRecordSeen $_
        Add-LogEntry "Error loading file" 'fail' -ErrorMessage $_.Exception.Message
        Set-StatusLabel $actionLabel "Error loading file." 'fail'
        $friendly = ConvertTo-FriendlyErrorMessage $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Error loading file:`n$friendly", "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

# Save to file
$btnSave.add_Click({
    # When JsonBasePath is configured: ask only for a filename, create the folder if needed,
    # and save there directly. Otherwise: fall back to the standard SaveFileDialog.
    if (Test-HasText $GLB_jsonBasePath) {
        # Create the folder automatically when it does not exist yet
        if (-not (Test-Path -LiteralPath $GLB_jsonBasePath)) {
            try {
                New-Item -ItemType Directory -Path $GLB_jsonBasePath -Force -ErrorAction Stop | Out-Null
            }
            catch {
                Mark-ErrorRecordSeen $_
                Add-LogEntry "Could not create JSON target folder" 'fail' -ErrorMessage $_.Exception.Message
                Set-StatusLabel $actionLabel "Could not create target folder for JSON." 'fail'
                [System.Windows.Forms.MessageBox]::Show("Could not create JSON folder:`n$GLB_jsonBasePath`n`n$($_.Exception.Message)", "Folder error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                return
            }
        }
        $fileName = Show-SaveNameDialog $GLB_jsonBasePath
        if (-not $fileName) { return }
        $targetPath = Join-Path $GLB_jsonBasePath $fileName
        # Overwrite confirmation when file already exists
        if (Test-Path $targetPath) {
            $ans = [System.Windows.Forms.MessageBox]::Show(
                "'$fileName' already exists. Overwrite?", "Confirm",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }
    } else {
        $targetPath = Show-SaveFileDialog -Title "Save ManagedBookmarks JSON" -Filter "JSON files (*.json)|*.json|All files (*.*)|*.*" -DefaultFileName "managed_bookmarks.json" -InitialDirectory ([System.Windows.Forms.Application]::StartupPath)
        if (-not (Test-HasText $targetPath)) { return }
    }

    Add-LogEntry "Saving file" 'step'
    Add-LogKeyValue "Path" $targetPath
    try {
        Build-Json | Set-Content $targetPath -Encoding UTF8
        Add-LogEntry "File saved" 'ok'
        Add-LogKeyValue "Root items" "$($script:bookmarks.Count)"
        Set-StatusLabel $actionLabel "Saved: $(Split-Path $targetPath -Leaf)" 'ok'
    } catch {
        Mark-ErrorRecordSeen $_
        Add-LogEntry "Error saving file" 'fail' -ErrorMessage $_.Exception.Message
        Set-StatusLabel $actionLabel "Error saving file." 'fail'
        $friendly = ConvertTo-FriendlyErrorMessage $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Error saving file:`n$friendly", "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

# Export deployment script
$btnExportScript.add_Click({
    if ($script:bookmarks.Count -eq 0) {
        Set-StatusLabel $actionLabel "Add at least one bookmark before exporting." 'warn'
        return
    }

    $params = Show-ExportScriptDialog
    if ($null -eq $params) { return }

    $scriptName = $params.name
    $description = $params.description
    $jsonValue  = Build-Json

    # Derive transcript/log root from Logging.BasePath (strip the \logs leaf)
    $loggingBase = if ($GLB_config -and $GLB_config.Logging.BasePath) {
        $GLB_config.Logging.BasePath
    } else { 'C:\ProgramData\Scripting' }
    $scriptingRoot = (Split-Path $loggingBase -Parent).TrimEnd('\') + '\'

    $authorName  = if ($GLB_companyAuthor) { $GLB_companyAuthor } else { $env:USERNAME }
    $companyName = if ($GLB_companyName)   { $GLB_companyName   } else { 'Company' }
    $regBase     = if ($GLB_companyRegBase){ $GLB_companyRegBase } else { 'HKLM:\Software\Company\Scripting' }
    $dateStr     = (Get-Date -Format "yyyy.MM.dd")
    $version     = "1.00"

    $content = New-ManagedBookmarksScript `
        -ScriptName    $scriptName `
        -Description   $description `
        -JsonValue     $jsonValue `
        -CompanyName   $companyName `
        -Author        $authorName `
        -RegistryBase  $regBase `
        -ScriptingRoot $scriptingRoot `
        -Date          $dateStr `
        -Version       $version

    $fileName   = "$scriptName.ps1"
    if (Test-HasText $GLB_scriptBasePath) {
        if (-not (Test-Path -LiteralPath $GLB_scriptBasePath)) {
            try {
                New-Item -ItemType Directory -Path $GLB_scriptBasePath -Force -ErrorAction Stop | Out-Null
            }
            catch {
                Mark-ErrorRecordSeen $_
                Add-LogEntry "Could not create script target folder" 'fail' -ErrorMessage $_.Exception.Message
                Set-StatusLabel $actionLabel "Could not create target folder for script export." 'fail'
                [System.Windows.Forms.MessageBox]::Show("Could not create script folder:`n$GLB_scriptBasePath`n`n$($_.Exception.Message)", "Folder error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                return
            }
        }
        $targetPath = Join-Path $GLB_scriptBasePath $fileName
    } else {
        $targetPath = Show-SaveFileDialog -Title "Export Deployment Script" -Filter "PowerShell scripts (*.ps1)|*.ps1|All files (*.*)|*.*" -DefaultFileName $fileName -InitialDirectory ([System.Windows.Forms.Application]::StartupPath)
        if (-not (Test-HasText $targetPath)) { return }
    }

    # Overwrite confirmation
    if (Test-Path $targetPath) {
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "'$fileName' already exists. Overwrite?", "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    Add-LogEntry "Exporting deployment script" 'step'
    Add-LogKeyValue "Path" $targetPath
    try {
        $content | Set-Content $targetPath -Encoding UTF8
        Add-LogEntry "Script exported" 'ok'
        Add-LogKeyValue "Script name" $scriptName
        Set-StatusLabel $actionLabel "Script saved: $fileName" 'ok'
    } catch {
        Mark-ErrorRecordSeen $_
        Add-LogEntry "Error exporting script" 'fail' -ErrorMessage $_.Exception.Message
        Set-StatusLabel $actionLabel "Error saving script." 'fail'
        $friendly = ConvertTo-FriendlyErrorMessage $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Error saving script:`n$friendly", "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

# Copy to clipboard
$btnCopy.add_Click({
    [System.Windows.Forms.Clipboard]::SetText((Build-Json))
    Add-LogEntry "JSON copied to clipboard." 'ok'
    Set-StatusLabel $actionLabel "JSON copied to clipboard." 'ok'
})

# Import from pasted string
$btnImportStr.add_Click({
    $raw = $txtImport.Text.Trim()
    if ($raw -eq "") {
        Set-StatusLabel $actionLabel "Paste a JSON string in the input field first." 'warn'
        return
    }
    Add-LogEntry "Importing JSON string" 'step'
    try {
        Push-UndoState
        Import-JsonString $raw
        $tabControl.SelectedTab = $tabPreview
        Add-LogEntry "JSON string imported" 'ok'
        Add-LogKeyValue "Root items" "$($script:bookmarks.Count)"
        Set-StatusLabel $actionLabel "JSON imported — $($script:bookmarks.Count) root items." 'ok'
    } catch {
        Mark-ErrorRecordSeen $_
        Add-LogEntry "Invalid JSON" 'fail' -ErrorMessage $_.Exception.Message
        Set-StatusLabel $actionLabel "Invalid JSON." 'fail'
        $friendly = ConvertTo-FriendlyErrorMessage $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Invalid JSON:`n$friendly", "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$btnUndo.add_Click({
    if (Invoke-Undo) {
        Add-LogEntry "Undo applied" 'ok'
        Set-StatusLabel $actionLabel "Undo applied." 'ok'
    } else {
        Set-StatusLabel $actionLabel "Nothing to undo." 'warn'
    }
})

$btnRedo.add_Click({
    if (Invoke-Redo) {
        Add-LogEntry "Redo applied" 'ok'
        Set-StatusLabel $actionLabel "Redo applied." 'ok'
    } else {
        Set-StatusLabel $actionLabel "Nothing to redo." 'warn'
    }
})

$btnValidateContract.add_Click({
    Add-LogEntry "Running JSON contract test" 'step'
    try {
        if (Test-JsonRoundtripContract) {
            Add-LogEntry "Contract test passed" 'ok'
            Set-StatusLabel $actionLabel "Contract test passed (import/export roundtrip)." 'ok'
        } else {
            Add-LogEntry "Contract test failed" 'fail' -ErrorMessage "Roundtrip produced a different JSON payload."
            Set-StatusLabel $actionLabel "Contract test failed." 'fail'
        }
    }
    catch {
        Mark-ErrorRecordSeen $_
        Add-LogEntry "Contract test error" 'fail' -ErrorMessage $_.Exception.Message
        Set-StatusLabel $actionLabel "Contract test error." 'fail'
    }
})

$btnClearImport.add_Click({ $txtImport.Clear() })

# Registry tab handlers (admin only)
if ($isAdmin) {
    $script:regPaths = @(
        @{ Path = 'HKLM:\SOFTWARE\Policies\Google\Chrome';    Name = 'ManagedBookmarks' },
        @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge';   Name = 'ManagedFavorites' }
    )

    $btnReadReg.add_Click({
        $entry = $script:regPaths[$cmbRegBrowser.SelectedIndex]
        Add-LogEntry "Reading registry: $($entry.Path) → $($entry.Name)" 'step'
        try {
            $val = Get-ItemPropertyValue -Path $entry.Path -Name $entry.Name -ErrorAction Stop
            $txtReg.Text = $val
            Add-LogEntry "Registry value read" 'ok'
            Set-StatusLabel $actionLabel "Registry value loaded ($($entry.Name))." 'ok'
        } catch {
            Mark-ErrorRecordSeen $_
            $txtReg.Text = ""
            Add-LogEntry "Registry read failed" 'fail' -ErrorMessage $_.Exception.Message
            Set-StatusLabel $actionLabel "Registry key not found or no value set." 'warn'
        }
    })

    $btnLoadReg.add_Click({
        $raw = $txtReg.Text.Trim()
        if ($raw -eq "") {
            Set-StatusLabel $actionLabel "Read a registry value first before loading." 'warn'
            return
        }
        Add-LogEntry "Importing registry value into editor" 'step'
        try {
            Push-UndoState
            Import-JsonString $raw
            $tabControl.SelectedTab = $tabPreview
            Add-LogEntry "Registry JSON imported" 'ok'
            Add-LogKeyValue "Root items" "$($script:bookmarks.Count)"
            Set-StatusLabel $actionLabel "Registry JSON loaded — $($script:bookmarks.Count) root items." 'ok'
        } catch {
            Mark-ErrorRecordSeen $_
            Add-LogEntry "Invalid JSON in registry value" 'fail' -ErrorMessage $_.Exception.Message
            Set-StatusLabel $actionLabel "Registry value is not valid JSON." 'fail'
            $friendly = ConvertTo-FriendlyErrorMessage $_.Exception.Message
            [System.Windows.Forms.MessageBox]::Show("Invalid JSON:`n$friendly", "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    })
}

# Write to Registry button (admin only)
if ($isAdmin) {
    $btnWriteRegistry.add_Click({
        if ($script:bookmarks.Count -eq 0) {
            Set-StatusLabel $actionLabel "No bookmarks to write — add items first." 'warn'
            return
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "This will overwrite the ManagedBookmarks / ManagedFavorites values in HKLM for Chrome and Edge.`n`nContinue?",
            "Write to Registry",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $json            = Build-Json
        $chromePath      = 'HKLM:\SOFTWARE\Policies\Google\Chrome'
        $edgePath        = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
        $errMessages     = @()

        Add-LogEntry "Writing bookmarks to registry" 'step'

        # Chrome
        try {
            if (-not (Test-Path $chromePath)) { New-Item -Path $chromePath -Force | Out-Null }
            Set-ItemProperty -Path $chromePath -Name 'ManagedBookmarks'   -Value $json -Type String
            Set-ItemProperty -Path $chromePath -Name 'BookmarkBarEnabled' -Value 1     -Type DWord
            Add-LogEntry "Chrome: ManagedBookmarks + BookmarkBarEnabled set" 'ok'
        } catch {
            Mark-ErrorRecordSeen $_
            $errMessages += "Chrome: $($_.Exception.Message)"
            Add-LogEntry "Chrome registry write failed" 'fail' -ErrorMessage $_.Exception.Message
        }

        # Edge
        try {
            if (-not (Test-Path $edgePath)) { New-Item -Path $edgePath -Force | Out-Null }
            Set-ItemProperty -Path $edgePath -Name 'ManagedFavorites'    -Value $json -Type String
            Set-ItemProperty -Path $edgePath -Name 'FavoritesBarEnabled' -Value 1     -Type DWord
            Add-LogEntry "Edge: ManagedFavorites + FavoritesBarEnabled set" 'ok'
        } catch {
            Mark-ErrorRecordSeen $_
            $errMessages += "Edge: $($_.Exception.Message)"
            Add-LogEntry "Edge registry write failed" 'fail' -ErrorMessage $_.Exception.Message
        }

        if ($errMessages.Count -eq 0) {
            Set-StatusLabel $actionLabel "Registry updated for Chrome and Edge." 'ok'
            [System.Windows.Forms.MessageBox]::Show(
                "Registry updated successfully.`n`nRestart Chrome and Edge fully (including system tray) to apply the new bookmarks.",
                "Write to Registry",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        } else {
            Set-StatusLabel $actionLabel "Registry write completed with errors." 'fail'
            [System.Windows.Forms.MessageBox]::Show(
                "One or more writes failed:`n`n$($errMessages -join "`n")",
                "Write to Registry — Errors",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    })
}
#endregion

$null = $form.ShowDialog()
$form.Dispose()

#region ── Write session log ─────────────────────────────────────────────────
if ($script:outputModuleLoaded -and $GLB_logPath -and $logListBox.Items.Count -gt 0) {
    Add-LogEntry '' 'separator'
    Add-LogEntry "Session ended" 'info'
    Add-LogKeyValue "Time"       "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
    Add-LogKeyValue "Root items" "$($script:bookmarks.Count)"
    try {
        Write-DecosterOutputLog -Listbox $logListBox -LogPath $GLB_logPath `
            -Username $env:USERNAME -Suffix "ManagedBookmarksCreator"
        Write-AppHostMessage "Log written to: $GLB_logPath" "Cyan"
    }
    catch {
        Mark-ErrorRecordSeen $_
        Write-AppHostMessage "Could not write log: $($_.Exception.Message)" "Yellow"
        Add-FallbackLogLine -Text 'Primary log write failed' -State 'warn' -ErrorMessage $_.Exception.Message
    }
}
#endregion




