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

$GLB_scriptVersion     = "2.9.1.0"
$GLB_ScriptUpdateDate  = "19/03/2026"
$GLB_scriptcontributer = "Decoster Hans"
$GLB_ScriptTitel       = "Managed Bookmarks Creator"

#region load assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#endregion

function Test-HasText {
    param([string]$Value)
    return -not [string]::IsNullOrWhiteSpace($Value)
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

    if ($script:isPackagedExecutable) { return }
    Write-Host $Message -ForegroundColor $Color
}

function Get-AppSearchRoots {
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

if ($script:outputModuleLoaded -and $GLB_config -and $GLB_config.Logging.BasePath) {
    $GLB_logPath = Join-Path $GLB_config.Logging.BasePath "ManagedBookmarksCreator"
    Remove-DecosterOutputOldLogs -LogPath $GLB_logPath -RetentionDays $GLB_config.Logging.LogRetentionDays
}
#endregion

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
        Write-AppHostMessage "[Log] $State : $Text $(if($Detail){"| $Detail"}) — $($_.Exception.Message)" "DarkYellow"
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
$script:bookmarks = [System.Collections.Generic.List[object]]::new()
#endregion

#region ── Data helpers ──────────────────────────────────────────────────────
function New-FolderNode($name) {
    return @{ type = "folder"; name = $name; children = [System.Collections.Generic.List[object]]::new() }
}

function New-UrlNode($name, $url) {
    return @{ type = "url"; name = $name; url = $url }
}

function ConvertTo-NodeObject($node) {
    if ($node.type -eq "folder") {
        return [PSCustomObject]@{
            name     = $node.name
            children = @($node.children | ForEach-Object { ConvertTo-NodeObject $_ })
        }
    }
    return [PSCustomObject]@{ name = $node.name; url = $node.url }
}

function Build-Json {
    $bookmarkObjects = @($script:bookmarks | ForEach-Object { ConvertTo-NodeObject $_ })
    $name = if ($null -ne $script:txtTopLevelName) { $script:txtTopLevelName.Text.Trim() } else { '' }
    if ($name -ne '') {
        $all = @([PSCustomObject]@{ toplevel_name = $name }) + $bookmarkObjects
    } else {
        $all = $bookmarkObjects
    }
    return (ConvertTo-Json -InputObject $all -Compress -Depth 20)
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
    $script:tree.Nodes.Clear()
    foreach ($node in $script:bookmarks) {
        $tv = Add-TreeNode $null $node
        if ($node.type -eq "folder") { Add-ChildrenToTree $tv $node }
    }
    $script:tree.ExpandAll()
}

function Add-ChildrenToTree($tvParent, $dataParent) {
    foreach ($child in $dataParent.children) {
        $tv = Add-TreeNode $tvParent $child
        if ($child.type -eq "folder") { Add-ChildrenToTree $tv $child }
    }
}

function Find-ParentList($dataNode, [ref]$outList, [ref]$outParent) {
    foreach ($n in $script:bookmarks) {
        if ($n -eq $dataNode) { $outList.Value = $script:bookmarks; $outParent.Value = $null; return $true }
        if ($n.type -eq "folder" -and (Search-Children $n $dataNode $outList $outParent)) { return $true }
    }
    return $false
}

function Search-Children($folder, $target, [ref]$outList, [ref]$outParent) {
    foreach ($c in $folder.children) {
        if ($c -eq $target) { $outList.Value = $folder.children; $outParent.Value = $folder; return $true }
        if ($c.type -eq "folder" -and (Search-Children $c $target $outList $outParent)) { return $true }
    }
    return $false
}

function Select-ByData($tvNode, $data) {
    if ($tvNode.Tag -eq $data) { $script:tree.SelectedNode = $tvNode; return }
    foreach ($child in $tvNode.Nodes) { Select-ByData $child $data }
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
    $lines.Add('Stop-Transcript | Out-Null')
    $lines.Add('Start-Transcript -Path "$transcriptPath$transcriptName" -Append')
    $lines.Add('')
    $lines.Add('# ---------------------------------------------------------------------------')
    $lines.Add('# Check admin rights')
    $lines.Add('# ---------------------------------------------------------------------------')
    $lines.Add('if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {')
    $lines.Add('    Write-Host "Not running as Administrator. Please restart with admin privileges." -ForegroundColor Red')
    $lines.Add('    Stop-Transcript')
    $lines.Add('    exit 1')
    $lines.Add('}')
    $lines.Add('')
    $lines.Add('# Get OS')
    $lines.Add('$os = (Get-WmiObject -Class Win32_OperatingSystem).Caption')
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
    $lines.Add('Stop-Transcript')

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

    if ($TextBox.PSObject.Properties.Name -contains 'PlaceholderText') {
        $TextBox.PlaceholderText = $Placeholder
        return
    }

    if (-not ('Win32CueBanner' -as [type])) { return }

    $TextBox.AccessibleDescription = $Placeholder
    $applyCueBanner = {
        param($sender, $e)
        if ($sender.IsHandleCreated) {
            [void][Win32CueBanner]::SendMessage($sender.Handle, 0x1501, [IntPtr]1, [string]$sender.AccessibleDescription)
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
#endregion

#region ── Dialogs ───────────────────────────────────────────────────────────
function Show-FolderDialog($title, $defaultName = "") {
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
        return @{ name = ($txtName.Text.Trim() -replace '\s+','_'); description = $txtDesc.Text.Trim() }
    }
    return $null
}

function Show-LinkDialog($title, $defaultName = "", $defaultUrl = "") {
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
        return @{ name = $txtN.Text.Trim(); url = $txtU.Text.Trim() }
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

$toolStrip.Items.AddRange(@(
    $btnAddFolder, $btnAddSubFolder, $btnAddLink, $btnAddRootLink,
    $sep1, $btnEdit, $btnDelete, $sep2, $btnMoveUp, $btnMoveDown,
    $sep3, $btnLoad, $btnSave, $btnCopy,
    $sep4, $btnExportScript,
    $sep5, $btnWriteRegistry
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

$script:txtTopLevelName.add_TextChanged({ Update-Preview })

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
    if ($script:configWarnings.Count -gt 0) {
        Set-StatusLabel $actionLabel "Config incomplete — missing keys: $($script:configWarnings -join ', ')" 'warn'
    }
    $btnExportScript.Available = $true

})
$form.add_Resize({ Invoke-Layout })
#endregion

#region ── Button events ─────────────────────────────────────────────────────

# Add root folder
$btnAddFolder.add_Click({
    $name = Show-FolderDialog "New root folder"
    if ($null -eq $name) { return }
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
    if ($idx -le 0) { return }
    $outList.RemoveAt($idx)
    $outList.Insert($idx - 1, $sel.Tag)
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
    if ($idx -ge $outList.Count - 1) { return }
    $outList.RemoveAt($idx)
    $outList.Insert($idx + 1, $sel.Tag)
    $data = $sel.Tag
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
        Import-JsonString (Get-Content $ofd.FileName -Raw -Encoding UTF8)
        Add-LogEntry "File loaded" 'ok'
        Add-LogKeyValue "Root items" "$($script:bookmarks.Count)"
        Set-StatusLabel $actionLabel "Loaded: $($ofd.SafeFileName)" 'ok'
    } catch {
        Add-LogEntry "Error loading file" 'fail' -ErrorMessage $_.Exception.Message
        Set-StatusLabel $actionLabel "Error loading file." 'fail'
        [System.Windows.Forms.MessageBox]::Show("Error loading file:`n$($_.Exception.Message)", "Error",
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
            New-Item -ItemType Directory -Path $GLB_jsonBasePath -Force | Out-Null
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
        Add-LogEntry "Error saving file" 'fail' -ErrorMessage $_.Exception.Message
        Set-StatusLabel $actionLabel "Error saving file." 'fail'
        [System.Windows.Forms.MessageBox]::Show("Error saving file:`n$($_.Exception.Message)", "Error",
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
            New-Item -ItemType Directory -Path $GLB_scriptBasePath -Force | Out-Null
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
        Add-LogEntry "Error exporting script" 'fail' -ErrorMessage $_.Exception.Message
        Set-StatusLabel $actionLabel "Error saving script." 'fail'
        [System.Windows.Forms.MessageBox]::Show("Error saving script:`n$($_.Exception.Message)", "Error",
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
        Import-JsonString $raw
        $tabControl.SelectedTab = $tabPreview
        Add-LogEntry "JSON string imported" 'ok'
        Add-LogKeyValue "Root items" "$($script:bookmarks.Count)"
        Set-StatusLabel $actionLabel "JSON imported — $($script:bookmarks.Count) root items." 'ok'
    } catch {
        Add-LogEntry "Invalid JSON" 'fail' -ErrorMessage $_.Exception.Message
        Set-StatusLabel $actionLabel "Invalid JSON." 'fail'
        [System.Windows.Forms.MessageBox]::Show("Invalid JSON:`n$($_.Exception.Message)", "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
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
            Import-JsonString $raw
            $tabControl.SelectedTab = $tabPreview
            Add-LogEntry "Registry JSON imported" 'ok'
            Add-LogKeyValue "Root items" "$($script:bookmarks.Count)"
            Set-StatusLabel $actionLabel "Registry JSON loaded — $($script:bookmarks.Count) root items." 'ok'
        } catch {
            Add-LogEntry "Invalid JSON in registry value" 'fail' -ErrorMessage $_.Exception.Message
            Set-StatusLabel $actionLabel "Registry value is not valid JSON." 'fail'
            [System.Windows.Forms.MessageBox]::Show("Invalid JSON:`n$($_.Exception.Message)", "Error",
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
        Write-AppHostMessage "Could not write log: $($_.Exception.Message)" "Yellow"
    }
}
#endregion




