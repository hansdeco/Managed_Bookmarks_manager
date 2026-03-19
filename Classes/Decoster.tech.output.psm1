#
# decoster.tech.output.psm1
# Shared output formatting and UI helper functions for DecosterOutput scripts
# (DecomVMWithGui, DHCPreservation_AND_VMCreation_GUI, Vsphere_Customisation_ADJoin_passwordChange).
#
# Usage:
#   Import-Module (Join-Path $PSScriptRoot "decoster.tech.output.psm1")
#
# Sections:
#   1. Text formatting   — return [string[]], fully thread-safe, usable in RunspacePool workers
#   2. UI helpers        — update WinForms controls; call from the UI thread only
#   3. Logging           — write ListBox contents to a timestamped log file
#
# Thread-safety note:
#   Formatting functions only build strings and never touch WinForms controls.
#   In a parallel worker (V0.5 pattern), call them and Enqueue each returned line:
#
#       foreach ($line in (Format-DecosterOutputOk $specName)) { $outputQueue.Enqueue($line) }
#
#   On the UI thread, pass the lines directly to Add-DecosterOutputOutput:
#
#       Add-DecosterOutputOutput -Listbox $textBoxOutPut -Lines (Format-DecosterOutputOk $specName) -AutoScroll
#

Set-StrictMode -Version Latest

# Load WinForms and Drawing assemblies at module load time so that the
# type constraints on the UI helper parameters resolve correctly.
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Drawing       -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# Divider constants — single definition, reused by all formatting functions.
# Change here once to update every divider across all scripts.
# ---------------------------------------------------------------------------
$script:DecosterOutputDividerHeavy = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
$script:DecosterOutputDividerThin  = '---'

#region Text formatting
# ---------------------------------------------------------------------------
# All Format-DecosterOutput* functions return [string[]] so the caller can either
# - Add-DecosterOutputOutput  (UI thread)
# - Enqueue each element into a ConcurrentQueue (worker thread)
# ---------------------------------------------------------------------------

function Format-DecosterOutputHeader {
    <#
    .SYNOPSIS  Section header line.
    .EXAMPLE   Format-DecosterOutputHeader "ITX" "server.decoster.tech"
               → @("=== ITX (server.decoster.tech) ===")
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$Url = ""
    )
    if ($Url) { return [string[]]@("=== $Name ($Url) ===") }
    return [string[]]@("=== $Name ===")
}

function Format-DecosterOutputOk {
    <#
    .SYNOPSIS  Success line.
    .EXAMPLE   Format-DecosterOutputOk "DT-W2022-DHCP-DomainJoined-DEV"
               → @("  [OK]   DT-W2022-DHCP-DomainJoined-DEV")
    #>
    param([Parameter(Mandatory)][string]$Name)
    return [string[]]@("  [OK]   $Name")
}

function Format-DecosterOutputFail {
    <#
    .SYNOPSIS  Failure line, with optional error detail on a second line.
    .EXAMPLE   Format-DecosterOutputFail "DT-W2022-..." "Access denied"
               → @("  [FAIL] DT-W2022-...", "         Access denied")
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$ErrorMessage = ""
    )
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("  [FAIL] $Name")
    if ($ErrorMessage) { $lines.Add("         $ErrorMessage") }
    return $lines.ToArray()
}

function Format-DecosterOutputSkip {
    <#
    .SYNOPSIS  Skipped line (item not found / not applicable).
    .EXAMPLE   Format-DecosterOutputSkip "DT-W2022-DHCP-DomainJoined-DEV"
               → @("  [SKIP] DT-W2022-DHCP-DomainJoined-DEV")
    #>
    param([Parameter(Mandatory)][string]$Name)
    return [string[]]@("  [SKIP] $Name")
}

function Format-DecosterOutputWarn {
    <#
    .SYNOPSIS  Warning line for a named item, with optional detail and error message.
    .EXAMPLE   Format-DecosterOutputWarn "DT-W2022-..." "using defaults: user / domain" "Could not read spec"
               → @("  [WARN] DT-W2022-...  (using defaults: user / domain)", "         Could not read spec")
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$Detail       = "",
        [string]$ErrorMessage = ""
    )
    $lines = [System.Collections.Generic.List[string]]::new()
    if ($Detail) {
        $lines.Add("  [WARN] $Name  ($Detail)")
    } else {
        $lines.Add("  [WARN] $Name")
    }
    if ($ErrorMessage) { $lines.Add("         $ErrorMessage") }
    return $lines.ToArray()
}

function Format-DecosterOutputWarnResult {
    <#
    .SYNOPSIS  Follow-up result line after a [WARN] prefix — the action still ran with defaults.
    .PARAMETER Success  $true → "-> [OK]";  $false → "-> [FAIL]"
    .EXAMPLE   Format-DecosterOutputWarnResult -Success $true  → @("         -> [OK]")
               Format-DecosterOutputWarnResult -Success $false "Timeout" → @("         -> [FAIL]", "         Timeout")
    #>
    param(
        [Parameter(Mandatory)][bool]$Success,
        [string]$ErrorMessage = ""
    )
    $lines = [System.Collections.Generic.List[string]]::new()
    $resultTag = if ($Success) { "[OK]" } else { "[FAIL]" }
    $lines.Add("         -> $resultTag")
    if (-not $Success -and $ErrorMessage) { $lines.Add("         $ErrorMessage") }
    return $lines.ToArray()
}

function Format-DecosterOutputError {
    <#
    .SYNOPSIS  Hard error line (e.g. module load failure, unexpected exception).
    .EXAMPLE   Format-DecosterOutputError "PowerCLI module could not be loaded"
               → @("  [ERROR] PowerCLI module could not be loaded")
    #>
    param([Parameter(Mandatory)][string]$Message)
    return [string[]]@("  [ERROR] $Message")
}

function Format-DecosterOutputInfo {
    <#
    .SYNOPSIS  Indented informational line (no status prefix).
    .EXAMPLE   Format-DecosterOutputInfo "Connected"  → @("  Connected")
    .EXAMPLE   Format-DecosterOutputInfo "User   : decoster.tech\svc-vmware"  → @("  User   : decoster.tech\svc-vmware")
    #>
    param([Parameter(Mandatory)][string]$Message)
    return [string[]]@("  $Message")
}

function Format-DecosterOutputConnected {
    <#
    .SYNOPSIS  Standard vCenter connection success line.
    #>
    return [string[]]@("  Connected")
}

function Format-DecosterOutputConnectionFail {
    <#
    .SYNOPSIS  Standard vCenter connection failure line.
    .EXAMPLE   Format-DecosterOutputConnectionFail "Timeout after 30s"
               → @("  FAILURE - Could not connect: Timeout after 30s")
    #>
    param([string]$ErrorMessage = "")
    if ($ErrorMessage) { return [string[]]@("  FAILURE - Could not connect: $ErrorMessage") }
    return [string[]]@("  FAILURE - Could not connect")
}

function Format-DecosterOutputSummary {
    <#
    .SYNOPSIS  Per-vCenter (or per-phase) summary line.
    .EXAMPLE   Format-DecosterOutputSummary -Ok 15 -Skipped 0 -Failed 0
               → @("  Summary: 15 OK | 0 skipped | 0 failed")
    #>
    param(
        [Parameter(Mandatory)][int]$Ok,
        [Parameter(Mandatory)][int]$Skipped,
        [Parameter(Mandatory)][int]$Failed
    )
    return [string[]]@("  Summary: $Ok OK | $Skipped skipped | $Failed failed")
}

function Format-DecosterOutputTotal {
    <#
    .SYNOPSIS  Grand total line at the end of a run.
    .EXAMPLE   Format-DecosterOutputTotal -Ok 45 -Skipped 0 -Failed 0
               → @("Total: 45 OK | 0 skipped | 0 failed")
    #>
    param(
        [Parameter(Mandatory)][int]$Ok,
        [Parameter(Mandatory)][int]$Skipped,
        [Parameter(Mandatory)][int]$Failed
    )
    return [string[]]@("Total: $Ok OK | $Skipped skipped | $Failed failed")
}

function Format-DecosterOutputSeparator {
    <#
    .SYNOPSIS  Thin separator line ("---"), used before the grand total.
    #>
    return [string[]]@($script:DecosterOutputDividerThin)
}

function Format-DecosterOutputDivider {
    <#
    .SYNOPSIS  Heavy divider line (XXXXX...), used in warning and failure blocks.
    .NOTES     Both the thin and heavy divider strings are defined once as module-level
               constants ($script:DecosterOutputDividerThin / $script:DecosterOutputDividerHeavy).
               All formatting functions that need a divider reference those constants,
               so changing the style here propagates everywhere automatically.
    #>
    return [string[]]@($script:DecosterOutputDividerHeavy)
}

function Format-DecosterOutputStep {
    <#
    .SYNOPSIS  Numbered or unnumbered step indicator for multi-step operations.
    .DESCRIPTION
        When Step and Total are both provided the prefix is "[Step/Total]".
        When omitted the prefix is "[>>]".
    .PARAMETER Message
        The step description.
    .PARAMETER Step
        Current step number (optional). Must be paired with Total.
    .PARAMETER Total
        Total number of steps (optional). Must be paired with Step.
    .EXAMPLE
        Format-DecosterOutputStep "Bestand laden" -Step 1 -Total 3
        → @("  [1/3] Bestand laden")
    .EXAMPLE
        Format-DecosterOutputStep "Validating input"
        → @("  [>>] Validating input")
    #>
    param(
        [Parameter(Mandatory)][string]$Message,
        [int]$Step  = 0,
        [int]$Total = 0
    )
    if ($Step -gt 0 -and $Total -gt 0) {
        return [string[]]@("  [$Step/$Total] $Message")
    }
    return [string[]]@("  [>>] $Message")
}

function Format-DecosterOutputSubItem {
    <#
    .SYNOPSIS  Indented sub-item detail line, visually nested under a preceding log line.
    .DESCRIPTION
        Use to add extra detail lines beneath an Ok / Warn / Fail entry without
        repeating the status prefix.
    .EXAMPLE
        Format-DecosterOutputSubItem "Naam : Intranet"
        → @("         └ Naam : Intranet")
    #>
    param([Parameter(Mandatory)][string]$Message)
    return [string[]]@("         └ $Message")
}

function Format-DecosterOutputWarningBlock {
    <#
    .SYNOPSIS  Heavy warning block (decom/creation style) — blank + divider + message + divider + blank.
    .EXAMPLE   Format-DecosterOutputWarningBlock "VM already present on vCenter"
    #>
    param([Parameter(Mandatory)][string]$Message)
    return [string[]]@("", $script:DecosterOutputDividerHeavy, "!! WARNING  $Message", $script:DecosterOutputDividerHeavy, "")
}

function Format-DecosterOutputFailureBlock {
    <#
    .SYNOPSIS  Heavy failure block (decom/creation style) — blank + divider + message + divider + blank.
    .EXAMPLE   Format-DecosterOutputFailureBlock "Could not disable AD object" "Access denied"
    #>
    param(
        [Parameter(Mandatory)][string]$Message,
        [string]$ErrorMessage = ""
    )
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("")
    $lines.Add($script:DecosterOutputDividerHeavy)
    $lines.Add("!! FAILURE !!  $Message")
    if ($ErrorMessage) { $lines.Add("!! ErrorMessage !!  $ErrorMessage") }
    $lines.Add($script:DecosterOutputDividerHeavy)
    $lines.Add("")
    return $lines.ToArray()
}

function Format-DecosterOutputKeyValue {
    <#
    .SYNOPSIS  Aligned key : value line (matches the set-output pattern in the decom script).
    .EXAMPLE   Format-DecosterOutputKeyValue "Server name" $serverName   → @("  Server name : DT-SERVER-001")
    #>
    param(
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)][string]$Value
    )
    return [string[]]@("  $Key : $Value")
}
#endregion

#region UI helpers (UI thread only)
# ---------------------------------------------------------------------------
# These functions directly modify WinForms controls.
# Must be called from the thread that created the form.
# ---------------------------------------------------------------------------

function Add-DecosterOutputOutput {
    <#
    .SYNOPSIS
        Adds one or more strings to a ListBox and optionally auto-scrolls to the last item.
    .PARAMETER Listbox
        The target System.Windows.Forms.ListBox.
    .PARAMETER Lines
        One or more strings to add. Accepts the [string[]] returned by Format-DecosterOutput* functions.
    .PARAMETER AutoScroll
        When present, scrolls the ListBox to show the most-recently added item.
    .EXAMPLE
        Add-DecosterOutputOutput -Listbox $textBoxOutPut -Lines (Format-DecosterOutputOk "spec-name") -AutoScroll
    #>
    param(
        [Parameter(Mandatory)][System.Windows.Forms.ListBox]$Listbox,
        [Parameter(Mandatory)][string[]]$Lines,
        [switch]$AutoScroll
    )
    foreach ($line in $Lines) {
        [void]$Listbox.Items.Add($line)
    }
    if ($AutoScroll -and $Listbox.Items.Count -gt 0) {
        $Listbox.TopIndex = $Listbox.Items.Count - 1
    }
}

function Set-DecosterOutputProgress {
    <#
    .SYNOPSIS
        Updates a ProgressBar value and an optional status Label in one call.
    .PARAMETER ProgressBar
        The target System.Windows.Forms.ProgressBar.
    .PARAMETER StatusLabel
        Optional System.Windows.Forms.Label to update with the Text parameter.
    .PARAMETER Value
        Progress value 0–100. Automatically clamped to [0, 100].
    .PARAMETER Text
        Status text to set on StatusLabel. Ignored when empty or StatusLabel is $null.
    .EXAMPLE
        Set-DecosterOutputProgress -ProgressBar $progressBar -StatusLabel $statusLabel -Value 50 -Text "Connecting..."
    #>
    param(
        [Parameter(Mandatory)][System.Windows.Forms.ProgressBar]$ProgressBar,
        [System.Windows.Forms.Label]$StatusLabel = $null,
        [Parameter(Mandatory)][int]$Value,
        [string]$Text = ""
    )
    $ProgressBar.Value = [Math]::Max(0, [Math]::Min(100, $Value))
    if ($null -ne $StatusLabel -and $Text) {
        $StatusLabel.Text = $Text
    }
}

function Set-DecosterOutputCredentialStatus {
    <#
    .SYNOPSIS
        Updates the credential indicator Panel colour and the adjacent Label text.
    .DESCRIPTION
        Matches the lightPanelcredInfo + credentialInfoLabel pattern used in the DecosterOutput scripts.
        Success  → Cyan panel, verified text.
        Failure  → Red panel, failed text.
        Pending  → Orange panel, custom or default pending text.
    .PARAMETER Panel
        The small coloured System.Windows.Forms.Panel used as a status LED.
    .PARAMETER Label
        The System.Windows.Forms.Label next to the panel.
    .PARAMETER State
        'ok', 'fail', or 'pending'.
    .PARAMETER Text
        Override the default message for this state.
    .EXAMPLE
        Set-DecosterOutputCredentialStatus -Panel $lightPanelcredInfo -Label $credentialInfoLabel -State 'ok'
    #>
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Panel]$Panel,
        [Parameter(Mandatory)][System.Windows.Forms.Label]$Label,
        [Parameter(Mandatory)][ValidateSet('ok','fail','pending')][string]$State,
        [string]$Text = ""
    )
    switch ($State) {
        'ok' {
            $Panel.BackColor = [System.Drawing.Color]::Cyan
            $Label.Text      = if ($Text) { $Text } else { "Domain credentials verified." }
        }
        'fail' {
            $Panel.BackColor = [System.Drawing.Color]::Red
            $Label.Text      = if ($Text) { $Text } else { "Domain credential verification failed." }
        }
        'pending' {
            $Panel.BackColor = [System.Drawing.Color]::Orange
            $Label.Text      = if ($Text) { $Text } else { "" }
        }
    }
}

function Set-DecosterOutputLabelStatus {
    <#
    .SYNOPSIS
        Sets a Label's text and ForeColor in one call.
    .DESCRIPTION
        Matches the pattern used for validation labels in the creation/decom scripts
        (LabelServerNameVerification, LabelJiraInfo, LabelServerIPInfo, LabelFetchStatus, ...).
    .PARAMETER Label
        The target System.Windows.Forms.Label.
    .PARAMETER Text
        The text to display.
    .PARAMETER State
        'ok' → Green, 'fail' → Red, 'warn' → Orange, 'info' → Black (default).
    .EXAMPLE
        Set-DecosterOutputLabelStatus -Label $LabelJiraInfo -Text "added" -State 'ok'
        Set-DecosterOutputLabelStatus -Label $LabelJiraInfo -Text "Ticket name incorrect" -State 'fail'
    #>
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Label]$Label,
        [Parameter(Mandatory)][string]$Text,
        [ValidateSet('ok','fail','warn','info')][string]$State = 'info'
    )
    $Label.Text = $Text
    $Label.ForeColor = switch ($State) {
        'ok'   { [System.Drawing.Color]::Green  }
        'fail' { [System.Drawing.Color]::Red    }
        'warn' { [System.Drawing.Color]::Orange }
        'info' { [System.Drawing.Color]::Black  }
    }
}
#endregion

#region Logging
function Remove-DecosterOutputOldLogs {
    <#
    .SYNOPSIS
        Deletes log files older than RetentionDays from a log folder.
    .DESCRIPTION
        Called at script startup to enforce the log-retention policy configured in
        decoster.tech.config.psd1 (Logging.LogRetentionDays).
        Only files with extension .txt or .log in the given folder are removed.
        Subfolders (e.g. arch\) are not touched. Errors are suppressed silently.
    .PARAMETER LogPath
        Folder to clean up. Returns silently if the folder does not exist yet.
    .PARAMETER RetentionDays
        Maximum age of log files in days. Files with LastWriteTime older than
        (today - RetentionDays) are deleted.
    .EXAMPLE
        Remove-DecosterOutputOldLogs -LogPath $GLB_logPath -RetentionDays $Config.Logging.LogRetentionDays
    #>
    param(
        [Parameter(Mandatory)][string]$LogPath,
        [Parameter(Mandatory)][int]$RetentionDays
    )
    if (-not (Test-Path -LiteralPath $LogPath)) { return }
    $cutoff = (Get-Date).AddDays(-$RetentionDays)
    Get-ChildItem -Path $LogPath -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in '.txt', '.log' -and $_.LastWriteTime -lt $cutoff } |
        ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
}

function Write-DecosterOutputLog {
    <#
    .SYNOPSIS
        Writes the contents of a ListBox to a timestamped log file.
    .DESCRIPTION
        Implements the set-log pattern shared across the DecosterOutput scripts.
        - Creates the log directory if it does not exist.
        - Appends a _1, _2, … counter suffix if the target filename already exists.
    .PARAMETER Listbox
        The ListBox whose Items will be written to the log.
    .PARAMETER LogPath
        Folder path where the log file will be created.
        Typically: Join-Path $Config.Logging.BasePath "SubfolderName"
    .PARAMETER Username
        Inserted in the filename: <username>_<yyyyMMdd_HHmmss>_<Suffix>.txt
    .PARAMETER Suffix
        Descriptive label for the run type.
        Examples: "ServerDecom", "DHCPreservation", "ServerCreation", "CustomisationPasswordChange"
    .EXAMPLE
        Write-DecosterOutputLog -Listbox $textBoxOutPut -LogPath $GLB_logPath -Username $GLB_currentUsername -Suffix "ServerDecom"
    #>
    param(
        [Parameter(Mandatory)][System.Windows.Forms.ListBox]$Listbox,
        [Parameter(Mandatory)][string]$LogPath,
        [Parameter(Mandatory)][string]$Username,
        [string]$Suffix = "run"
    )
    $dateStr  = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $baseName = "${Username}_${dateStr}_${Suffix}"
    $fileName = "$baseName.txt"
    $filePath = Join-Path -Path $LogPath -ChildPath $fileName

    if (-not (Test-Path -Path $LogPath)) {
        New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
    }

    if (Test-Path -Path $filePath) {
        $counter = 1
        while (Test-Path -Path (Join-Path -Path $LogPath -ChildPath "${baseName}_${counter}.txt")) {
            $counter++
        }
        $filePath = Join-Path -Path $LogPath -ChildPath "${baseName}_${counter}.txt"
    }

    $content = $Listbox.Items | Out-String
    Add-Content -Path $filePath -Value $content
}
#endregion

function Test-DecosterOutputConfig {
    <#
    .SYNOPSIS
        Validates that all required Section.Key paths exist and are non-empty in a loaded config hashtable.
    .DESCRIPTION
        Call this after Import-PowerShellDataFile succeeds to verify the config contains all keys
        expected by the script. Returns a list of missing or empty keys so the caller can show a
        meaningful error in the GUI.
    .PARAMETER Config
        The hashtable returned by Import-PowerShellDataFile.
    .PARAMETER RequiredKeys
        Array of "Section.Key" strings, e.g. @('Logging.BasePath','Paths.ExeBasePath').
    .OUTPUTS
        [string[]] — list of missing or empty keys. Empty array means the config is valid.
    .EXAMPLE
        $warnings = Test-DecosterOutputConfig -Config $GLB_config `
                        -RequiredKeys @('Logging.BasePath','Logging.LogRetentionDays','Paths.ExeBasePath')
        if ($warnings.Count -gt 0) { Write-Warning "Missing config keys: $($warnings -join ', ')" }
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string[]]$RequiredKeys
    )
    $missing = [System.Collections.Generic.List[string]]::new()
    foreach ($key in $RequiredKeys) {
        $parts   = $key -split '\.', 2
        $section = $parts[0]
        $name    = $parts[1]
        if (-not $Config.ContainsKey($section) -or
            -not $Config[$section].ContainsKey($name) -or
            [string]::IsNullOrWhiteSpace([string]$Config[$section][$name])) {
            $missing.Add($key)
        }
    }
    return $missing.ToArray()
}

Export-ModuleMember -Function `
    Remove-DecosterOutputOldLogs,
    Format-DecosterOutputHeader,
    Format-DecosterOutputOk,
    Format-DecosterOutputFail,
    Format-DecosterOutputSkip,
    Format-DecosterOutputWarn,
    Format-DecosterOutputWarnResult,
    Format-DecosterOutputError,
    Format-DecosterOutputInfo,
    Format-DecosterOutputConnected,
    Format-DecosterOutputConnectionFail,
    Format-DecosterOutputSummary,
    Format-DecosterOutputTotal,
    Format-DecosterOutputSeparator,
    Format-DecosterOutputDivider,
    Format-DecosterOutputWarningBlock,
    Format-DecosterOutputFailureBlock,
    Format-DecosterOutputKeyValue,
    Format-DecosterOutputStep,
    Format-DecosterOutputSubItem,
    Test-DecosterOutputConfig,
    Add-DecosterOutputOutput,
    Set-DecosterOutputProgress,
    Set-DecosterOutputCredentialStatus,
    Set-DecosterOutputLabelStatus,
    Write-DecosterOutputLog
