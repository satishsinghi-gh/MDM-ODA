# -- param must be first statement for PS5 compatibility --
param([switch]$SkipUpgrade)

# ============================================================
# MDM-ODA - Orchestrator Launcher
# Author      : Satish Singhi
# Description : PS5-compatible launcher. Installs/upgrades
#               PowerShell 7 via winget (user scope, no admin)
#               then relaunches self in PS7 to show the GUI.
#
# Distribute alongside: Orchestrator_GUI.ps1
# Users run this file only. Orchestrator_GUI.ps1 is copied
# to a staging directory and launched via PS7.
#
# -- DEVELOPER CONFIG -----------------------------------------
# ============================================================

# ============================================================
#region ORCHESTRATOR CONFIG
# ============================================================
$OrcConfig = @{
    StagingRoot  = "C:\Logs\MDM-ODA"
    StagingPath  = "C:\Logs\MDM-ODA\Staging"
    LogPath      = "C:\Logs\MDM-ODA\Logs"
    ToolFileName = "MDM-ODA.ps1"
}
# Ensure all folders exist upfront
foreach ($dir in @($OrcConfig.StagingRoot, $OrcConfig.StagingPath, $OrcConfig.LogPath)) {
    if (-not (Test-Path $dir)) {
        try { New-Item -ItemType Directory -Path $dir -Force | Out-Null } catch {}
    }
}
#endregion ORCHESTRATOR CONFIG
# ============================================================

# ============================================================
# PS5 SECTION  -  Everything below here runs in PS5.
# PS7 code begins after the exit statement further down.
# PS5 exits before reaching the PS7 section so PS7-only
# syntax there never causes a parse error in PS5.
# ============================================================

if ($PSVersionTable.PSVersion.Major -lt 7) {

    Write-Host ""
    Write-Host "  MDM ODA Tool" -ForegroundColor Cyan
    Write-Host "  Checking PowerShell 7..." -ForegroundColor Cyan
    Write-Host ""

    # -- Verify winget is available -------------------------------------------
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        Write-Host "  winget is not available on this device." -ForegroundColor Red
        Write-Host "  Please install PowerShell 7 from:" -ForegroundColor Yellow
        Write-Host "  https://github.com/PowerShell/PowerShell/releases/latest" -ForegroundColor Cyan
        Write-Host ""
        Read-Host "  Press Enter to exit"
        exit 1
    }

    # -- PS7 detection: validate real executable (not a Windows Store stub) -----
    Write-Host "  Checking PowerShell 7..." -ForegroundColor Gray

    # Windows 11/10 places a tiny stub at WindowsApps\pwsh.exe even when PS7 is NOT
    # installed — Test-Path returns $true but running it just opens the Store.
    # Filter stubs out by requiring file size > 50 KB (real pwsh.exe is ~200 KB).
    function Test-RealPwsh {
        param([string]$Path)
        if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
        if (-not (Test-Path $Path -ErrorAction SilentlyContinue)) { return $false }
        $item = Get-Item $Path -ErrorAction SilentlyContinue
        if (-not $item) { return $false }
        if ($item.Length -lt 51200) {
            Write-Host "  Skipping stub at: $Path (size $($item.Length) bytes)" -ForegroundColor DarkGray
            return $false
        }
        return $true
    }

    # Probe in priority order:
    #   1. System-wide MSI  (most reliable on corporate devices)
    #   2. User-scope winget non-Store
    #   3. PATH (covers custom/Chocolatey/SCCM installs)
    #   4. WindowsApps  (Store install — validated last because stubs live here)
    $probePaths = [System.Collections.Generic.List[string]]::new()
    $probePaths.Add('C:\Program Files\PowerShell\7\pwsh.exe')                              # system MSI
    $probePaths.Add((Join-Path $env:LOCALAPPDATA 'Microsoft\PowerShell\7\pwsh.exe'))       # winget user-scope
    $pwshCmd = Get-Command pwsh -All -ErrorAction SilentlyContinue
    if ($pwshCmd) {
        foreach ($p in ($pwshCmd | Where-Object { $_.Source -notlike '*WindowsApps*' })) {
            if (-not $probePaths.Contains($p.Source)) { $probePaths.Add($p.Source) }
        }
    }
    $probePaths.Add((Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\pwsh.exe'))        # Store (checked LAST; stubs are filtered by size)

    $pwsh7Detected = $null
    foreach ($p in $probePaths) {
        if (Test-RealPwsh $p) { $pwsh7Detected = $p; break }
    }
    $ps7Installed = $null -ne $pwsh7Detected

    if (-not $ps7Installed) {
        Write-Host "  PowerShell 7 not found. Installing (user-scope, no admin required)..." -ForegroundColor Yellow
        Write-Host "  This may take 1-2 minutes. Please wait." -ForegroundColor Gray
        & winget install --id Microsoft.PowerShell --exact --scope user --source winget `
            --accept-source-agreements --accept-package-agreements --silent
        Start-Sleep -Seconds 3
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","User") + ";" +
                    [System.Environment]::GetEnvironmentVariable("Path","Machine")
        # Re-probe after install (using same stub-safe validator)
        foreach ($p in $probePaths) {
            if (Test-RealPwsh $p) { $pwsh7Detected = $p; break }
        }
        $ps7Installed = $null -ne $pwsh7Detected
        if ($ps7Installed) {
            Write-Host "  PowerShell 7 installed at: $pwsh7Detected" -ForegroundColor Green
        } else {
            Write-Host "  Installation may have succeeded but path not yet visible." -ForegroundColor Yellow
            Write-Host "  Please close this window and open a new terminal, then run the script again." -ForegroundColor Yellow
            Read-Host "  Press Enter to exit"
            exit 1
        }
    } else {
        $scopeLabel = if ($pwsh7Detected -like "*LocalAppData*") { "(user-scope)" } else { "(system-wide)" }
        Write-Host "  PowerShell 7 detected $scopeLabel : $pwsh7Detected" -ForegroundColor Green
    }

    # -- Upgrade to latest stable every launch (skip previews) -------------------
    if (-not $SkipUpgrade) {
        Write-Host "  Checking for updates (stable only)..." -ForegroundColor Gray

        $showOutput = & winget show --id Microsoft.PowerShell --exact --source winget 2>&1
        $verLine    = $showOutput | Where-Object { $_ -match '^\s*Version\s*:' } | Select-Object -First 1

        $latestStable = $null
        if ($verLine -and $verLine -match ':\s*([\d\.]+)') {
            $rawVer = $Matches[1]
            if ($rawVer -notmatch 'preview|rc|beta|alpha') { $latestStable = $rawVer }
        }

        if ($latestStable) {
            Write-Host "  Latest stable PS7: $latestStable" -ForegroundColor Gray
            $upgradeOut = & winget upgrade --id Microsoft.PowerShell --exact --scope user `
                --source winget --accept-source-agreements --accept-package-agreements --silent 2>&1
            $done = $upgradeOut | Where-Object {
                $_ -match 'No applicable upgrade' -or
                $_ -match 'already installed'     -or
                $_ -match 'Successfully installed'
            }
            if ($done) {
                Write-Host "  PowerShell 7 is up to date (v$latestStable)." -ForegroundColor Green
            } else {
                Write-Host "  Update applied." -ForegroundColor Green
            }
        }
    }

    # -- Find pwsh.exe and relaunch in PS7 ---------------------------------------
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","User") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path","Machine")

    # Use already-detected path; fall back to re-probe if needed
    $pwsh7 = $pwsh7Detected
    if (-not $pwsh7) {
        $pwshAll = Get-Command pwsh -All -ErrorAction SilentlyContinue
        if ($pwshAll) {
            foreach ($p in ($pwshAll | Where-Object { $_.Source -notlike '*WindowsApps*' })) {
                if (Test-Path $p.Source -ErrorAction SilentlyContinue) { $pwsh7 = $p.Source; break }
            }
        }
    }

    if ($pwsh7) {
        Write-Host "  Staging and launching in PowerShell 7 ($pwsh7)..." -ForegroundColor Green
        # Hide this console window before handing off
        try {
            Add-Type -TypeDefinition @"
using System.Runtime.InteropServices;
public class OrcLaunchConsole {
    [DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);
    [DllImport("kernel32.dll")] public static extern System.IntPtr GetConsoleWindow();
}
"@ -ErrorAction SilentlyContinue
            $null = [OrcLaunchConsole]::ShowWindow([OrcLaunchConsole]::GetConsoleWindow(), 0)
        } catch {}

        # Extract PS7 payload from the block-comment section embedded in this script.
        # PS5 never compiled that code — it was just a comment token to the PS5 tokenizer.
        $scriptText  = [System.IO.File]::ReadAllText($PSCommandPath, [System.Text.Encoding]::UTF8)
        $startTag    = '<#__PS7' + '_PAYLOAD_START__'
        $endTag      = '__PS7_PAYLOAD' + '_END__#>'
        $startOffset = $scriptText.IndexOf($startTag)
        $endOffset   = $scriptText.IndexOf($endTag)

        if ($startOffset -lt 0 -or $endOffset -le $startOffset) {
            Write-Host "  ERROR: PS7 payload boundary markers not found in script." -ForegroundColor Red
            Read-Host "  Press Enter to exit"
            exit 1
        }

        $payloadText   = $scriptText.Substring($startOffset + $startTag.Length, $endOffset - ($startOffset + $startTag.Length)).TrimStart("`r", "`n")
        $stagingScript = Join-Path $OrcConfig.StagingPath $OrcConfig.ToolFileName

        # Ensure staging folders exist
        foreach ($dir in @($OrcConfig.StagingRoot, $OrcConfig.StagingPath, $OrcConfig.LogPath)) {
            if (-not (Test-Path $dir)) {
                try { New-Item -ItemType Directory -Path $dir -Force | Out-Null } catch {}
            }
        }

        # Write PS7 payload to staging directory as MDM-ODA.ps1
        [System.IO.File]::WriteAllText($stagingScript, $payloadText, [System.Text.Encoding]::UTF8)
        Write-Host "  Staged: $stagingScript" -ForegroundColor Green

        # Launch staged PS7 script — hidden console window, wait for exit
        Start-Process -FilePath $pwsh7 `
            -ArgumentList "-ExecutionPolicy Bypass -STA -File `"$stagingScript`"" `
            -WindowStyle Hidden `
            -Wait
        exit 0
    } else {
        Write-Host ""
        Write-Host "  Could not locate pwsh.exe." -ForegroundColor Red
        Write-Host "  Please close this window, open a new terminal and run the script again." -ForegroundColor Yellow
        Write-Host ""
        Read-Host "  Press Enter to exit"
        exit 1
    }
}

# ============================================================
# PS7 PAYLOAD — embedded as a PS5 block comment so PS5 never
#   compiles this code. The PS5 launcher reads this file at
#   runtime, extracts the payload between the sentinel tags,
#   stages it as MDM-ODA.ps1, and launches it with pwsh.exe.
#   DO NOT edit or remove the sentinel tags below.
# ============================================================
<#__PS7_PAYLOAD_START__
# ── Orchestrator config ─────────────────────────────────────────────────────
# Defined here so MDM-ODA.ps1 is self-contained when staged and run directly.
# Values must match the $OrcConfig in the PS5 launcher section above.
$OrcConfig = @{
    StagingRoot  = 'C:\Logs\MDM-ODA'
    StagingPath  = 'C:\Logs\MDM-ODA\Staging'
    LogPath      = 'C:\Logs\MDM-ODA\Logs'
    ToolFileName = 'MDM-ODA.ps1'
}

# ── Startup diagnostic log (captures fatal errors before the form appears) ─
$script:OrcDiagLog = Join-Path $OrcConfig.LogPath ("OrcDiag_" + (Get-Date -f 'yyyyMMdd_HHmmss') + ".log")
function Write-OrcDiag { param([string]$Msg)
    try { [System.IO.Directory]::CreateDirectory($OrcConfig.LogPath)|Out-Null
          [System.IO.File]::AppendAllText($script:OrcDiagLog,"[$(Get-Date -f 'HH:mm:ss')] $Msg`n",[System.Text.Encoding]::UTF8) } catch {} }
trap { Write-OrcDiag "FATAL: $($_.Exception.Message)`n$($_.ScriptStackTrace)"; break }
Write-OrcDiag "MDM-ODA.ps1 started. PSCommandPath=$PSCommandPath"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -- Read WindowTitle from the embedded GUI section ----------------------------
$toolTitle = "Launching..."
try {
    $selfContent = [System.IO.File]::ReadAllText($PSCommandPath, [System.Text.Encoding]::UTF8)
    if ($selfContent -match 'WindowTitle\s*=\s*"([^"]+)"') { $toolTitle = $Matches[1] }
} catch {}

# -- Colours ------------------------------------------------------------------
$C = @{
    Header   = [System.Drawing.ColorTranslator]::FromHtml('#007A00')
    HeaderFg = [System.Drawing.Color]::White
    BtnGreen = [System.Drawing.ColorTranslator]::FromHtml('#007A00')
    BtnFg    = [System.Drawing.Color]::White
    Bg       = [System.Drawing.ColorTranslator]::FromHtml('#F8FBF8')
    LogBg    = [System.Drawing.ColorTranslator]::FromHtml('#1A1F2E')
    LogFg    = [System.Drawing.ColorTranslator]::FromHtml('#9DB8D8')
    Success  = [System.Drawing.ColorTranslator]::FromHtml('#5DC98B')
    Warning  = [System.Drawing.ColorTranslator]::FromHtml('#F5A623')
    Error    = [System.Drawing.ColorTranslator]::FromHtml('#F06C6C')
    Info     = [System.Drawing.ColorTranslator]::FromHtml('#C8DEF5')
}

# -- Build form ----------------------------------------------------------------
$form                 = [System.Windows.Forms.Form]::new()
$form.Text            = $toolTitle
$form.Width           = 580
$form.Height          = 380
$form.MinimumSize     = [System.Drawing.Size]::new(480, 300)
$form.StartPosition   = 'CenterScreen'
$form.BackColor       = $C.Bg
$form.Font            = [System.Drawing.Font]::new('Segoe UI', 10)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false

$header              = [System.Windows.Forms.Panel]::new()
$header.Dock         = 'Top'
$header.Height       = 52
$header.BackColor    = $C.Header
$headerLbl           = [System.Windows.Forms.Label]::new()
$headerLbl.Text      = $toolTitle
$headerLbl.ForeColor = $C.HeaderFg
$headerLbl.Font      = [System.Drawing.Font]::new('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
$headerLbl.AutoSize  = $true
$headerLbl.Location  = [System.Drawing.Point]::new(16, 13)
$header.Controls.Add($headerLbl)

$content              = [System.Windows.Forms.TableLayoutPanel]::new()
$content.Dock         = 'Fill'
$content.Padding      = [System.Windows.Forms.Padding]::new(14, 10, 14, 10)
$content.BackColor    = $C.Bg
$content.ColumnCount  = 1
$content.RowCount     = 3
$null = $content.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
$null = $content.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 30))
$null = $content.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
$null = $content.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 44))
$form.Controls.Add($content)
$form.Controls.Add($header)

$statusLbl           = [System.Windows.Forms.Label]::new()
$statusLbl.Text      = 'Initialising...'
$statusLbl.Dock      = 'Fill'
$statusLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$statusLbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#1A1F2E')
$content.Controls.Add($statusLbl, 0, 0)

$logBox              = [System.Windows.Forms.RichTextBox]::new()
$logBox.ReadOnly     = $true
$logBox.BackColor    = $C.LogBg
$logBox.ForeColor    = $C.LogFg
$logBox.Font         = [System.Drawing.Font]::new('Consolas', 9)
$logBox.BorderStyle  = 'None'
$logBox.Dock         = 'Fill'
$content.Controls.Add($logBox, 0, 1)

$btnPanel            = [System.Windows.Forms.Panel]::new()
$btnPanel.Dock       = 'Fill'
$btnPanel.BackColor  = $C.Bg
$content.Controls.Add($btnPanel, 0, 2)

$btnClose            = [System.Windows.Forms.Button]::new()
$btnClose.Text       = 'Close'
$btnClose.Width      = 100
$btnClose.Height     = 32
$btnClose.Anchor     = 'Right,Bottom'
$btnClose.Location   = [System.Drawing.Point]::new(440, 6)
$btnClose.BackColor  = $C.BtnGreen
$btnClose.ForeColor  = $C.BtnFg
$btnClose.FlatStyle  = 'Flat'
$btnClose.FlatAppearance.BorderSize = 0
$btnClose.Visible    = $false
$btnClose.Add_Click({ $form.Close() })
$btnPanel.Controls.Add($btnClose)

function OLog {
    param([string]$Line, [string]$Level = 'Info')
    $col = switch ($Level) {
        'Success' { $C.Success } 'Warning' { $C.Warning }
        'Error'   { $C.Error   } 'Action'  { $C.Info    }
        default   { $C.LogFg }
    }
    $logBox.SelectionStart  = $logBox.TextLength
    $logBox.SelectionLength = 0
    $logBox.SelectionColor  = $col
    $logBox.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] $Line`n")
    $logBox.ScrollToCaret()
    $form.Refresh()
}
function OStatus { param([string]$T) $statusLbl.Text = $T; $form.Refresh() }

$form.Add_Shown({

    # -- Extract GUI tool content FIRST (before clearing staging) ----------------
    # MDM-ODA.ps1 runs FROM the staging directory — read it before deleting it.
    OStatus 'Extracting GUI tool...'
    $toolPath  = Join-Path $OrcConfig.StagingPath $OrcConfig.ToolFileName
    $beginMark = '#' + '=' * 10 + ' BEGIN_GUI_TOOL_CONTENT ' + '=' * 10
    $endMark   = '#' + '=' * 10 + ' END_GUI_TOOL_CONTENT '   + '=' * 11
    $guiLines  = $null
    try {
        $selfLines = [System.IO.File]::ReadAllLines($PSCommandPath, [System.Text.Encoding]::UTF8)
        $inBlock   = $false
        $guiLines  = [System.Collections.Generic.List[string]]::new()
        foreach ($line in $selfLines) {
            if ($line -eq $beginMark) { $inBlock = $true;  continue }
            if ($line -eq $endMark)   { $inBlock = $false; break }
            if ($inBlock)              { $guiLines.Add($line) }
        }
        if ($guiLines.Count -eq 0) { throw 'GUI tool section not found in orchestrator.' }
        OLog ('GUI tool read (' + $guiLines.Count + ' lines).') 'Success'
        Write-OrcDiag "GUI tool read: $($guiLines.Count) lines"
    } catch {
        OLog "Extraction failed: $($_.Exception.Message)" 'Error'
        OStatus 'Error - see log.'
        Write-OrcDiag "Extraction failed: $($_.Exception.Message)"
        $btnClose.Visible = $true; return
    }

    # -- Prepare staging (AFTER reading self — Remove-Item deletes PSCommandPath) --
    OStatus 'Preparing staging...'
    OLog "Staging: $($OrcConfig.StagingPath)" 'Info'
    try {
        if (Test-Path $OrcConfig.StagingPath) {
            Remove-Item $OrcConfig.StagingPath -Recurse -Force -ErrorAction Stop
        }
        New-Item -ItemType Directory -Path $OrcConfig.StagingPath -Force | Out-Null
        OLog 'Staging ready.' 'Success'
    } catch {
        OLog "Staging failed: $($_.Exception.Message)" 'Error'
        OStatus 'Error - see log.'
        $btnClose.Visible = $true; return
    }

    # -- Write GUI tool to staging ---------------------------------------------
    try {
        [System.IO.File]::WriteAllText($toolPath, ($guiLines -join "`n"), [System.Text.Encoding]::UTF8)
        OLog ('GUI tool staged (' + $guiLines.Count + ' lines).') 'Success'
        Write-OrcDiag "GUI tool written to: $toolPath"
    } catch {
        OLog "Write to staging failed: $($_.Exception.Message)" 'Error'
        OStatus 'Error - see log.'
        $btnClose.Visible = $true; return
    }


    # -- Launch GUI tool -------------------------------------------------------
    OStatus "Launching $toolTitle..."
    $pwsh7 = (Get-Process -Id $PID).MainModule.FileName
    OLog "Launching via: $pwsh7" 'Action'

    # Write orchestrator launch log for diagnostics
    $orcLogPath = Join-Path $OrcConfig.LogPath ("OrcLaunch_{0}.log" -f (Get-Date -f 'yyyyMMdd_HHmmss'))
    Set-Content -Path $orcLogPath -Value "Orchestrator Launch Log" -Encoding UTF8
    Add-Content -Path $orcLogPath -Value "========================" -Encoding UTF8
    Add-Content -Path $orcLogPath -Value ("Timestamp  : " + (Get-Date)) -Encoding UTF8
    Add-Content -Path $orcLogPath -Value ("Host       : " + $env:COMPUTERNAME) -Encoding UTF8
    Add-Content -Path $orcLogPath -Value ("User       : " + $env:USERNAME) -Encoding UTF8
    Add-Content -Path $orcLogPath -Value ("PS7 Path   : " + $pwsh7) -Encoding UTF8
    Add-Content -Path $orcLogPath -Value ("Tool Path  : " + $toolPath) -Encoding UTF8
    Add-Content -Path $orcLogPath -Value ("Tool Lines : " + $guiLines.Count) -Encoding UTF8
    Add-Content -Path $orcLogPath -Value ("Staging    : " + $OrcConfig.StagingPath) -Encoding UTF8
    Add-Content -Path $orcLogPath -Value ("Log Dir    : " + $OrcConfig.LogPath) -Encoding UTF8
    OLog "Launch log: $orcLogPath" 'Info'

    # Transcript path for the GUI tool startup diagnostics
    $transcriptPath = Join-Path $OrcConfig.LogPath ("MDMODA_Transcript_{0}.log" -f (Get-Date -f 'yyyyMMdd_HHmmss'))

    # Pass log paths to GUI tool via environment variables.
    # Note: param() blocks in mid-file cause PS5 parse errors, so env vars are used.
    $env:MDMODA_LogFolder  = $OrcConfig.LogPath
    $env:MDMODA_Transcript = $transcriptPath

    $launchArgs = "-ExecutionPolicy Bypass -STA -File `"$toolPath`""
    try {
        # WindowStyle Normal (visible) so startup crashes surface in a console window.
        # Switch to Hidden once the tool is confirmed stable.
        $proc = Start-Process $pwsh7 -ArgumentList $launchArgs -WindowStyle Hidden -PassThru
        Add-Content -Path $orcLogPath -Value ("Process ID : " + $proc.Id) -Encoding UTF8
        OLog ("Tool launched. PID=" + $proc.Id + " Transcript=" + $transcriptPath) 'Success'
        OStatus ($toolTitle + " launched.")
        # Hide orchestrator console window now that the GUI tool is running
        try {
            if (-not ([System.Management.Automation.PSTypeName]'OrcConsole').Type) {
                Add-Type -TypeDefinition @"
using System.Runtime.InteropServices;
public class OrcConsole {
    [DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);
    [DllImport("kernel32.dll")] public static extern System.IntPtr GetConsoleWindow();
}
"@
            }
            $null = [OrcConsole]::ShowWindow([OrcConsole]::GetConsoleWindow(), 0)
        } catch {}
        Start-Sleep -Milliseconds 800
        $form.Close()
    } catch {
        $errMsg = "Launch failed: " + $_.Exception.Message
        Add-Content -Path $orcLogPath -Value ("ERROR: " + $errMsg) -Encoding UTF8
        OLog $errMsg 'Error'
        OStatus 'Launch failed. Check log.'
        $btnClose.Visible = $true
    }
})

$form.ShowDialog() | Out-Null
exit 0

#========== BEGIN_GUI_TOOL_CONTENT ==========
# Note: #Requires omitted - orchestrator guarantees PS7
# Module requirements checked at runtime by the prereq window (see #region PREREQ CHECK)
# Values passed from orchestrator via environment variables (no param block  - 
# a param() in the middle of the orchestrator file causes PS5 parse errors)
$OverrideLogFolder = if ($env:MDMODA_LogFolder)    { $env:MDMODA_LogFolder }    else { "" }
$TranscriptPath    = if ($env:MDMODA_Transcript)   { $env:MDMODA_Transcript }   else { "" }

# -- Early-crash diagnostic: written to TEMP before DevConfig initialised --
$_earlyLog = Join-Path $env:TEMP ("MDM-ODA_Early_{0}.log" -f (Get-Date -f 'yyyyMMdd_HHmmss'))
function Write-EarlyLog { param([string]$Msg) try { Add-Content -Path $_earlyLog -Value "[$(Get-Date -f 'HH:mm:ss')] $Msg" -Encoding UTF8 } catch {} }
Write-EarlyLog "GUI tool started. PS=$($PSVersionTable.PSVersion) User=$($env:USERNAME)"


# .SYNOPSIS
    # MDM On-Demand Actions (MDM-ODA) - PowerShell 7 & WPF tool for Entra ID and Intune on-demand operations.
#
# .DESCRIPTION
    # A single-file PowerShell 7+ WPF GUI application for Entra ID and Intune on-demand
    # operations with an embedded WPF interface. Features include:
      # - Group Management: Search, list members, create, rename, compare, bulk owners,
      #   add user devices, find common/distinct, dynamic membership, object membership
      # - Device Management: Device info, Intune policy assignments across all policy types
      # - Productivity: Session notes, keyword filter, verbose logging, XLSX export
      # - Security: OAuth 2.0 delegated auth, read-only API scopes by default,
      #   validation-before-commit for all write operations, least privilege enforcement
    # All write operations require mandatory Entra validation + user confirmation before execution.
    # Real-time verbose logging in the lower pane, simultaneously written to a local log file.
#
# .NOTES
    # Author      : Satish Singhi
    # Version     : 0.7
    # Requires    : PowerShell 7+, Microsoft.Graph SDK
    # Auth        : Azure App Registration - Delegated permissions (interactive browser)
                  # Connect-MgGraph is called WITHOUT -Scopes parameter
    # Permissions : User.Read, User.Read.All, Group.Read.All, GroupMember.Read.All,
                   # Directory.Read.All, Device.Read.All, DeviceManagementConfiguration.Read.All,
                   # DeviceManagementManagedDevices.Read.All, DeviceManagementRBAC.Read.All, offline_access
#
    # ===================== DEVELOPER CONFIG  -  EDIT THIS SECTION =====================
    # All developer-facing settings are in the #region DEVELOPER CONFIG block below.
    # This includes: theme colours, background image, logo, log path, window size, etc.
    # ================================================================================


Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
#region DEVELOPER CONFIG
# ============================================================

$DevConfig = [ordered]@{

    # ---- Azure App Registration ----
    TenantId = ""
    ClientId = ""

    # ---- Local log file path ----
    # A timestamped log file will be created here each session
    LogFolder = "C:\Logs\MDM-ODA\Logs"

    # ---- Window ----
    WindowTitle  = "MDM On-Demand Actions  v0.7"
    WindowWidth  = 1100
    WindowHeight = 780
    MinWidth     = 900
    MinHeight    = 600

    # ---- Verbose pane ----
    VerbosePaneDefaultHeight = 200   # pixels; user can drag the splitter

    # ---- Background image ----
    # Local file path takes priority over URL. Set both to "" to use solid colour.
    BackgroundImagePath    = ""      # e.g. "C:\Assets\bg.jpg"  (local file)
    BackgroundImageUrl     = ""      # e.g. "https://intranet/assets/bg.jpg"  (web URL)
    BackgroundImageOpacity = 0.08   # 0.0 (invisible) to 1.0 (fully opaque)

    # ---- Logo ----
    # Local file path takes priority over URL. Set both to "" to hide the logo.
    LogoImagePath   = ""            # e.g. "C:\Assets\logo.png"  (local file)
    LogoImageUrl    = ""            # e.g. "https://intranet/assets/logo.png"  (web URL)
    LogoWidth       = 240
    LogoHeight      = 80
    LogoBase64      = ""            # Base64-encoded image (PNG/JPG/ICO). Takes priority over Path/URL.
    LogoTitleSpacing = 12              # Pixels between logo and title text
    # To encode: [Convert]::ToBase64String([IO.File]::ReadAllBytes("C:\logo.png"))
    # To encode from URL: [Convert]::ToBase64String((Invoke-WebRequest "https://...").Content)

    # ---- Colour theme ----
    # Primary colour used for header bar, buttons, accents
    ColorPrimary        = "#007A00"   # Primary green
    ColorPrimaryLight   = "#009900"   # Primary green (hover)
    ColorPrimaryDark    = "#006400"   # Primary green (pressed)
    ColorAccent         = "#007A00"   # Accent green (unified with primary)
    ColorAccentLight    = "#009900"
    ColorBackground     = "#F8FBF8"   # Page background
    ColorSurface        = "#FFFFFF"   # Card / panel background
    ColorBorder         = "#B0D0B0"
    ColorText           = "#1A1F2E"
    ColorTextMuted      = "#5A6480"
    ColorSuccess        = "#1E7A4A"
    ColorWarning        = "#B45B00"
    ColorError          = "#C0392B"
    ColorInfo           = "#1A5FA8"

    # ---- Verbose pane log colours (foreground hex) ----
    LogColorInfo    = "#9DB8D8"
    LogColorSuccess = "#5DC98B"
    LogColorWarning = "#F5A623"
    LogColorError   = "#F06C6C"
    LogColorAction  = "#C8DEF5"

    # ---- Font ----
    FontFamily   = "Segoe UI"
    FontSizeBase = 13

    # ---- Max objects per batch Graph request ----
    GraphBatchSize = 20

    # ---- PIM role auto-refresh interval (minutes). Set to 0 to disable.
    PimRefreshIntervalMinutes = 5

    # ---- Feedback URL (opened when Feedback button is clicked) ----
    # Set to the URL of your feedback form, Teams channel, or any web page.
    # Leave empty to hide the Feedback button.
    FeedbackUrl = ""              # e.g. "https://forms.office.com/yourformid"

    # ---- M365 Group Mail Domain ----
    # Domain suffix shown next to the mail nickname field when creating M365 groups.
    # e.g. "contoso.onmicrosoft.com" or "contoso.com"
    # Leave empty to show only "@" (tenant assigns the domain automatically).
    M365MailDomain = ""           # e.g. "contoso.onmicrosoft.com"
}

#endregion DEVELOPER CONFIG
# ============================================================


# ============================================================
#region BOOTSTRAP  -  STA check & assemblies
# ============================================================

# ── Start transcript immediately for startup diagnostics ─────────────────────
$script:TranscriptStarted = $false
try {
    $tsPath = if ($TranscriptPath -and $TranscriptPath -ne "") {
        $TranscriptPath
    } else {
        Join-Path $env:TEMP ("MDMODA_Transcript_{0}.log" -f (Get-Date -f "yyyyMMdd_HHmmss"))
    }
    $tsDir = Split-Path $tsPath -Parent
    if (-not (Test-Path $tsDir)) { New-Item -ItemType Directory -Path $tsDir -Force | Out-Null }
    Start-Transcript -Path $tsPath -Append -ErrorAction Stop
    $script:TranscriptStarted = $true
    Write-Host "[STARTUP] Transcript: $tsPath"
} catch {
    Write-Host "[STARTUP] Transcript failed: $($_.Exception.Message)"
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ── WPF Application singleton ─────────────────────────────────────────────────
# Required for transparent/custom-chrome windows (WindowStyle=None).
# ShutdownMode=OnExplicitShutdown keeps the dispatcher alive across
# multiple windows (prereq -> main). Each window uses DispatcherFrame
# to pump its own message loop without calling Application.Run().
if (-not [System.Windows.Application]::Current) {
    $script:WpfApp = [System.Windows.Application]::new()
    $script:WpfApp.ShutdownMode = [System.Windows.ShutdownMode]::OnExplicitShutdown
} else {
    [System.Windows.Application]::Current.ShutdownMode = [System.Windows.ShutdownMode]::OnExplicitShutdown
}

# ── DPI awareness  -  must be set before any WPF window is created ──
# Per-Monitor v2 tells WPF to render at native DPI on each monitor,
# preventing the window from being clipped or mis-sized on high-DPI displays.
try {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class DpiAware {
    [DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr value);
    // DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
    public static void SetPerMonitorV2() { SetProcessDpiAwarenessContext(new IntPtr(-4)); }
}
'@ -ErrorAction SilentlyContinue
    [DpiAware]::SetPerMonitorV2()
} catch {}

# Ensure we are on an STA thread (required for WPF)
# The orchestrator launcher guarantees -STA; this is a safety net for direct launches.
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Write-Error "This script must run on an STA thread. Use the orchestrator launcher, or run: pwsh -STA -File '$PSCommandPath'"
    exit 1
}

# Override LogFolder from orchestrator if provided
if (-not [string]::IsNullOrWhiteSpace($OverrideLogFolder)) {
    $DevConfig.LogFolder = $OverrideLogFolder
    Write-Host "[STARTUP] LogFolder overridden to: $OverrideLogFolder"
}

# Ensure log folder exists
if (-not [string]::IsNullOrWhiteSpace($DevConfig.LogFolder)) {
    if (-not (Test-Path $DevConfig.LogFolder)) {
        New-Item -ItemType Directory -Path $DevConfig.LogFolder -Force | Out-Null
    }
}

$script:SessionStamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$script:LogFile = if (-not [string]::IsNullOrWhiteSpace($DevConfig.LogFolder)) {
    Join-Path $DevConfig.LogFolder ("MDM-ODA_{0}.log" -f $script:SessionStamp)
} else { $null }

#endregion


# ============================================================
#region PREREQ CHECK WINDOW
# ============================================================

function Start-PrereqCheck {

    # Prereq window  -  runs checks on a background runspace.
    # Results are posted to a ConcurrentQueue drained by a DispatcherTimer.
    # No STA thread blocking. Find-Module (network) runs in the same runspace.


    $requiredModules = @(
        [PSCustomObject]@{ Name = 'Microsoft.Graph.Authentication';               Label = 'Graph . Authentication' }
    )
    $optionalModules = @(
        [PSCustomObject]@{ Name = 'ImportExcel'; Label = 'ImportExcel (XLSX export)' }
    )

    # ── Shared queue between runspace and UI ──
    Write-Host "[PREREQ] Entering Start-PrereqCheck"
    $queue = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
    Write-Host "[PREREQ] Queue created"

    # ── Build XAML (simple Grid rows  -  no data binding) ──
    [xml]$prereqXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$($DevConfig.WindowTitle)  -  System Check"
        Width="640" Height="560" MinWidth="520" MinHeight="420"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize"
        Background="#F2FAF2" FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True"
        SnapsToDevicePixels="True"
        TextOptions.TextFormattingMode="Display"
        TextOptions.TextRenderingMode="ClearType"
        RenderOptions.ClearTypeHint="Enabled">
  <Window.Resources>
    <ControlTemplate x:Key="PrereqBtnTpl" TargetType="Button">
      <Grid>
        <Border x:Name="BtnBd"
                Background="{TemplateBinding Background}"
                CornerRadius="9"
                Padding="{TemplateBinding Padding}"
                SnapsToDevicePixels="True">
          <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
        <Border CornerRadius="9,9,0,0" Height="14" VerticalAlignment="Top"
                IsHitTestVisible="False" Margin="1,1,1,0"
                SnapsToDevicePixels="True"
                RenderOptions.ClearTypeHint="Enabled">
          <Border.Background>
            <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
              <GradientStop Color="#33FFFFFF" Offset="0"/>
              <GradientStop Color="#0AFFFFFF" Offset="0.6"/>
              <GradientStop Color="#00FFFFFF" Offset="1"/>
            </LinearGradientBrush>
          </Border.Background>
        </Border>
      </Grid>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter TargetName="BtnBd" Property="Opacity" Value="0.88"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
          <Setter TargetName="BtnBd" Property="Opacity" Value="0.72"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
          <Setter TargetName="BtnBd" Property="Opacity" Value="0.4"/>
        </Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Window.Resources>
  <DockPanel>
    <!-- Green gradient header bar (matches main tool) -->
    <Border DockPanel.Dock="Top" Padding="14,12">
      <Border.Background>
        <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
          <GradientStop Color="#006400" Offset="0"/>
          <GradientStop Color="#009900" Offset="1"/>
        </LinearGradientBrush>
      </Border.Background>
      <Border.BorderBrush>
        <SolidColorBrush Color="#0FFFFFFF"/>
      </Border.BorderBrush>
      <Border.BorderThickness>0,0,0,1</Border.BorderThickness>
      <StackPanel HorizontalAlignment="Center" RenderOptions.ClearTypeHint="Enabled">
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
          <Image x:Name="PrereqLogoImg" Height="32" MaxWidth="120"
                 Stretch="Uniform" VerticalAlignment="Center" Visibility="Collapsed"
                 Margin="0,0,10,0"/>
          <TextBlock Text="System Prerequisites Check" FontSize="16" FontWeight="ExtraBold"
                     Foreground="#F1F9F1" VerticalAlignment="Center"/>
        </StackPanel>
        <TextBlock Text="Verifying all requirements before launching the tool."
                   FontSize="11" Foreground="#CCFFCC" HorizontalAlignment="Center" Margin="0,3,0,0"/>
      </StackPanel>
    </Border>

    <!-- Content area -->
    <Grid Margin="20,14,20,20">
      <Grid.RowDefinitions>
        <RowDefinition Height="*"/>
        <RowDefinition Height="130"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <Border Grid.Row="0" Background="White" BorderBrush="#B0D0B0" BorderThickness="1" SnapsToDevicePixels="True"
              CornerRadius="12" Margin="0,0,0,10">
        <Border.Effect>
          <DropShadowEffect BlurRadius="16" ShadowDepth="3" Direction="270" Color="#000000" Opacity="0.10"/>
        </Border.Effect>
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <StackPanel x:Name="SpChecks" Margin="4,4,4,4"/>
        </ScrollViewer>
      </Border>

      <Border Grid.Row="1" Background="#0D1824" CornerRadius="8" Margin="0,0,0,10" SnapsToDevicePixels="True">
        <RichTextBox x:Name="RtbPrereqLog" Background="Transparent" BorderThickness="0"
                     IsReadOnly="True" FontFamily="Consolas" FontSize="10.5"
                     Foreground="#9DB8D8" Padding="8,5"
                     VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"/>
      </Border>

      <Grid Grid.Row="2">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="8"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="TxtPrereqSummary" Grid.Column="0" FontSize="11"
                   Foreground="#5A6480" VerticalAlignment="Center" TextWrapping="Wrap"/>
        <Button x:Name="BtnPrereqInstall" Grid.Column="1" Content="Install Missing"
                Foreground="White" BorderThickness="0"
                Padding="14,8" FontWeight="SemiBold" Cursor="Hand" Visibility="Collapsed"
                Template="{StaticResource PrereqBtnTpl}">
          <Button.Background>
            <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
              <GradientStop Color="#006400" Offset="0"/>
              <GradientStop Color="#009900" Offset="1"/>
            </LinearGradientBrush>
          </Button.Background>
        </Button>
        <Button x:Name="BtnPrereqContinue" Grid.Column="3" Content="Continue &#x2192;"
                Foreground="White" BorderThickness="0"
                Padding="14,8" FontWeight="SemiBold" Cursor="Hand" IsEnabled="False"
                Template="{StaticResource PrereqBtnTpl}">
          <Button.Background>
            <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
              <GradientStop Color="#006400" Offset="0"/>
              <GradientStop Color="#009900" Offset="1"/>
            </LinearGradientBrush>
          </Button.Background>
        </Button>
      </Grid>
    </Grid>
  </DockPanel>
</Window>
"@

    Write-Host "[PREREQ] Parsing XAML..."
    $reader    = [System.Xml.XmlNodeReader]::new($prereqXaml)
    Write-Host "[PREREQ] XmlNodeReader ready"
    $prereqWin = [Windows.Markup.XamlReader]::Load($reader)
    Write-Host "[PREREQ] Window loaded: $($prereqWin -ne $null)"
    Write-Host "[PREREQ] Wiring elements..."
    $spChecks     = $prereqWin.FindName('SpChecks')
    $rtbLog       = $prereqWin.FindName('RtbPrereqLog')
    $btnInstall   = $prereqWin.FindName('BtnPrereqInstall')
    $btnContinue  = $prereqWin.FindName('BtnPrereqContinue')
    $txtSummary   = $prereqWin.FindName('TxtPrereqSummary')
    $prereqLogo   = $prereqWin.FindName('PrereqLogoImg')

    # ── Load logo into prereq header (same DevConfig sources as main tool) ──
    try {
        $logoSrc = $null
        # Priority 1: Base64
        if (-not [string]::IsNullOrWhiteSpace($DevConfig.LogoBase64)) {
            $bytes = [Convert]::FromBase64String($DevConfig.LogoBase64.Trim())
            $ms    = [System.IO.MemoryStream]::new($bytes)
            $bmp   = [System.Windows.Media.Imaging.BitmapImage]::new()
            $bmp.BeginInit()
            $bmp.StreamSource = $ms
            $bmp.CacheOption  = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $bmp.EndInit(); $null = $bmp.Freeze()
            $logoSrc = $bmp
        }
        # Priority 2: Local path
        if (-not $logoSrc -and -not [string]::IsNullOrWhiteSpace($DevConfig.LogoImagePath) -and (Test-Path $DevConfig.LogoImagePath)) {
            $logoSrc = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]::new((Resolve-Path $DevConfig.LogoImagePath).Path))
        }
        # Priority 3: URL
        if (-not $logoSrc -and -not [string]::IsNullOrWhiteSpace($DevConfig.LogoImageUrl)) {
            $bmp = [System.Windows.Media.Imaging.BitmapImage]::new()
            $bmp.BeginInit()
            $bmp.UriSource   = [Uri]::new($DevConfig.LogoImageUrl)
            $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $bmp.EndInit()
            $logoSrc = $bmp
        }
        if ($logoSrc) {
            $prereqLogo.Source     = $logoSrc
            $prereqLogo.Visibility = 'Visible'
            Write-Host "[PREREQ] Logo loaded into header"
        }
    } catch { Write-Host "[PREREQ] Logo load skipped: $($_.Exception.Message)" }
    Write-Host "[PREREQ] Elements wired ok"

    $script:PrereqResult     = $false
    $script:MissingMandatory = [System.Collections.Generic.List[string]]::new()
    $script:MissingOptional  = [System.Collections.Generic.List[string]]::new()
    $script:UpdateAvailable  = [System.Collections.Generic.List[string]]::new()

    # Row registry: key = row-id string, value = hashtable of TextBlock refs
    $rowRegistry = [hashtable]::Synchronized(@{})

    # ── UI helpers (run on dispatcher) ──────────────────────────────────────
    function Add-CheckRow {
        param([string]$RowId, [string]$Label, [string]$Detail)
        $prereqWin.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Normal, [Action]{
            $outer = [System.Windows.Controls.Grid]::new()
            $outer.Margin = [System.Windows.Thickness]::new(8,5,8,5)
            $c0 = [System.Windows.Controls.ColumnDefinition]::new(); $c0.Width = [System.Windows.GridLength]::new(22)
            $c1 = [System.Windows.Controls.ColumnDefinition]::new(); $c1.Width = [System.Windows.GridLength]::new(1,[System.Windows.GridUnitType]::Star)
            $c2 = [System.Windows.Controls.ColumnDefinition]::new(); $c2.Width = [System.Windows.GridLength]::new(1,[System.Windows.GridUnitType]::Auto)
            $outer.ColumnDefinitions.Add($c0); $outer.ColumnDefinitions.Add($c1); $outer.ColumnDefinitions.Add($c2)

            # Icon — hourglass U+231B, rendered in Segoe UI Symbol for reliable glyph
            $tbIcon = [System.Windows.Controls.TextBlock]::new()
            $tbIcon.Text = [string][char]0x231B; $tbIcon.FontSize = 13
            $tbIcon.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe UI Symbol, Segoe UI Emoji, Segoe UI")
            $tbIcon.VerticalAlignment = 'Center'
            $tbIcon.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#999999'))
            $tbIcon.Opacity = 0.45
            # Opacity pulse animation (pre-built, started on demand)
            $pulseAnim = [System.Windows.Media.Animation.DoubleAnimation]::new()
            $pulseAnim.From = 0.35; $pulseAnim.To = 1.0
            $pulseAnim.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(700))
            $pulseAnim.AutoReverse = $true
            $pulseAnim.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever
            [System.Windows.Controls.Grid]::SetColumn($tbIcon, 0)

            $inner = [System.Windows.Controls.StackPanel]::new()
            $inner.Margin = [System.Windows.Thickness]::new(6,0,0,0)
            $inner.VerticalAlignment = 'Center'
            $tbLabel = [System.Windows.Controls.TextBlock]::new()
            $tbLabel.Text = $Label; $tbLabel.FontSize = 12; $tbLabel.FontWeight = 'SemiBold'
            $tbLabel.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#888888'))
            $tbDetail = [System.Windows.Controls.TextBlock]::new()
            $tbDetail.Text = $Detail; $tbDetail.FontSize = 10.5; $tbDetail.TextWrapping = 'Wrap'
            $tbDetail.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#AAAAAA'))
            $tbDetail.Margin = [System.Windows.Thickness]::new(0,1,0,0)
            $inner.Children.Add($tbLabel) | Out-Null
            $inner.Children.Add($tbDetail) | Out-Null
            [System.Windows.Controls.Grid]::SetColumn($inner, 1)

            $tbStatus = [System.Windows.Controls.TextBlock]::new()
            $tbStatus.Text = 'Queued'; $tbStatus.FontSize = 11; $tbStatus.FontWeight = 'SemiBold'
            $tbStatus.VerticalAlignment = 'Center'
            $tbStatus.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#AAAAAA'))
            [System.Windows.Controls.Grid]::SetColumn($tbStatus, 2)

            $outer.Children.Add($tbIcon)   | Out-Null
            $outer.Children.Add($inner)    | Out-Null
            $outer.Children.Add($tbStatus) | Out-Null
            $spChecks.Children.Add($outer) | Out-Null

            $rowRegistry[$RowId] = @{ Icon = $tbIcon; Label = $tbLabel; Detail = $tbDetail; Status = $tbStatus; Outer = $outer; PulseAnimation = $pulseAnim }
        })
    }

    function Apply-RowUpdate {
        param([hashtable]$Msg)
        $rid = $Msg['RowId']
        if (-not $rowRegistry.ContainsKey($rid)) { return }
        $row = $rowRegistry[$rid]
        # Stop opacity pulse animation
        $row['Icon'].BeginAnimation([System.Windows.UIElement]::OpacityProperty, $null)
        $row['Icon'].Opacity = 1.0
        # Reset icon font to default for result glyphs (tick/cross render fine in Segoe UI)
        $row['Icon'].FontFamily = [System.Windows.Media.FontFamily]::new("Segoe UI")
        # Clear row highlight background
        if ($row['Outer']) { $row['Outer'].Background = $null }
        # Apply field updates
        if ($Msg['Icon'])        { $row['Icon'].Text = $Msg['Icon'] }
        if ($Msg['Status'])      { $row['Status'].Text = $Msg['Status'] }
        if ($Msg['Detail'])      { $row['Detail'].Text = $Msg['Detail'] }
        if ($Msg['LabelColor'])  { $row['Label'].Foreground  = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($Msg['LabelColor'])) }
        if ($Msg['StatusColor']) { $row['Status'].Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($Msg['StatusColor'])) }
        if ($Msg['IconColor'])   { $row['Icon'].Foreground   = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($Msg['IconColor'])) }
    }

    function Append-LogLine {
        param([string]$Line, [string]$Color = '#9DB8D8')
        $para = [System.Windows.Documents.Paragraph]::new()
        $run  = [System.Windows.Documents.Run]::new($Line)
        $run.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($Color))
        $para.Margin = [System.Windows.Thickness]::new(0)
        $para.Inlines.Add($run)
        $rtbLog.Document.Blocks.Add($para)
        $rtbLog.ScrollToEnd()
    }

    # ── Pre-create all rows before runspace starts ───────────────────────────
    Add-CheckRow 'ps'     'PowerShell 7+'              'Required for WPF and modern syntax'
    Add-CheckRow 'os'     'Windows OS + WPF'           'WPF requires Windows'
    Add-CheckRow 'ep'     'Execution Policy'           'Must allow script execution'
    foreach ($m in $requiredModules) { Add-CheckRow $m.Name $m.Label "Required module  -  $($m.Name)" }
    foreach ($m in $optionalModules) { Add-CheckRow $m.Name $m.Label "Optional  -  $($m.Name)" }

    # ── Launch background runspace ───────────────────────────────────────────
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'MTA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('Queue',           $queue)
    $rs.SessionStateProxy.SetVariable('RequiredModules', $requiredModules)
    $rs.SessionStateProxy.SetVariable('OptionalModules', $optionalModules)

    $ps = [powershell]::Create()
    $ps.Runspace = $rs

    $null = $ps.AddScript({
        function Q {
            param([hashtable]$Msg)
            $Queue.Enqueue($Msg)
        }
        function TS { (Get-Date).ToString('HH:mm:ss') }

        Q @{ Type='log'; Line="[$(TS)] Starting prerequisite checks..."; Color='#C8DEF5' }
        Start-Sleep -Milliseconds 200

        # 1. PowerShell version
        Q @{ Type='checking'; RowId='ps' }
        Start-Sleep -Milliseconds 250
        Q @{ Type='log'; Line="[$(TS)]   Checking PowerShell version..."; Color='#9DB8D8' }
        $psv = $PSVersionTable.PSVersion
        if ($psv.Major -ge 7) {
            Q @{ Type='rowlog'; Line="[$(TS)]   $([char]0x2713) PowerShell $($psv.ToString())"; Color='#5DC98B'; RowId='ps'; Icon="$([char]0x2713)"; Status="OK  v$($psv.Major).$($psv.Minor).$($psv.Build)"; StatusColor='#14532D'; IconColor='#14532D'; LabelColor='#14532D' }
        } else {
            Q @{ Type='rowlog'; Line="[$(TS)]   $([char]0x2717) PowerShell $($psv.ToString())  -  need 7+"; Color='#F06C6C'; RowId='ps'; Icon="$([char]0x2717)"; Status="FAIL  v$($psv.ToString())"; LabelColor='#8B0000'; StatusColor='#8B0000'; IconColor='#8B0000' }
            Q @{ Type='mandatory'; Item='PowerShell 7+   -   https://github.com/PowerShell/PowerShell/releases' }
        }

        # 2. Windows OS
        Q @{ Type='checking'; RowId='os' }
        Start-Sleep -Milliseconds 250
        Q @{ Type='log'; Line="[$(TS)]   Checking OS..."; Color='#9DB8D8' }
        if ($IsWindows) {
            Q @{ Type='rowlog'; Line="[$(TS)]   $([char]0x2713) Windows  -  WPF available"; Color='#5DC98B'; RowId='os'; Icon="$([char]0x2713)"; Status="OK  Windows"; StatusColor='#14532D'; IconColor='#14532D'; LabelColor='#14532D'; Detail=[System.Environment]::OSVersion.VersionString }
        } else {
            Q @{ Type='rowlog'; Line="[$(TS)]   $([char]0x2717) Not Windows  -  tool is Windows-only"; Color='#F06C6C'; RowId='os'; Icon="$([char]0x2717)"; Status='NOT WINDOWS'; LabelColor='#8B0000'; StatusColor='#8B0000'; IconColor='#8B0000' }
            Q @{ Type='mandatory'; Item='Windows OS   -   WPF is Windows-only' }
        }

        # 3. Execution Policy
        Q @{ Type='checking'; RowId='ep' }
        Start-Sleep -Milliseconds 250
        Q @{ Type='log'; Line="[$(TS)]   Checking execution policy..."; Color='#9DB8D8' }
        $ep  = Get-ExecutionPolicy -Scope CurrentUser
        $epm = Get-ExecutionPolicy -Scope LocalMachine
        $epOk = $ep -in @('RemoteSigned','Unrestricted','Bypass') -or $epm -in @('RemoteSigned','Unrestricted','Bypass')
        if ($epOk) {
            Q @{ Type='rowlog'; Line="[$(TS)]   $([char]0x2713) Execution policy: $ep"; Color='#5DC98B'; RowId='ep'; Icon="$([char]0x2713)"; Status="OK  ($ep)"; StatusColor='#14532D'; IconColor='#14532D'; LabelColor='#14532D' }
        } else {
            Q @{ Type='rowlog'; Line="[$(TS)]   [!] Policy '$ep' may block scripts"; Color='#F5A623'; RowId='ep'; Icon='[!]'; Status="$ep"; StatusColor='#B45B00' }
            Q @{ Type='mandatory'; Item="ExecutionPolicy   -   run: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser" }
        }

        # 4. Required modules (local check only, fast)
        foreach ($mod in $RequiredModules) {
            Q @{ Type='checking'; RowId=$mod.Name }
            Start-Sleep -Milliseconds 250
            Q @{ Type='log'; Line="[$(TS)]   Checking $($mod.Name)..."; Color='#9DB8D8' }
            $inst = Get-Module -ListAvailable -Name $mod.Name -ErrorAction SilentlyContinue |
                    Sort-Object Version -Descending | Select-Object -First 1
            if ($inst) {
                $v = $inst.Version.ToString()
                Q @{ Type='rowlog'; Line="[$(TS)]   $([char]0x2713) $($mod.Name)  v$v"; Color='#5DC98B'; RowId=$mod.Name; Icon="$([char]0x2713)"; Status="OK  v$v"; StatusColor='#14532D'; IconColor='#14532D'; LabelColor='#14532D'; Detail="Installed: v$v" }
                Q @{ Type='installed'; ModName=$mod.Name; Version=$v }
            } else {
                Q @{ Type='rowlog'; Line="[$(TS)]   $([char]0x2717) $($mod.Name)  -  NOT INSTALLED"; Color='#F06C6C'; RowId=$mod.Name; Icon="$([char]0x2717)"; Status='NOT INSTALLED'; LabelColor='#8B0000'; StatusColor='#8B0000'; IconColor='#8B0000'; Detail="Missing" }
                Q @{ Type='mandatory'; Item=$mod.Name }
            }
        }

        # 5. Optional modules
        foreach ($mod in $OptionalModules) {
            Q @{ Type='checking'; RowId=$mod.Name }
            Start-Sleep -Milliseconds 250
            Q @{ Type='log'; Line="[$(TS)]   Checking $($mod.Name) (optional)..."; Color='#9DB8D8' }
            $inst = Get-Module -ListAvailable -Name $mod.Name -ErrorAction SilentlyContinue |
                    Sort-Object Version -Descending | Select-Object -First 1
            if ($inst) {
                $v = $inst.Version.ToString()
                Q @{ Type='rowlog'; Line="[$(TS)]   $([char]0x2713) $($mod.Name)  v$v"; Color='#5DC98B'; RowId=$mod.Name; Icon="$([char]0x2713)"; Status="OK  v$v"; StatusColor='#14532D'; IconColor='#14532D'; LabelColor='#14532D'; Detail="Installed: v$v" }
                Q @{ Type='installed'; ModName=$mod.Name; Version=$v }
            } else {
                Q @{ Type='rowlog'; Line="[$(TS)]   o $($mod.Name) not installed (optional)"; Color='#F5A623'; RowId=$mod.Name; Icon='o'; Status='Not installed'; StatusColor='#B45B00'; Detail='Optional  -  export falls back to CSV' }
                Q @{ Type='optional'; Item=$mod.Name }
            }
        }

        Q @{ Type='log'; Line="[$(TS)] ---------------------------------"; Color='#333D4D' }
        Q @{ Type='log'; Line="[$(TS)]   Local checks complete. Checking for module updates online..."; Color='#9DB8D8' }

        # 6. Online update check (slow  -  runs after local checks)
        $allModNames = ($RequiredModules + $OptionalModules).Name
        foreach ($modName in $allModNames) {
            $inst = Get-Module -ListAvailable -Name $modName -ErrorAction SilentlyContinue |
                    Sort-Object Version -Descending | Select-Object -First 1
            if (-not $inst) { continue }
            try {
                Q @{ Type='log'; Line="[$(TS)]   Checking online: $modName..."; Color='#9DB8D8' }
                $gallery = Find-Module -Name $modName -ErrorAction SilentlyContinue
                if ($gallery -and [version]$gallery.Version -gt [version]$inst.Version) {
                    Q @{ Type='rowlog'; Line="[$(TS)]   ^ Update: $modName  $($inst.Version) -> $($gallery.Version)"; Color='#9DB8D8'; RowId=$modName; Icon='^'; Status="Update $($gallery.Version)"; StatusColor='#1A5FA8'; Detail="Installed: v$($inst.Version)  ->  Available: v$($gallery.Version)" }
                    Q @{ Type='update'; Item=$modName }
                } else {
                    Q @{ Type='log'; Line="[$(TS)]   $([char]0x2713) $modName up to date"; Color='#5DC98B' }
                }
            } catch {
                Q @{ Type='log'; Line="[$(TS)]   [!] Could not check online for $modName"; Color='#F5A623' }
            }
        }

        Q @{ Type='log'; Line="[$(TS)] ---------------------------------"; Color='#333D4D' }
        Q @{ Type='done' }
    })

    $bgHandle = $ps.BeginInvoke()

    # ── DispatcherTimer  -  drains queue on UI thread ──────────────────────────
    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(120)

    $timer.Add_Tick({
        try {
        $msg = $null
        while ($queue.TryDequeue([ref]$msg)) {
            switch ($msg['Type']) {
                'log' {
                    Append-LogLine -Line $msg['Line'] -Color $msg['Color']
                }
                'checking' {
                    $rid = $msg['RowId']
                    if ($rowRegistry.ContainsKey($rid)) {
                        $row = $rowRegistry[$rid]
                        # Activate: pulsing hourglass, teal colour, row highlight
                        $row['Icon'].Text = [string][char]0x231B
                        $row['Icon'].FontFamily = [System.Windows.Media.FontFamily]::new('Segoe UI Symbol, Segoe UI Emoji, Segoe UI')
                        $row['Icon'].Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#0A9396'))
                        $row['Status'].Text = 'Checking...'
                        $row['Status'].Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#0A9396'))
                        $row['Label'].Foreground  = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#1A1F2E'))
                        $row['Detail'].Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#5A6480'))
                        $row['Outer'].Background  = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#F0FAF0'))
                        # Start opacity pulse animation
                        $row['Icon'].BeginAnimation([System.Windows.UIElement]::OpacityProperty, $row['PulseAnimation'])
                    }
                }
                'row' {
                    Apply-RowUpdate -Msg $msg
                }
                'rowlog' {
                    # Atomic: render log line THEN row update in same dispatcher callback
                    Append-LogLine -Line $msg['Line'] -Color $msg['Color']
                    Apply-RowUpdate -Msg $msg
                }
                'mandatory' {
                    $script:MissingMandatory.Add($msg['Item'])
                }
                'optional' {
                    $script:MissingOptional.Add($msg['Item'])
                }
                'update' {
                    $script:UpdateAvailable.Add($msg['Item'])
                }
                'done' {
                    $timer.Stop()
                    $ps.EndInvoke($bgHandle)
                    $rs.Close(); $ps.Dispose()

                    $hasMand  = $script:MissingMandatory.Count -gt 0
                    $hasOpt   = $script:MissingOptional.Count -gt 0
                    $hasUpd   = $script:UpdateAvailable.Count -gt 0

                    # Summary
                    $parts = @()
                    if ($hasMand) { $parts += "$($script:MissingMandatory.Count) required missing" }
                    if ($hasOpt)  { $parts += "$($script:MissingOptional.Count) optional missing" }
                    if ($hasUpd)  { $parts += "$($script:UpdateAvailable.Count) update(s) available" }
                    if ($parts.Count -eq 0) {
                        $txtSummary.Text = "$([char]0x2713)  All checks passed."
                        Append-LogLine "[$(Get-Date -f 'HH:mm:ss')]   $([char]0x2713) All checks passed  -  ready to launch." '#5DC98B'
                    } else {
                        $txtSummary.Text = $parts -join '   |   '
                        if ($hasMand) { Append-LogLine "[$(Get-Date -f 'HH:mm:ss')]   $([char]0x2717) Mandatory items missing." '#F06C6C' }
                    }

                    # Install button
                    $installable = [System.Collections.Generic.List[string]]::new()
                    $allModNames = ($requiredModules + $optionalModules).Name
                    foreach ($item in ($script:MissingMandatory + $script:MissingOptional + $script:UpdateAvailable)) {
                        if ($allModNames -contains $item) { $installable.Add($item) }
                    }
                    if ($installable.Count -gt 0) {
                        $btnInstall.Visibility = 'Visible'
                        if   ($hasMand -and $hasUpd) { $btnInstall.Content = 'Install Missing + Update' }
                        elseif ($hasMand)             { $btnInstall.Content = 'Install Missing' }
                        elseif ($hasUpd)              { $btnInstall.Content = 'Update Modules' }
                        else                          { $btnInstall.Content = 'Install Optional' }
                    }

                     if (-not $hasMand) {
                         $btnContinue.IsEnabled = $true
                         $txtSummary.Text = "✓  All checks passed  -  click Continue to proceed"
                     }
                }
            }
        }
        } catch {
            Write-ErrorLog "PrereqTimer error: $($_.Exception.Message)"
        }
    })

    $timer.Start()

    # ── Native WPF confirm dialog (replaces MessageBox) ──────────────────────
    function Show-PrereqConfirm {
        param([string]$Title, [string]$Message, [bool]$IsWarning = $false)
        # Title and Message set in code  -  never interpolated into XML (avoids
        # XML-special chars: &, <, >, newlines breaking the document parse)
        [xml]$dlgXml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Confirm" Width="480" SizeToContent="Height"
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize"
        Background="#F8FBF8" FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True" SnapsToDevicePixels="True"
        TextOptions.TextFormattingMode="Display"
        TextOptions.TextRenderingMode="ClearType">
  <Grid Margin="22,18,22,18">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="14"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock x:Name="TxtMsg" Grid.Row="0" TextWrapping="Wrap" LineHeight="20"
               Foreground="#1A1F2E" FontSize="12"/>
    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="BtnNo" Content="Cancel" Width="90" Padding="0,8" Margin="0,0,10,0"
              Background="Transparent" BorderBrush="#B0D0B0" BorderThickness="1"
              Cursor="Hand" FontSize="12">
        <Button.Template>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="5" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="#E8F5E8"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Button.Template>
      </Button>
      <Button x:Name="BtnYes" Content="Proceed" Width="100" Padding="0,8"
              Background="#007A00" BorderThickness="0" Cursor="Hand"
              FontSize="12" FontWeight="SemiBold" Foreground="White">
        <Button.Template>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" CornerRadius="5"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="#009900"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Button.Template>
      </Button>
    </StackPanel>
  </Grid>
</Window>
'@
        $dlgReader = [System.Xml.XmlNodeReader]::new($dlgXml)
        $dlg       = [Windows.Markup.XamlReader]::Load($dlgReader)
        $dlg.Owner = $prereqWin
        $dlg.Title = $Title
        $dlg.FindName('TxtMsg').Text = $Message

        # Colour the Proceed button  -  evaluate colour before passing to method
        $yesBtn    = $dlg.FindName('BtnYes')
        $btnColor  = if ($IsWarning) { '#C0392B' } else { '#007A00' }
        $yesBtn.Background = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString($btnColor))

        $script:PrereqDlgResult = $false
        $dlg.FindName('BtnNo').Add_Click({ $script:PrereqDlgResult = $false; $dlg.Close() })
        $yesBtn.Add_Click({ $script:PrereqDlgResult = $true; $dlg.Close() })
        $dlg.ShowDialog() | Out-Null
        return $script:PrereqDlgResult
    }

    # ── Install button ────────────────────────────────────────────────────────
    $btnInstall.Add_Click({
        $allModNames = ($requiredModules + $optionalModules).Name
        $toInstall   = [System.Collections.Generic.List[string]]::new()
        foreach ($item in ($script:MissingMandatory + $script:MissingOptional + $script:UpdateAvailable)) {
            if ($allModNames -contains $item -and -not $toInstall.Contains($item)) {
                $toInstall.Add($item)
            }
        }
        if ($toInstall.Count -eq 0) {
            $null = Show-PrereqConfirm -Title "Nothing to Install" `
                -Message "No installable modules found.`nPS version, OS and execution policy issues must be resolved manually."
            return
        }

        $listTxt = ($toInstall | ForEach-Object { "  - $_" }) -join "`n"

        # Install using CurrentUser scope  -  no admin rights required
        $confirmMsg  = "The following modules will be installed or updated:`n`n$listTxt`n`n"
        $confirmMsg += "Modules will be installed in CurrentUser scope (no admin required)."
        if (-not (Show-PrereqConfirm -Title "Install Modules" -Message $confirmMsg)) { return }

            $btnInstall.IsEnabled = $false
            $btnContinue.IsEnabled = $false
            $btnInstall.Content   = "[wait] Installing..."
            Append-LogLine "[$(Get-Date -f 'HH:mm:ss')] -- Installing modules (elevated session) --" '#C8DEF5'

            # Use a runspace so Install-Module output streams to our queue live
            # All vars promoted to script scope  -  closure capture is unreliable in WPF handlers
            $script:InstallQueue = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
            $script:InstallRs = [runspacefactory]::CreateRunspace()
            $script:InstallRs.ApartmentState = 'MTA'
            $script:InstallRs.ThreadOptions  = 'ReuseThread'
            $script:InstallRs.Open()
            $script:InstallRs.SessionStateProxy.SetVariable('InstallQueue', $script:InstallQueue)
            $script:InstallRs.SessionStateProxy.SetVariable('ModulesToInstall', $toInstall)

            $script:InstallPs = [powershell]::Create()
            $script:InstallPs.Runspace = $script:InstallRs
            $null = $script:InstallPs.AddScript({
                function IQ { param([hashtable]$M) $InstallQueue.Enqueue($M) }
                function TS  { (Get-Date).ToString('HH:mm:ss') }
                $allOk = $true
                foreach ($modName in $ModulesToInstall) {
                    IQ @{ Type='log'; Line="[$(TS)]   Installing $modName ..."; Color='#C8DEF5' }
                    try {
                        # Redirect verbose/progress into variable to avoid console output
                        $null = Install-Module $modName -Scope CurrentUser -Force -AllowClobber `
                            -ErrorAction Stop -Verbose 4>&1 | ForEach-Object {
                                IQ @{ Type='log'; Line="[$(TS)]     $_"; Color='#9DB8D8' }
                            }
                        IQ @{ Type='log'; Line="[$(TS)]   $([char]0x2713) $modName  -  done."; Color='#5DC98B' }
                    } catch {
                        IQ @{ Type='log'; Line="[$(TS)]   $([char]0x2717) $modName  -  failed: $($_.Exception.Message)"; Color='#F06C6C' }
                        $allOk = $false
                    }
                }
                IQ @{ Type='done'; AllOk=$allOk }
            })
            $script:InstallHandle = $script:InstallPs.BeginInvoke()

            # Drain install queue via a separate timer
            $script:InstallTimer = [System.Windows.Threading.DispatcherTimer]::new()
            $script:InstallTimer.Interval = [TimeSpan]::FromMilliseconds(150)
            $script:InstallTimer.Add_Tick({
                $m = $null
                while ($script:InstallQueue.TryDequeue([ref]$m)) {
                    if ($m['Type'] -eq 'log') {
                        Append-LogLine $m['Line'] $m['Color']
                    } elseif ($m['Type'] -eq 'done') {
                        $script:InstallTimer.Stop()
                        $script:InstallPs.EndInvoke($script:InstallHandle)
                        $script:InstallRs.Close(); $script:InstallPs.Dispose()

                        if ($m['AllOk']) {
                            Append-LogLine "[$(Get-Date -f 'HH:mm:ss')] $([char]0x2713) All modules installed successfully." '#5DC98B'
                            Append-LogLine "[$(Get-Date -f 'HH:mm:ss')]   Note: A PowerShell session restart may be needed" '#F5A623'
                            Append-LogLine "[$(Get-Date -f 'HH:mm:ss')]   for newly installed modules to load correctly." '#F5A623'
                        } else {
                            Append-LogLine "[$(Get-Date -f 'HH:mm:ss')] [!] Some modules failed  -  see log above." '#F5A623'
                        }

                        # Show native restart prompt
                        $restartMsg  = "Modules installed successfully.`n`n"
                        $restartMsg += "PowerShell modules load at session start  -  to ensure the`n"
                        $restartMsg += "tool uses the updated versions, it needs to relaunch in a`n"
                        $restartMsg += "fresh session.`n`nRelaunch now?"

                        $btnInstall.Content    = "Install Missing"
                        $btnInstall.IsEnabled  = $true

                        if (Show-PrereqConfirm -Title "Relaunch Required" -Message $restartMsg) {
                            Append-LogLine "[$(Get-Date -f 'HH:mm:ss')]   Relaunching tool in new session..." '#C8DEF5'
                            $scriptPath = $PSCommandPath
                            Start-Process pwsh -ArgumentList "-STA -File `"$scriptPath`""
                            $prereqWin.Close()
                            [System.Environment]::Exit(0)
                        } else {
                            # User declined relaunch  -  re-run checks so they see current state
                            $script:RerunPrereq = $true
                            $script:PrereqResult = $false
                            $prereqWin.Close()
                        }
                    }
                }
            })
            $script:InstallTimer.Start()
            return
    })

    # ── Continue / Close ──────────────────────────────────────────────────────
    $btnContinue.Add_Click({
        $script:PrereqResult = $true
        # Show feedback before window closes - main window takes ~20s to build
        $btnContinue.IsEnabled = $false
        $btnContinue.Content   = 'Loading...'
        $btnInstall.IsEnabled  = $false
        if ($txtSummary) { $txtSummary.Text = 'Building main window, please wait...' }
        $prereqWin.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [Action]{})
        $prereqWin.Close()
    })

    $prereqWin.Add_Closed({
        $timer.Stop()
        if (-not $script:PrereqResult) { exit 0 }
    })

    $prereqWin.ShowDialog() | Out-Null
    return $script:PrereqResult
}



# ── Error log + trap defined BEFORE prereq so the trap can log errors ──
# ── Diagnostic: log any unhandled errors to a temp file ──────────────────────
# Visible error log so we can diagnose failures when the console is hidden.
# Remove once stable.
$_errDir = if (-not [string]::IsNullOrWhiteSpace($DevConfig.LogFolder) -and
             (Test-Path $DevConfig.LogFolder)) { $DevConfig.LogFolder } else { $env:TEMP }
$script:ErrorLogPath = Join-Path $_errDir ("MDMODA_Error_{0}.log" -f $script:SessionStamp)
try { Remove-Item $script:ErrorLogPath -Force -ErrorAction SilentlyContinue } catch {}
Write-Host "[STARTUP] Error log: $script:ErrorLogPath"
function Write-ErrorLog {
    param([string]$Msg)
    try {
        $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Add-Content -Path $script:ErrorLogPath -Value "[$ts] $Msg" -Encoding UTF8
    } catch {}
}
Write-ErrorLog "Prereq passed. Building main window..."
Write-ErrorLog "PS Version  : $($PSVersionTable.PSVersion)"
Write-ErrorLog "OS Version  : $([System.Environment]::OSVersion.VersionString)"
Write-ErrorLog "User        : $($env:USERNAME)@$($env:USERDOMAIN)"
Write-ErrorLog "LogFolder   : $($DevConfig.LogFolder)"
Write-ErrorLog "Error log   : $script:ErrorLogPath"
Write-ErrorLog "Transcript  : $(if ($script:TranscriptStarted) { $tsPath } else { 'NOT STARTED' })"
$Error.Clear()
Write-EarlyLog "Checkpoint 1: Starting XAML definition..."
Write-ErrorLog "Checkpoint 1: Starting XAML definition..."

# ── Catch-all trap  -  logs any terminating error to the error log ─────────────
trap {
    $errMsg = "TRAP at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
    $errStack = $_.ScriptStackTrace
    Write-ErrorLog $errMsg
    Write-ErrorLog "Stack: $errStack"
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        [System.Windows.Forms.MessageBox]::Show(
            "$errMsg`n`nSee: $script:ErrorLogPath",
            "MDM-ODA - Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    } catch {}
    break
}

# Run the prereq check  -  loop supports in-session reinstall when already admin
$script:RerunPrereq = $false
do {
    $script:RerunPrereq = $false
    if (-not (Start-PrereqCheck)) { exit 0 }
} while ($script:RerunPrereq)

#endregion PREREQ CHECK




# ============================================================
#region XAML  -  UI Definition
# ============================================================

# ── Compute window size from available screen working area ──
# Uses SystemParameters (WPF logical pixels already DPI-adjusted) so the
# window fits on any screen size / scaling factor without clipping.
$_screen = [System.Windows.SystemParameters]
$_workW  = [int]$_screen::WorkArea.Width
$_workH  = [int]$_screen::WorkArea.Height
# Target 92% of working area, clamped to config defaults as ceiling
$_winW   = [int][Math]::Min([Math]::Floor($_workW * 0.92), $DevConfig.WindowWidth)
$_winH   = [int][Math]::Min([Math]::Floor($_workH * 0.92), $DevConfig.WindowHeight)
# Never go below the minimum sizes
$_winW   = [Math]::Max($_winW, $DevConfig.MinWidth)
$_winH   = [Math]::Max($_winH, $DevConfig.MinHeight)

[xml]$Xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="$($DevConfig.WindowTitle) &#x26A1;"
    Width="$_winW"
    Height="$_winH"
    MinWidth="$($DevConfig.MinWidth)"
    MinHeight="$($DevConfig.MinHeight)"
    WindowStartupLocation="CenterScreen"
    Background="#F2FAF2"
    FontFamily="$($DevConfig.FontFamily)"
    FontSize="$($DevConfig.FontSizeBase)"
    UseLayoutRounding="True"
    SnapsToDevicePixels="True"
    TextOptions.TextFormattingMode="Display"
    TextOptions.TextRenderingMode="ClearType"
    RenderOptions.ClearTypeHint="Enabled">

  <Window.Resources>

    <!-- Glassy button base template -->
    <ControlTemplate x:Key="GlassyBtnTemplate" TargetType="Button">
      <Grid>
        <Border x:Name="BtnBorder"
                Background="{TemplateBinding Background}"
                CornerRadius="9"
                Padding="{TemplateBinding Padding}"
                SnapsToDevicePixels="True">
          <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
        <Border CornerRadius="9,9,0,0" Height="14" VerticalAlignment="Top"
                IsHitTestVisible="False" Margin="1,1,1,0"
                SnapsToDevicePixels="True"
                RenderOptions.ClearTypeHint="Enabled">
          <Border.Background>
            <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
              <GradientStop Color="#33FFFFFF" Offset="0"/>
              <GradientStop Color="#0AFFFFFF" Offset="0.6"/>
              <GradientStop Color="#00FFFFFF" Offset="1"/>
            </LinearGradientBrush>
          </Border.Background>
        </Border>
      </Grid>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter TargetName="BtnBorder" Property="Opacity" Value="0.88"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
          <Setter TargetName="BtnBorder" Property="Opacity" Value="0.72"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
          <Setter TargetName="BtnBorder" Property="Opacity" Value="0.4"/>
        </Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>

    <!-- Primary button (aqua blue) -->
    <Style x:Key="PrimaryBtn" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="14,9"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template" Value="{StaticResource GlassyBtnTemplate}"/>
      <Setter Property="Background">
        <Setter.Value>
          <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
            <GradientStop Color="#005F73" Offset="0"/>
            <GradientStop Color="#0A9396" Offset="1"/>
          </LinearGradientBrush>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsEnabled" Value="False">
          <Setter Property="Opacity" Value="0.45"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- Accent button (Connect - aqua blue with glow) -->
    <Style x:Key="AccentBtn" TargetType="Button" BasedOn="{StaticResource PrimaryBtn}">
      <Setter Property="Effect">
        <Setter.Value>
          <DropShadowEffect BlurRadius="18" ShadowDepth="0" Color="#0A9396" Opacity="0.30"/>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Disconnect button (amber) -->
    <Style x:Key="AmberBtn" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="14,9"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template" Value="{StaticResource GlassyBtnTemplate}"/>
      <Setter Property="Effect">
        <Setter.Value>
          <DropShadowEffect BlurRadius="14" ShadowDepth="0" Color="#F59E0B" Opacity="0.22"/>
        </Setter.Value>
      </Setter>
      <Setter Property="Background">
        <Setter.Value>
          <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
            <GradientStop Color="#78350F" Offset="0"/>
            <GradientStop Color="#D97706" Offset="1"/>
          </LinearGradientBrush>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsEnabled" Value="False">
          <Setter Property="Opacity" Value="0.45"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- Stop button (blood red) -->
    <Style x:Key="StopBtn" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="12,9"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template" Value="{StaticResource GlassyBtnTemplate}"/>
      <Setter Property="Effect">
        <Setter.Value>
          <DropShadowEffect BlurRadius="16" ShadowDepth="0" Color="#DC2626" Opacity="0.30"/>
        </Setter.Value>
      </Setter>
      <Setter Property="Background">
        <Setter.Value>
          <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
            <GradientStop Color="#7F1D1D" Offset="0"/>
            <GradientStop Color="#DC2626" Offset="1"/>
          </LinearGradientBrush>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Ghost / icon button -->
    <Style x:Key="GhostBtn" TargetType="Button">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="#4A7060"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="{TemplateBinding Background}"
                    CornerRadius="7" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#E8F5E8"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Outline button -->
    <Style x:Key="OutlineBtn" TargetType="Button">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="#22C55E"/>
      <Setter Property="BorderBrush" Value="#22C55E"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="12,7"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="7" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="#D4EDDA"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Op tile button (sidebar) -->
    <Style x:Key="OpTileBtn" TargetType="Button">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="#374151"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Grid>
              <Border x:Name="AccentBar" Width="3" HorizontalAlignment="Left"
                      VerticalAlignment="Stretch" CornerRadius="0,2,2,0" Opacity="0"/>
              <Border x:Name="TileBg" Background="{TemplateBinding Background}"
                      BorderBrush="#B0D0B0" BorderThickness="1"
                      CornerRadius="8" Margin="0,1,0,1" Padding="10,9,10,9">
                <ContentPresenter HorizontalAlignment="Stretch"/>
              </Border>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="TileBg" Property="Background" Value="#E8F5E8"/>
                <Setter TargetName="TileBg" Property="BorderBrush" Value="#B0D0B0"/>
                <Setter TargetName="AccentBar" Property="Opacity" Value="0.6"/>
                <Setter TargetName="AccentBar" Property="Background" Value="#22C55E"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="TileBg" Property="Background" Value="#D4EDDA"/>
                <Setter TargetName="TileBg" Property="BorderBrush" Value="#86C18A"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- TextBox -->
    <Style x:Key="InputBox" TargetType="TextBox">
      <Setter Property="Background" Value="#FFFFFF"/>
      <Setter Property="Foreground" Value="#374151"/>
      <Setter Property="BorderBrush" Value="#B0D0B0"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="10,8"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="CaretBrush" Value="#22C55E"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="Bd" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="8">
              <Grid TextElement.Foreground="{TemplateBinding Foreground}">
                <ScrollViewer x:Name="PART_ContentHost" TextElement.Foreground="{TemplateBinding Foreground}" Margin="{TemplateBinding Padding}"
                              VerticalAlignment="Center"/>
                <TextBlock x:Name="Placeholder" Text="{TemplateBinding Tag}"
                           Foreground="#9CA3AF" FontSize="{TemplateBinding FontSize}"
                           Margin="{TemplateBinding Padding}" VerticalAlignment="Top"
                           IsHitTestVisible="False" TextWrapping="Wrap" Visibility="Collapsed"/>
              </Grid>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsFocused" Value="True">
                <Setter TargetName="Bd" Property="BorderBrush" Value="#4022C55E"/>
                <Setter Property="Foreground" Value="#1A1F2E"/>
              </Trigger>
              <MultiTrigger>
                <MultiTrigger.Conditions>
                  <Condition Property="Text" Value=""/>
                  <Condition Property="IsFocused" Value="False"/>
                </MultiTrigger.Conditions>
                <Setter TargetName="Placeholder" Property="Visibility" Value="Visible"/>
              </MultiTrigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ComboBox -->
    <Style x:Key="StyledCombo" TargetType="ComboBox">
      <Setter Property="Background" Value="#FFFFFF"/>
      <Setter Property="Foreground" Value="#374151"/>
      <Setter Property="BorderBrush" Value="#B0D0B0"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="10,8"/>
      <Setter Property="FontSize" Value="12"/>
    </Style>

    <!-- Field label -->
    <Style x:Key="FieldLabel" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#14532D"/>
      <Setter Property="FontSize" Value="10"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="Margin" Value="0,0,0,5"/>
    </Style>

    <!-- Card panel -->
    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="#FFFFFF"/>
      <Setter Property="BorderBrush" Value="#B0D0B0"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="12"/>
      <Setter Property="Padding" Value="18"/>
      <Setter Property="Effect">
        <Setter.Value>
          <DropShadowEffect BlurRadius="20" ShadowDepth="4" Direction="270"
                            Color="#000000" Opacity="0.35"/>
        </Setter.Value>
      </Setter>
    </Style>

  </Window.Resources>

  <Grid>

    <!-- Background image layer -->
    <Image x:Name="BgImage" Stretch="UniformToFill"
           Opacity="$($DevConfig.BackgroundImageOpacity)"
           HorizontalAlignment="Stretch" VerticalAlignment="Stretch"
           IsHitTestVisible="False" SnapsToDevicePixels="True"/>

    <DockPanel>

      <!-- HEADER BAR -->
      <Border x:Name="HeaderBorder" DockPanel.Dock="Top" Padding="18,14" SnapsToDevicePixels="True" RenderOptions.ClearTypeHint="Enabled">
        <Border.Background>
          <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
            <GradientStop Color="#006400" Offset="0"/>
            <GradientStop Color="#009900" Offset="1"/>
          </LinearGradientBrush>
        </Border.Background>
        <Border.BorderBrush>
          <SolidColorBrush Color="#0FFFFFFF"/>
        </Border.BorderBrush>
        <Border.BorderThickness>0,0,0,1</Border.BorderThickness>
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>

           <Border Grid.Column="0" Width="Auto" MaxWidth="$($DevConfig.LogoWidth)"
                   Height="$($DevConfig.LogoHeight)" Margin="0,0,0,0"
                   VerticalAlignment="Center" ClipToBounds="True" Background="Transparent">
             <Image x:Name="LogoImg" Height="$($DevConfig.LogoHeight)" MaxWidth="$($DevConfig.LogoWidth)"
                    Stretch="Uniform" HorizontalAlignment="Left"
                    VerticalAlignment="Center" Visibility="Collapsed"/>
           </Border>

          <!-- Title + bolt + status -->
          <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="$($DevConfig.LogoTitleSpacing),0,0,0">
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="$($DevConfig.WindowTitle)" Foreground="#F1F9F1"
                         FontSize="17" FontWeight="Black" VerticalAlignment="Center"/>
              <TextBlock Text=" &#x26A1;" Foreground="#FFD700" FontSize="15"
                         VerticalAlignment="Center"/>
            </StackPanel>
            <TextBlock x:Name="TxtStatusBar" Text="Not connected"
                       Foreground="#CCFFCC" FontSize="11" Margin="0,2,0,0"/>
          </StackPanel>

              <Border Grid.Column="2"/>

          <!-- Button cluster -->
          <StackPanel Grid.Column="3" Orientation="Horizontal" VerticalAlignment="Center">
            <Button x:Name="BtnConnect" Content="Connect" Style="{StaticResource AccentBtn}"
                    Width="90" Margin="0,0,7,0"/>
            <Button x:Name="BtnClearInputs" Content="x  Clear" Visibility="Collapsed"
                    Padding="10,9" FontSize="12" FontWeight="Bold" Cursor="Hand"
                    Width="80" Margin="0,0,7,0" Foreground="White" BorderThickness="0">
              <Button.Background>
                <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                  <GradientStop Color="#374151" Offset="0"/>
                  <GradientStop Color="#6B7280" Offset="1"/>
                </LinearGradientBrush>
              </Button.Background>
              <Button.Template>
                <ControlTemplate TargetType="Button">
                  <Border Background="{TemplateBinding Background}" CornerRadius="9"
                          Padding="{TemplateBinding Padding}">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                  </Border>
                  <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True">
                      <Setter Property="Opacity" Value="0.85"/>
                    </Trigger>
                  </ControlTemplate.Triggers>
                </ControlTemplate>
              </Button.Template>
            </Button>
            <Button x:Name="BtnStop" Style="{StaticResource StopBtn}"
                    Visibility="Collapsed" Width="88" Margin="0,0,7,0">
              <Button.Content>
                <StackPanel Orientation="Horizontal">
                  <TextBlock Text="&#xE71A;" FontFamily="Segoe MDL2 Assets" FontSize="12"
                             VerticalAlignment="Center" Margin="0,0,5,0" Foreground="White"/>
                  <TextBlock Text="Stop" VerticalAlignment="Center"/>
                </StackPanel>
              </Button.Content>
            </Button>
            <Button x:Name="BtnDisconnect" Content="Disconnect" Style="{StaticResource AmberBtn}"
                    Width="106" IsEnabled="False" Margin="0,0,7,0"/>
            <Button x:Name="BtnFeedback"
                    Padding="10,9" FontSize="11" FontWeight="SemiBold" Cursor="Hand"
                    Width="96" Margin="0,0,7,0" Foreground="#1C1C00" BorderThickness="1"
                    BorderBrush="#D4A017" ToolTip="Open feedback link">
               <Button.Content>
                 <StackPanel Orientation="Horizontal">
                   <TextBlock Text="&#xE72D;" FontFamily="Segoe MDL2 Assets" FontSize="12"
                              VerticalAlignment="Center" Margin="0,0,5,0"
                              Foreground="#1C1C00"/>
                   <TextBlock Text="Feedback" VerticalAlignment="Center"/>
                 </StackPanel>
               </Button.Content>
              <Button.Background>
                <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
                  <GradientStop Color="#FFCE3A" Offset="0"/>
                  <GradientStop Color="#F5A800" Offset="1"/>
                </LinearGradientBrush>
              </Button.Background>
              <Button.Template>
                <ControlTemplate TargetType="Button">
                  <Border x:Name="FbBg" Background="{TemplateBinding Background}"
                          BorderBrush="{TemplateBinding BorderBrush}"
                          BorderThickness="{TemplateBinding BorderThickness}"
                          CornerRadius="9" Padding="{TemplateBinding Padding}">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                  </Border>
                  <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True">
                      <Setter TargetName="FbBg" Property="Background" Value="#EDBA00"/>
                      <Setter TargetName="FbBg" Property="BorderBrush" Value="#B8860B"/>
                    </Trigger>
                    <Trigger Property="IsPressed" Value="True">
                      <Setter TargetName="FbBg" Property="Background" Value="#CC9900"/>
                    </Trigger>
                  </ControlTemplate.Triggers>
                </ControlTemplate>
              </Button.Template>
            </Button>
            <Button x:Name="BtnSessionNotes"
                    Width="30" Height="30" Margin="2,0,4,0"
                    Background="Transparent" BorderThickness="0" Cursor="Hand"
                    VerticalAlignment="Center"
                    ToolTip="Session Notes  (in-memory only — cleared on close)">
              <TextBlock Text="&#x1F5D2;" FontSize="15" Foreground="#B3FFB3"
                         VerticalAlignment="Center" HorizontalAlignment="Center"/>
            </Button>
            <Button x:Name="BtnAuthCancel" Style="{StaticResource StopBtn}"
                    Visibility="Collapsed" Width="114" FontSize="11">
              <Button.Content>
                <StackPanel Orientation="Horizontal">
                  <TextBlock Text="&#xE71A;" FontFamily="Segoe MDL2 Assets" FontSize="12"
                             VerticalAlignment="Center" Margin="0,0,5,0" Foreground="White"/>
                  <TextBlock Text="Cancel Auth" VerticalAlignment="Center"/>
                </StackPanel>
              </Button.Content>
            </Button>
          </StackPanel>


        </Grid>
      </Border>

      <!-- Author attribution strip -->
      <Border DockPanel.Dock="Top" Padding="16,5" Background="#F8FAFC">
        <Border.BorderBrush><SolidColorBrush Color="#20000000"/></Border.BorderBrush>
        <Border.BorderThickness>0,0,0,1</Border.BorderThickness>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
          <TextBlock Text="Author: " FontSize="10.5" Foreground="#6B7280" VerticalAlignment="Center" FontWeight="Normal"/>
          <TextBlock x:Name="TxtAuthorLink" Text="Satish Singhi" FontSize="10.5" Foreground="#0A66C2" FontWeight="SemiBold"
                     VerticalAlignment="Center" Cursor="Hand" TextDecorations="Underline"
                     ToolTip="https://www.linkedin.com/in/satish-singhi-791163167/"/>
          <TextBlock Text="  &#x1F517;" FontSize="10" Foreground="#0A66C2" VerticalAlignment="Center"/>
        </StackPanel>
      </Border>

      <!-- Notification strip -->
      <Border x:Name="NotifStrip" DockPanel.Dock="Top" Visibility="Collapsed"
              Padding="16,8" Background="#FFFBEB">
        <Border.BorderBrush><SolidColorBrush Color="#40F59E0B"/></Border.BorderBrush>
        <Border.BorderThickness>0,0,0,1</Border.BorderThickness>
        <TextBlock x:Name="TxtNotif" Foreground="#F59E0B" FontSize="12" TextWrapping="Wrap"/>
      </Border>

      <!-- PIM role strip -->
      <Border x:Name="PimStrip" DockPanel.Dock="Top" Visibility="Collapsed"
              Padding="18,0" Background="#EFF7EF">
        <Border.BorderBrush><SolidColorBrush Color="#0FFFFFFF"/></Border.BorderBrush>
        <Border.BorderThickness>0,0,0,1</Border.BorderThickness>
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <TextBlock Grid.Column="0" Text="&#x1F511;  Active PIM Roles"
                     Foreground="#14532D" FontSize="10" FontWeight="Bold"
                     VerticalAlignment="Center" Margin="0,0,12,0"/>
          <ScrollViewer Grid.Column="1" HorizontalScrollBarVisibility="Auto"
                        VerticalScrollBarVisibility="Disabled" Margin="0,5,0,5">
            <StackPanel x:Name="SpPimRoles" Orientation="Horizontal"/>
          </ScrollViewer>
          <Button x:Name="BtnPimRefresh" Grid.Column="2" Content="&#x21BB;"
                  Style="{StaticResource GhostBtn}"
                  Foreground="#14532D" FontSize="14" Padding="6,4"
                  ToolTip="Refresh PIM roles"/>
        </Grid>
      </Border>

      <!-- LOG PANE (bottom docked, resizable) -->
      <Grid DockPanel.Dock="Bottom">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto" MinHeight="0"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Padding="12,5" Cursor="SizeNS">
          <Border.Background>
            <LinearGradientBrush StartPoint="0,0" EndPoint="0,1">
              <GradientStop Color="#D4EDDA" Offset="0"/>
              <GradientStop Color="#C8E6C9" Offset="1"/>
            </LinearGradientBrush>
          </Border.Background>
          <Border.BorderBrush><SolidColorBrush Color="#80B0D0B0"/></Border.BorderBrush>
          <Border.BorderThickness>0,1,0,0</Border.BorderThickness>
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
              <Border Width="6" Height="6" CornerRadius="3" Background="#166534"
                      Margin="0,0,8,0" VerticalAlignment="Center">
                <Border.Effect>
                  <DropShadowEffect BlurRadius="4" ShadowDepth="0" Color="#22C55E" Opacity="0.8"/>
                </Border.Effect>
              </Border>
              <TextBlock Text="VERBOSE LOG" Foreground="#132210" FontSize="10"
                         FontWeight="Bold" FontFamily="Consolas" VerticalAlignment="Center"/>
            </StackPanel>
            <StackPanel Grid.Column="2" Orientation="Horizontal">
              <Button x:Name="BtnLogClear" Content="Clear" Style="{StaticResource GhostBtn}"
                      Foreground="#132210" FontSize="10" Padding="6,3" Margin="0,0,4,0"/>
              <Button x:Name="BtnLogToggle" Content="&#x25BC;  Hide"
                      Style="{StaticResource GhostBtn}"
                      Foreground="#132210" FontSize="10" Padding="6,3"/>
            </StackPanel>
          </Grid>
        </Border>

        <Border Grid.Row="1" x:Name="LogPane" Background="#0D1824"
                Height="$($DevConfig.VerbosePaneDefaultHeight)">
          <RichTextBox x:Name="RtbLog" Background="Transparent" BorderThickness="0"
                       IsReadOnly="True" FontFamily="Consolas" FontSize="11"
                       Foreground="#9DB8D8" Padding="14,8"
                       VerticalScrollBarVisibility="Auto"
                       HorizontalScrollBarVisibility="Disabled"
                       IsDocumentEnabled="True"/>
        </Border>

        <GridSplitter Grid.Row="0" Height="5" HorizontalAlignment="Stretch"
                      Background="Transparent" Cursor="SizeNS" VerticalAlignment="Bottom"
                      ResizeDirection="Rows" ResizeBehavior="PreviousAndNext"/>
      </Grid>

      <!-- MAIN CONTENT AREA -->
      <Grid Margin="16,12,16,12">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="224"/>
          <ColumnDefinition Width="12"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <!-- Left nav -->
        <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto"
                      HorizontalScrollBarVisibility="Disabled" PanningMode="VerticalOnly">
          <StackPanel>
            <TextBlock Text="OPERATIONS" Foreground="#0F1F14" FontSize="9.5"
                       FontWeight="Black" Margin="10,0,0,8"/>

            <!-- Group Management blade -->
            <Border Background="#F5FBF5" BorderBrush="#B0D0B0"
                    BorderThickness="1" CornerRadius="10" Margin="0,0,0,6">
              <StackPanel>
                <Button x:Name="BtnBladeGroupMgmt" Background="#E8F5E8"
                        BorderThickness="0" Cursor="Hand" Padding="10,10">
                  <Button.Template>
                    <ControlTemplate TargetType="Button">
                      <Border Background="{TemplateBinding Background}"
                              CornerRadius="9,9,0,0" Padding="{TemplateBinding Padding}">
                        <ContentPresenter/>
                      </Border>
                      <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True">
                          <Setter Property="Background" Value="#D4EDDA"/>
                        </Trigger>
                      </ControlTemplate.Triggers>
                    </ControlTemplate>
                  </Button.Template>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0" Orientation="Horizontal">
                      <TextBlock Text="&#x1F5C2;" FontSize="14" VerticalAlignment="Center" Margin="0,0,8,0"/>
                      <TextBlock Text="Group Management" FontWeight="Bold" FontSize="12"
                                 Foreground="#166534" VerticalAlignment="Center"/>
                    </StackPanel>
                    <TextBlock x:Name="TxtBladeArrow" Grid.Column="1" Text="&#x2212;"
                               FontSize="15" FontWeight="Bold" Foreground="#2A7A4A" VerticalAlignment="Center"/>
                  </Grid>
                </Button>
                <Border x:Name="BladeDivider" Height="1"><Border.Background><SolidColorBrush Color="#B0D0B0"/></Border.Background>
                </Border>
                <StackPanel x:Name="BladeGroupMgmtContent" Margin="6,5,6,6">

                  <Button x:Name="BtnOpSearch" Style="{StaticResource OpTileBtn}" Margin="0,0,0,1">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x1F50D;" FontSize="13" Foreground="#0EA5E9" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Search Entra Objects" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="Users · Devices · Groups" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>

                  <Button x:Name="BtnOpListMembers" Style="{StaticResource OpTileBtn}" Margin="0,0,0,1">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x1F4CB;" FontSize="13" Foreground="#8B5CF6" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="List Group Members" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="Users · Devices · Groups" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>


                  <Button x:Name="BtnOpObjectMembership" Style="{StaticResource OpTileBtn}" Margin="0,0,0,1">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x1F517;" FontSize="13" Foreground="#D97706" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Object Membership" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="Transitive group memberships" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>
                  <Button x:Name="BtnOpFindGroupsByOwners" Style="{StaticResource OpTileBtn}" Margin="0,0,0,1">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x1F464;" FontSize="13" Foreground="#7C3AED" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Find Groups by Owners" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="Groups owned by specified UPNs" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>
                  <Button x:Name="BtnOpCreate" Style="{StaticResource OpTileBtn}" Margin="0,0,0,1">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x2295;" FontSize="14" Foreground="#22C55E" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Create Group" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="Security or M365" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>

                  <Button x:Name="BtnOpAdd" Style="{StaticResource OpTileBtn}" Margin="0,0,0,1">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x2295;" FontSize="14" Foreground="#3B82F6" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Add Members" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="Users &#xB7; Groups &#xB7; Devices" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>

                  <Button x:Name="BtnOpRemove" Style="{StaticResource OpTileBtn}" Margin="0,0,0,1">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x2296;" FontSize="14" Foreground="#EF4444" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Remove Members" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="Users &#xB7; Groups &#xB7; Devices" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>

                  <Button x:Name="BtnOpExport" Style="{StaticResource OpTileBtn}" Margin="0,0,0,1">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x2191;" FontSize="15" Foreground="#F59E0B" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Export Members" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="To XLSX file" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>

                  <Button x:Name="BtnOpDynamic" Style="{StaticResource OpTileBtn}" Margin="0,0,0,1">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x26A1;" FontSize="13" Foreground="#8B5CF6" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Set Dynamic Query" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="MembershipRule string" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>

                  <Button x:Name="BtnOpRename" Style="{StaticResource OpTileBtn}" Margin="0,0,0,1">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x270F;" FontSize="13" Foreground="#F59E0B" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Rename Group" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="Change display name" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>

                  <Button x:Name="BtnOpOwner" Style="{StaticResource OpTileBtn}" Margin="0,0,0,1">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x2299;" FontSize="14" Foreground="#3B82F6" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Set Group Owner" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="Assign owner UPN(s)" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>

                  <Button x:Name="BtnOpUserDevices" Style="{StaticResource OpTileBtn}" Margin="0,0,0,1">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x25A3;" FontSize="13" Foreground="#6366F1" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Add User Devices" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="By UPN &#xB7; Platform &#xB7; Ownership" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>

                  <Button x:Name="BtnOpFindCommon" Style="{StaticResource OpTileBtn}">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x2229;" FontSize="15" Foreground="#0EA5E9" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Find Common Groups" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="Groups shared by all inputs" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>

                  <Button x:Name="BtnOpFindDistinct" Style="{StaticResource OpTileBtn}">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x2296;" FontSize="15" Foreground="#7C3AED" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Find Distinct Groups" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="Groups unique to each input" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>

                  <Button x:Name="BtnOpCompareGroups" Style="{StaticResource OpTileBtn}" Margin="0,0,0,1">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x21D4;" FontSize="13" Foreground="#D97706" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Compare Groups" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="Membership across groups" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>

                  <Button x:Name="BtnOpRemoveGroups" Style="{StaticResource OpTileBtn}" Margin="0,0,0,1">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x1F5D1;" FontSize="13" Foreground="#DC2626" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Remove Groups" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="Delete multiple Entra groups" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>

                </StackPanel>
              </StackPanel>
            </Border>

            <!-- Device Management blade -->
            <Border Background="#F5FBF5" BorderBrush="#B0D0B0"
                    BorderThickness="1" CornerRadius="10" Margin="0,0,0,6">
              <StackPanel>
                <Button x:Name="BtnBladeDevMgmt" Background="#E8F5E8"
                        BorderThickness="0" Cursor="Hand" Padding="10,10">
                  <Button.Template>
                    <ControlTemplate TargetType="Button">
                      <Border Background="{TemplateBinding Background}"
                              CornerRadius="9,9,0,0" Padding="{TemplateBinding Padding}">
                        <ContentPresenter/>
                      </Border>
                      <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True">
                          <Setter Property="Background" Value="#D4EDDA"/>
                        </Trigger>
                      </ControlTemplate.Triggers>
                    </ControlTemplate>
                  </Button.Template>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0" Orientation="Horizontal">
                      <TextBlock Text="&#x1F4BB;" FontSize="14" VerticalAlignment="Center" Margin="0,0,8,0"/>
                      <TextBlock Text="Device Management" FontWeight="Bold" FontSize="12"
                                 Foreground="#166534" VerticalAlignment="Center"/>
                    </StackPanel>
                    <TextBlock x:Name="TxtBladeDevArrow" Grid.Column="1" Text="&#x2212;"
                               FontSize="15" FontWeight="Bold" Foreground="#2A7A4A" VerticalAlignment="Center"/>
                  </Grid>
                </Button>
                <Border x:Name="BladeDevDivider" Height="1"><Border.Background><SolidColorBrush Color="#B0D0B0"/></Border.Background>
                </Border>
                <StackPanel x:Name="BladeDevMgmtContent" Margin="6,5,6,6">
                  <Button x:Name="BtnOpGetPolicyAssignments" Style="{StaticResource OpTileBtn}" Margin="0,0,0,1">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x1F4CB;" FontSize="13" Foreground="#7C3AED" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Get Policy Info" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="Intune policy assignment report" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>
                  <Button x:Name="BtnOpGetDeviceInfo" Style="{StaticResource OpTileBtn}">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x2139;" FontSize="14" Foreground="#14B8A6" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Get Device Info" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="Entra + Intune combined report" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>
                  <Button x:Name="BtnOpGetDiscoveredApps" Style="{StaticResource OpTileBtn}" Margin="0,1,0,0">
                    <StackPanel>
                      <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#x1F4E6;" FontSize="13" Foreground="#7C3AED" Margin="0,0,7,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Get App Info" FontWeight="SemiBold" FontSize="12" Foreground="#374151"/>
                      </StackPanel>
                      <TextBlock Text="Apps detected on managed devices" FontSize="10" Foreground="#14532D" Margin="20,2,0,0"/>
                    </StackPanel>
                  </Button>
                </StackPanel>
              </StackPanel>
            </Border>

          </StackPanel>
        </ScrollViewer>

        <!-- Separator -->
        <Border Grid.Column="1" Width="1" HorizontalAlignment="Center">
          <Border.Background><SolidColorBrush Color="#D0D8D0"/></Border.Background>
        </Border>

        <!-- Right content: op panels -->
        <ScrollViewer Grid.Column="2" VerticalScrollBarVisibility="Auto"
                      HorizontalScrollBarVisibility="Disabled">
          <Grid>

            <!-- === WELCOME panel === -->
            <Border x:Name="PanelWelcome" Style="{StaticResource Card}" Visibility="Visible">
              <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center" Margin="0,40">
                <TextBlock Text="Welcome to $($DevConfig.WindowTitle)"
                           FontSize="18" FontWeight="SemiBold"
                           Foreground="$($DevConfig.ColorPrimary)"
                           HorizontalAlignment="Center"/>
                <TextBlock Text="Connect to Microsoft Graph, then select an operation from the left panel."
                           FontSize="13" Foreground="$($DevConfig.ColorTextMuted)"
                           HorizontalAlignment="Center" Margin="0,8,0,0" TextWrapping="Wrap"/>
              </StackPanel>
            </Border>

            <!-- === SEARCH ENTRA OBJECTS panel === -->
            <Border x:Name="PanelSearch" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Search Entra Objects" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,4"
                           Foreground="$($DevConfig.ColorPrimary)"/>
                <TextBlock Text="Search for Users, Devices, Security Groups and M365 Groups. Accepts bulk input: UPN, Device Name, Device ID, Group Name, or Group ID - one per line."
                           FontSize="11" Foreground="$($DevConfig.ColorTextMuted)" Margin="0,0,0,16" TextWrapping="Wrap"/>
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="SEARCH INPUT  -  one entry per line *" Style="{StaticResource FieldLabel}"/>
                  <Grid Margin="0,4,0,0">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtSearchKeyword" Grid.Column="0" Style="{StaticResource InputBox}"
                             Height="110" TextWrapping="Wrap" AcceptsReturn="True"
                             VerticalScrollBarVisibility="Auto" Tag="One entry per line: UPN, Device Name, Device ID, Group Name, or Group ID"/>
                    <Button x:Name="BtnEntraSearch" Grid.Column="2"
                            Content="&#x1F50D;  Search"
                            Style="{StaticResource PrimaryBtn}" Width="110" VerticalAlignment="Top"/>
                  </Grid>
                </StackPanel>
                <StackPanel Margin="0,0,0,16">
                  <TextBlock Text="OBJECT TYPE FILTER  (none checked = all types)" Style="{StaticResource FieldLabel}"/>
                  <WrapPanel Margin="0,6,0,0">
                    <CheckBox x:Name="ChkSrchUsers"   Content="Users"           IsChecked="True" Margin="0,0,20,4" FontSize="12"/>
                    <CheckBox x:Name="ChkSrchGetManager" Content="Get Manager" IsChecked="False" Margin="0,0,20,4" FontSize="12" ToolTip="When checked, adds Manager UPN column for User objects"/>
                    <CheckBox x:Name="ChkSrchDevices" Content="Devices"         IsChecked="True" Margin="0,0,20,4" FontSize="12"/>
                    <CheckBox x:Name="ChkSrchSG"      Content="Security Groups" IsChecked="True" Margin="0,0,20,4" FontSize="12"/>
                    <CheckBox x:Name="ChkSrchM365"    Content="M365 Groups"     IsChecked="True" Margin="0,0,0,4"  FontSize="12"/>
                  </WrapPanel>
                    <CheckBox x:Name="ChkSrchExactMatch" Content="Exact Match only" IsChecked="False" Margin="0,8,0,0" FontSize="12"/>
                </StackPanel>
                <Border x:Name="PnlSearchResults" Visibility="Collapsed">
                  <StackPanel>
                    <Grid Margin="0,0,0,8">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock x:Name="TxtSearchCount" Grid.Column="0"
                                 FontSize="11" Foreground="#6B7280" VerticalAlignment="Center"/>
                      <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <Button x:Name="BtnSearchFilterClear" Content="&#x2715;"
                                Style="{StaticResource PrimaryBtn}" ToolTip="Clear filter"
                                IsEnabled="False" Height="30" Width="30" FontSize="12"
                                Padding="0" Margin="0,0,4,0"/>
                        <TextBox x:Name="TxtSearchFilter" Width="150" Height="30" FontSize="12"
                                VerticalContentAlignment="Center" Padding="6,0" Margin="0,0,6,0"
                                IsEnabled="False" Tag="keyword filter..."/>
                        <Button x:Name="BtnSearchFilter" Content="&#x1F50D; Filter"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,12,0"/>
                        <Button x:Name="BtnSearchCopyValue"
                                Content="Copy Value"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnSearchCopyRow"
                                Content="Copy Row"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnSearchCopyAll"
                                Content="Copy All"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnSearchExportXlsx"
                                Content="Export XLSX"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4"/>
                      </StackPanel>
                    </Grid>
                    <DataGrid x:Name="DgSearchResults"
                              AutoGenerateColumns="False" IsReadOnly="True"
                              SelectionMode="Extended" SelectionUnit="CellOrRowHeader"
                              HeadersVisibility="Column"
                              GridLinesVisibility="All"
                              VerticalGridLinesBrush="#555555"
                              HorizontalGridLinesBrush="#D0D0D0"
                              RowBackground="White" AlternatingRowBackground="#F5FBF5"
                              BorderBrush="#B0D0B0" BorderThickness="1"
                              CanUserReorderColumns="False" CanUserResizeRows="False" CanUserResizeColumns="True"
                              CanUserSortColumns="True" MaxHeight="420" FontSize="12"
                              ScrollViewer.HorizontalScrollBarVisibility="Auto">
                      <DataGrid.ColumnHeaderStyle>
                        <Style TargetType="DataGridColumnHeader">
                          <Setter Property="FontWeight" Value="SemiBold"/>
                          <Setter Property="Background" Value="#E8F5E9"/>
                          <Setter Property="Foreground" Value="#14532D"/>
                          <Setter Property="Padding"    Value="8,5"/>
                          <Setter Property="BorderBrush" Value="#555555"/>
                          <Setter Property="BorderThickness" Value="0,0,1,1"/>
                        </Style>
                      </DataGrid.ColumnHeaderStyle>
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Object Name" Binding="{Binding DisplayName}" Width="*"   MinWidth="160"/>
                        <DataGridTextColumn Header="Details"     Binding="{Binding Detail}"      Width="220"/>
                        <DataGridTextColumn Header="Object Type" Binding="{Binding ObjectType}"  Width="140"/>
                      </DataGrid.Columns>
                    </DataGrid>
                  </StackPanel>
                </Border>
                <TextBlock x:Name="TxtSearchNoResults" Visibility="Collapsed"
                           Text="No matching objects found." FontSize="12"
                           Foreground="#6B7280" Margin="0,10,0,0"/>
              </StackPanel>
            </Border>

            <!-- === LIST GROUP MEMBERS panel === -->
            <Border x:Name="PanelListMembers" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="List Group Members" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,4"
                           Foreground="$($DevConfig.ColorPrimary)"/>
                <TextBlock Text="Query members from one or more groups. Enter display names or Object IDs, one per line."
                           FontSize="11" Foreground="$($DevConfig.ColorTextMuted)" Margin="0,0,0,16" TextWrapping="Wrap"/>

                <!-- Group input -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="GROUPS  *  (one per line  —  display name or Object ID)" Style="{StaticResource FieldLabel}"/>
                  <Grid Margin="0,4,0,0">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtLMGroupInput" Grid.Column="0" Style="{StaticResource InputBox}"
                             Height="80" TextWrapping="Wrap" AcceptsReturn="True"
                             VerticalScrollBarVisibility="Auto"
                             Tag="SG-Marketing-All&#10;a1b2c3d4-0000-0000-0000-000000000000&#10;IT-Devices"/>
                    <Button x:Name="BtnLMQuery" Grid.Column="2"
                            Content="&#x1F4CB;  List Members"
                            Style="{StaticResource PrimaryBtn}" Width="130" VerticalAlignment="Top"/>
                  </Grid>
                </StackPanel>

                <!-- Output type filter -->
                <StackPanel Margin="0,0,0,16">
                  <TextBlock Text="OUTPUT TYPE FILTER  (none checked = all types)" Style="{StaticResource FieldLabel}"/>
                  <WrapPanel Margin="0,6,0,0">
                    <CheckBox x:Name="ChkLMUsers"   Content="Users"                     IsChecked="True" Margin="0,0,20,4" FontSize="12"/>
                    <CheckBox x:Name="ChkLMDevices" Content="Devices"                   IsChecked="True" Margin="0,0,20,4" FontSize="12"/>
                    <CheckBox x:Name="ChkLMGroups"  Content="Security &amp; M365 Groups" IsChecked="True" Margin="0,0,0,4"  FontSize="12"/>
                  </WrapPanel>
                </StackPanel>

                <!-- Progress overlay (shown while querying) -->
                <Border x:Name="PnlLMProgress" Visibility="Collapsed"
                        Background="#1A2E1A" CornerRadius="6" Padding="16,12" Margin="0,0,0,12">
                  <StackPanel>
                    <TextBlock x:Name="TxtLMProgressMsg" Text="Querying groups..."
                               Foreground="#A3E4C0" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <ProgressBar x:Name="PbLMProgress" Height="6" IsIndeterminate="True"
                                 Background="#2A4A2A" Foreground="#22C55E" BorderThickness="0"/>
                    <TextBlock x:Name="TxtLMProgressDetail" Text=""
                               Foreground="#6BB88B" FontSize="11" Margin="0,5,0,0"/>
                  </StackPanel>
                </Border>

                <!-- Results area -->
                <Border x:Name="PnlLMResults" Visibility="Collapsed">
                  <StackPanel>
                    <Grid Margin="0,0,0,8">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock x:Name="TxtLMCount" Grid.Column="0"
                                 FontSize="11" Foreground="#6B7280" VerticalAlignment="Center"/>
                      <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <Button x:Name="BtnLMFilterClear" Content="&#x2715;"
                                Style="{StaticResource PrimaryBtn}" ToolTip="Clear filter"
                                IsEnabled="False" Height="30" Width="30" FontSize="12"
                                Padding="0" Margin="0,0,4,0"/>
                        <TextBox x:Name="TxtLMFilter" Width="150" Height="30" FontSize="12"
                                VerticalContentAlignment="Center" Padding="6,0" Margin="0,0,6,0"
                                IsEnabled="False" Tag="keyword filter..."/>
                        <Button x:Name="BtnLMFilter" Content="&#x1F50D; Filter"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,12,0"/>
                        <Button x:Name="BtnLMCopyValue" Content="Copy Value"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnLMCopyRow" Content="Copy Row"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnLMCopyAll" Content="Copy All"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnLMExportXlsx" Content="Export XLSX"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4"/>
                      </StackPanel>
                    </Grid>
                    <DataGrid x:Name="DgLMResults"
                              AutoGenerateColumns="False" IsReadOnly="True"
                              SelectionMode="Extended" SelectionUnit="CellOrRowHeader"
                              HeadersVisibility="Column"
                              GridLinesVisibility="All"
                              VerticalGridLinesBrush="#555555"
                              HorizontalGridLinesBrush="#D0D0D0"
                              RowBackground="White" AlternatingRowBackground="#F5FBF5"
                              BorderBrush="#B0D0B0" BorderThickness="1"
                              CanUserReorderColumns="False" CanUserResizeRows="False" CanUserResizeColumns="True"
                              CanUserSortColumns="True" MaxHeight="420" FontSize="12"
                              ScrollViewer.HorizontalScrollBarVisibility="Auto">
                      <DataGrid.ColumnHeaderStyle>
                        <Style TargetType="DataGridColumnHeader">
                          <Setter Property="FontWeight" Value="SemiBold"/>
                          <Setter Property="Background" Value="#E8F5E9"/>
                          <Setter Property="Foreground" Value="#14532D"/>
                          <Setter Property="Padding"    Value="8,5"/>
                          <Setter Property="BorderBrush" Value="#555555"/>
                          <Setter Property="BorderThickness" Value="0,0,1,1"/>
                        </Style>
                      </DataGrid.ColumnHeaderStyle>
                      <!-- Columns built dynamically in code-behind based on result schema -->
                    </DataGrid>
                  </StackPanel>
                </Border>
                <TextBlock x:Name="TxtLMNoResults" Visibility="Collapsed"
                           Text="No members found matching the current filters."
                           FontSize="12" Foreground="#6B7280" Margin="0,10,0,0"/>
              </StackPanel>
            </Border>


            <!-- === OBJECT MEMBERSHIP panel === -->
            <Border x:Name="PanelObjectMembership" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Object Membership" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,4"
                           Foreground="$($DevConfig.ColorPrimary)"/>
                <TextBlock Text="Resolve transitive group memberships for users, devices, or groups. Enter identifiers one per line."
                           FontSize="11" Foreground="$($DevConfig.ColorTextMuted)" Margin="0,0,0,14" TextWrapping="Wrap"/>

                <!-- Object type -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="OBJECT TYPE" Style="{StaticResource FieldLabel}"/>
                  <ComboBox x:Name="CmbOMObjectType" Style="{StaticResource StyledCombo}"
                            Width="260" HorizontalAlignment="Left" Margin="0,4,0,0">
                    <ComboBoxItem Content="Users (UPN)" IsSelected="True"/>
                    <ComboBoxItem Content="Devices (Name)"/>
                    <ComboBoxItem Content="Devices (ID)"/>
                    <ComboBoxItem Content="Security Groups"/>
                    <ComboBoxItem Content="M365 Groups"/>
                  </ComboBox>
                </StackPanel>

                <!-- Input list -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock x:Name="TxtOMInputLabel" Text="USER UPNs  -  one per line"
                             Style="{StaticResource FieldLabel}"/>
                  <Grid Margin="0,4,0,0">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtOMInputList" Grid.Column="0" Style="{StaticResource InputBox}"
                             Height="110" TextWrapping="Wrap" AcceptsReturn="True"
                             VerticalScrollBarVisibility="Auto" Tag="user@domain.com"/>
                    <Button x:Name="BtnOMRun" Grid.Column="2"
                            Content="&#x25B6;  Get Memberships"
                            Style="{StaticResource PrimaryBtn}" Width="180" VerticalAlignment="Top"/>
                  </Grid>
                </StackPanel>

                <!-- Progress -->
                <Border x:Name="PnlOMProgress" Visibility="Collapsed"
                        Background="#1A2E1A" CornerRadius="6" Padding="16,12" Margin="0,0,0,12">
                  <StackPanel>
                    <TextBlock x:Name="TxtOMProgressMsg" Text="Resolving objects..."
                               Foreground="#A3E4C0" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <ProgressBar x:Name="PbOMProgress" Height="6" IsIndeterminate="True"
                                 Background="#2A4A2A" Foreground="#22C55E" BorderThickness="0"/>
                    <TextBlock x:Name="TxtOMProgressDetail" Text=""
                               Foreground="#6BB88B" FontSize="11" Margin="0,5,0,0"/>
                  </StackPanel>
                </Border>

                <!-- Results -->
                <Border x:Name="PnlOMResults" Visibility="Collapsed">
                  <StackPanel>
                    <Grid Margin="0,0,0,8">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock x:Name="TxtOMCount" Grid.Column="0"
                                 FontSize="11" Foreground="#6B7280" VerticalAlignment="Center"/>
                      <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <Button x:Name="BtnOMFilterClear" Content="&#x2715;"
                                Style="{StaticResource PrimaryBtn}" ToolTip="Clear filter"
                                IsEnabled="False" Height="30" Width="30" FontSize="12"
                                Padding="0" Margin="0,0,4,0"/>
                        <TextBox x:Name="TxtOMFilter" Width="150" Height="30" FontSize="12"
                                VerticalContentAlignment="Center" Padding="6,0" Margin="0,0,6,0"
                                IsEnabled="False" Tag="keyword filter..."/>
                        <Button x:Name="BtnOMFilter" Content="&#x1F50D; Filter"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,12,0"/>
                        <Button x:Name="BtnOMCopyValue" Content="Copy Value"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnOMCopyRow" Content="Copy Row"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnOMCopyAll" Content="Copy All"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnOMExportXlsx" Content="Export XLSX"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4"/>
                      </StackPanel>
                    </Grid>
                    <DataGrid x:Name="DgOMResults"
                              AutoGenerateColumns="False" IsReadOnly="True"
                              SelectionMode="Extended" SelectionUnit="CellOrRowHeader"
                              HeadersVisibility="Column"
                              GridLinesVisibility="All"
                              VerticalGridLinesBrush="#555555"
                              HorizontalGridLinesBrush="#D0D0D0"
                              RowBackground="White" AlternatingRowBackground="#F5FBF5"
                              BorderBrush="#B0D0B0" BorderThickness="1"
                              CanUserReorderColumns="False" CanUserResizeRows="False" CanUserResizeColumns="True"
                              CanUserSortColumns="True" MaxHeight="420" FontSize="12"
                              ScrollViewer.HorizontalScrollBarVisibility="Auto">
                      <DataGrid.ColumnHeaderStyle>
                        <Style TargetType="DataGridColumnHeader">
                          <Setter Property="FontWeight" Value="SemiBold"/>
                          <Setter Property="Background" Value="#E8F5E9"/>
                          <Setter Property="Foreground" Value="#14532D"/>
                          <Setter Property="Padding"    Value="8,5"/>
                          <Setter Property="BorderBrush" Value="#555555"/>
                          <Setter Property="BorderThickness" Value="0,0,1,1"/>
                        </Style>
                      </DataGrid.ColumnHeaderStyle>
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Input Object"     Binding="{Binding InputObject}"    Width="2*" MinWidth="150"/>
                        <DataGridTextColumn Header="Object Type"      Binding="{Binding ObjectType}"     Width="*"  MinWidth="80"/>
                        <DataGridTextColumn Header="Group Name"       Binding="{Binding GroupName}"      Width="2*" MinWidth="160"/>
                        <DataGridTextColumn Header="Group ID"         Binding="{Binding GroupId}"        Width="2*" MinWidth="130"/>
                        <DataGridTextColumn Header="Group Type"       Binding="{Binding GroupType}"      Width="*"  MinWidth="100"/>
                        <DataGridTextColumn Header="Description"      Binding="{Binding Description}"    Width="2*" MinWidth="150"/>
                        <DataGridTextColumn Header="Membership Type"  Binding="{Binding MembershipType}" Width="*"  MinWidth="100"/>
                      </DataGrid.Columns>
                    </DataGrid>
                  </StackPanel>
                </Border>
                <TextBlock x:Name="TxtOMNoResults" Visibility="Collapsed"
                           Text="No group memberships found for the provided input."
                           FontSize="12" Foreground="#6B7280" Margin="0,12,0,0" TextWrapping="Wrap"/>
              </StackPanel>
            </Border>


            <!-- === FIND GROUPS BY OWNERS panel === -->
            <Border x:Name="PanelFindGroupsByOwners" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Find Groups by Owners" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,4"
                           Foreground="$($DevConfig.ColorPrimary)"/>
                <TextBlock Text="Enter owner UPNs (one per line) to find all groups they own."
                           FontSize="11" Foreground="$($DevConfig.ColorTextMuted)" Margin="0,0,0,14" TextWrapping="Wrap"/>

                <!-- Input list -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="OWNER UPNs  -  one per line" Style="{StaticResource FieldLabel}"/>
                  <Grid Margin="0,4,0,0">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtFGOInputList" Grid.Column="0" Style="{StaticResource InputBox}"
                             Height="110" TextWrapping="Wrap" AcceptsReturn="True"
                             VerticalScrollBarVisibility="Auto" Tag="user@domain.com"/>
                    <Button x:Name="BtnFGORun" Grid.Column="2"
                            Content="&#x25B6;  Find Groups"
                            Style="{StaticResource PrimaryBtn}" Width="180" VerticalAlignment="Top"/>
                  </Grid>
                </StackPanel>

                <!-- Progress -->
                <Border x:Name="PnlFGOProgress" Visibility="Collapsed"
                        Background="#1A2E1A" CornerRadius="6" Padding="16,12" Margin="0,0,0,12">
                  <StackPanel>
                    <TextBlock x:Name="TxtFGOProgressMsg" Text="Resolving owners..."
                               Foreground="#A3E4C0" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <ProgressBar x:Name="PbFGOProgress" Height="6" IsIndeterminate="True"
                                 Background="#2A4A2A" Foreground="#22C55E" BorderThickness="0"/>
                    <TextBlock x:Name="TxtFGOProgressDetail" Text=""
                               Foreground="#6BB88B" FontSize="11" Margin="0,5,0,0"/>
                  </StackPanel>
                </Border>

                <!-- Results -->
                <Border x:Name="PnlFGOResults" Visibility="Collapsed">
                  <StackPanel>
                    <Grid Margin="0,0,0,8">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock x:Name="TxtFGOCount" Grid.Column="0"
                                 FontSize="11" Foreground="#6B7280" VerticalAlignment="Center"/>
                      <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <Button x:Name="BtnFGOFilterClear" Content="&#x2715;"
                                Style="{StaticResource PrimaryBtn}" ToolTip="Clear filter"
                                IsEnabled="False" Height="30" Width="30" FontSize="12"
                                Padding="0" Margin="0,0,4,0"/>
                        <TextBox x:Name="TxtFGOFilter" Width="150" Height="30" FontSize="12"
                                VerticalContentAlignment="Center" Padding="6,0" Margin="0,0,6,0"
                                IsEnabled="False" Tag="keyword filter..."/>
                        <Button x:Name="BtnFGOFilter" Content="&#x1F50D; Filter"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,12,0"/>
                        <Button x:Name="BtnFGOCopyValue" Content="Copy Value"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnFGOCopyRow" Content="Copy Row"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnFGOCopyAll" Content="Copy All"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnFGOExportXlsx" Content="Export XLSX"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4"/>
                      </StackPanel>
                    </Grid>
                    <DataGrid x:Name="DgFGOResults"
                              AutoGenerateColumns="False" IsReadOnly="True"
                              SelectionMode="Extended" SelectionUnit="CellOrRowHeader"
                              HeadersVisibility="Column"
                              GridLinesVisibility="All"
                              VerticalGridLinesBrush="#555555"
                              HorizontalGridLinesBrush="#D0D0D0"
                              RowBackground="White" AlternatingRowBackground="#F5FBF5"
                              BorderBrush="#B0D0B0" BorderThickness="1"
                              CanUserReorderColumns="False" CanUserResizeRows="False" CanUserResizeColumns="True"
                              CanUserSortColumns="True" MaxHeight="420" FontSize="12"
                              ScrollViewer.HorizontalScrollBarVisibility="Auto">
                      <DataGrid.ColumnHeaderStyle>
                        <Style TargetType="DataGridColumnHeader">
                          <Setter Property="FontWeight" Value="SemiBold"/>
                          <Setter Property="Background" Value="#E8F5E9"/>
                          <Setter Property="Foreground" Value="#14532D"/>
                          <Setter Property="Padding"    Value="8,5"/>
                          <Setter Property="BorderBrush" Value="#555555"/>
                          <Setter Property="BorderThickness" Value="0,0,1,1"/>
                        </Style>
                      </DataGrid.ColumnHeaderStyle>
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Owner"             Binding="{Binding Owner}"          Width="2*" MinWidth="160"/>
                        <DataGridTextColumn Header="Group Name"        Binding="{Binding GroupName}"       Width="2*" MinWidth="160"/>
                        <DataGridTextColumn Header="Group ID"          Binding="{Binding GroupId}"         Width="2*" MinWidth="130"/>
                        <DataGridTextColumn Header="Group Type"        Binding="{Binding GroupType}"       Width="*"  MinWidth="100"/>
                        <DataGridTextColumn Header="Membership Type"   Binding="{Binding MembershipType}"  Width="*"  MinWidth="110"/>
                        <DataGridTextColumn Header="Member Count"      Binding="{Binding MemberCount}"     Width="*"  MinWidth="90"/>
                      </DataGrid.Columns>
                    </DataGrid>
                  </StackPanel>
                </Border>
                <TextBlock x:Name="TxtFGONoResults" Visibility="Collapsed"
                           Text="No groups found matching the criteria."
                           FontSize="12" Foreground="#6B7280" Margin="0,12,0,0" TextWrapping="Wrap"/>
              </StackPanel>
            </Border>

            <!-- === CREATE GROUP panel === -->
            <Border x:Name="PanelCreate" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Create Group" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,12"
                           Foreground="$($DevConfig.ColorPrimary)"/>

                <!-- Group type -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="GROUP TYPE" Style="{StaticResource FieldLabel}"/>
                  <ComboBox x:Name="CmbCreateType" Style="{StaticResource StyledCombo}" Width="220"
                            HorizontalAlignment="Left">
                    <ComboBoxItem Content="Security" IsSelected="True"/>
                    <ComboBoxItem Content="Microsoft 365 (M365)"/>
                  </ComboBox>
                </StackPanel>

                <!-- Display name -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="DISPLAY NAME *" Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="TxtCreateName" Style="{StaticResource InputBox}"
                           Tag="e.g. SG-Marketing-All"/>
                </StackPanel>

                <!-- M365-only fields (email address + sensitivity label) -->
                <StackPanel x:Name="PnlMailNick" Margin="0,0,0,0">

                  <!-- Group Email Address -->
                  <StackPanel Margin="0,0,0,12">
                    <TextBlock Text="GROUP EMAIL ADDRESS *" Style="{StaticResource FieldLabel}"/>
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBox x:Name="TxtCreateMailNick" Grid.Column="0"
                               Style="{StaticResource InputBox}"
                               Tag="mail-nickname (local part only)"/>
                      <Border Grid.Column="1" Background="$($DevConfig.ColorBackground)"
                              BorderBrush="$($DevConfig.ColorBorder)" BorderThickness="0,1,1,1"
                              CornerRadius="0,4,4,0" Padding="10,0,10,0">
                        <TextBlock x:Name="TxtMailDomainSuffix" VerticalAlignment="Center"
                                   FontSize="13" Foreground="$($DevConfig.ColorTextMuted)"
                                   Text="@"/>
                      </Border>
                    </Grid>
                  </StackPanel>

                  <!-- Sensitivity Label (populated at runtime; hidden if tenant has none) -->
                  <StackPanel x:Name="PnlSensitivityLabel" Margin="0,0,0,12" Visibility="Collapsed">
                    <TextBlock Text="SENSITIVITY LABEL (optional)" Style="{StaticResource FieldLabel}"/>
                    <ComboBox x:Name="CmbSensitivityLabel" Style="{StaticResource StyledCombo}"
                              HorizontalAlignment="Left" Width="320">
                      <ComboBoxItem Content="(None)" IsSelected="True" Tag=""/>
                    </ComboBox>
                  </StackPanel>

                </StackPanel>

                <!-- Description -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="DESCRIPTION" Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="TxtCreateDesc" Style="{StaticResource InputBox}"
                           Height="60" TextWrapping="Wrap" AcceptsReturn="True"
                           VerticalScrollBarVisibility="Auto"
                           Tag="Optional description"/>
                </StackPanel>

                <!-- Initial owner UPNs -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="INITIAL OWNER UPNs (optional  -  one per line)"
                             Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="TxtCreateOwner" Style="{StaticResource InputBox}"
                           Height="60" TextWrapping="Wrap" AcceptsReturn="True"
                           VerticalScrollBarVisibility="Auto"
                           Tag="owner1@domain.com&#10;owner2@domain.com"/>
                </StackPanel>

                <!-- Initial members -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="INITIAL MEMBERS (optional  -  one per line, UPNs / Group names or IDs / Device names or IDs)"
                             Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="TxtCreateMembers" Style="{StaticResource InputBox}"
                           Height="80" TextWrapping="Wrap" AcceptsReturn="True"
                           VerticalScrollBarVisibility="Auto"
                           Tag="user@domain.com&#10;GroupName or GroupObjectID&#10;DeviceName or DeviceObjectID"/>
                </StackPanel>

                <!-- Dynamic query -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="DYNAMIC MEMBERSHIP RULE (optional  -  converts group to dynamic)"
                             Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="TxtCreateDynamic" Style="{StaticResource InputBox}"
                           Tag='e.g. (user.department -eq "Marketing")'/>
                </StackPanel>

                <!-- Administrative Unit (optional) -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="ADMINISTRATIVE UNIT (optional)" Style="{StaticResource FieldLabel}"/>
                  <ComboBox x:Name="CmbCreateAU" Style="{StaticResource StyledCombo}"
                            HorizontalAlignment="Left" Width="360">
                    <ComboBoxItem Content="(None)" IsSelected="True" Tag=""/>
                  </ComboBox>
                  <TextBlock x:Name="TxtAULoading" Text="Loading Administrative Units..."
                             FontSize="10" Foreground="Gray" Margin="0,2,0,0" Visibility="Collapsed"/>
                </StackPanel>

                <Button Margin="4,4,0,0" x:Name="BtnCreateValidate" Content="Validate &amp; Preview &#x2192;"
                        Style="{StaticResource PrimaryBtn}" HorizontalAlignment="Left" Width="180"/>
              </StackPanel>
            </Border>

            <!-- === ADD / REMOVE MEMBERS panel (shared) === -->
            <Border x:Name="PanelMembership" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock x:Name="TxtMembershipTitle" Text="Add Members"
                           FontSize="15" FontWeight="SemiBold"
                           Foreground="$($DevConfig.ColorPrimary)"/>

                <!-- Group picker -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="TARGET GROUP" Style="{StaticResource FieldLabel}"/>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtGroupSearch" Grid.Column="0" Style="{StaticResource InputBox}"
                             Tag="Type group name, paste Object ID, or enter display name"/>
                    <Button x:Name="BtnGroupSearch" Grid.Column="2" Content="Search"
                            Style="{StaticResource OutlineBtn}" Width="70" Margin="0,0,8,0"/>
                    <Button x:Name="BtnGroupById" Grid.Column="4" Content="Use as ID"
                            Style="{StaticResource OutlineBtn}" Width="80"/>
                  </Grid>

                  <!-- Search results -->
                  <Border Margin="0,4,0,0" BorderBrush="$($DevConfig.ColorBorder)"
                          BorderThickness="1" CornerRadius="5"
                          x:Name="GroupSearchResults" Visibility="Collapsed" MaxHeight="150">
                    <ListBox x:Name="LstGroupResults" Background="$($DevConfig.ColorSurface)"
                             BorderThickness="0" FontSize="12"/>
                  </Border>

                  <!-- Selected group display -->
                  <Border x:Name="SelectedGroupBadge" Margin="0,6,0,0" Visibility="Collapsed"
                          Background="#EEF5FF" BorderBrush="$($DevConfig.ColorPrimary)"
                          BorderThickness="1" CornerRadius="5" Padding="10,6">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Text="Selected:" FontSize="11"
                                 Foreground="$($DevConfig.ColorTextMuted)"/>
                      <TextBlock x:Name="TxtSelectedGroupName" FontSize="12" FontWeight="SemiBold"
                                 Foreground="$($DevConfig.ColorPrimary)"/>
                      <TextBlock x:Name="TxtSelectedGroupId" FontSize="10"
                                 Foreground="$($DevConfig.ColorTextMuted)"/>
                      <TextBlock x:Name="TxtSelectedGroupMemCount" FontSize="10"
                                 Foreground="$($DevConfig.ColorTextMuted)"/>
                    </StackPanel>
                  </Border>
                </StackPanel>

                <!-- Member input -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock x:Name="TxtMemberInputLabel"
                             Text="MEMBERS  -  one per line (User UPNs / Group names or IDs / Device names or IDs)"
                             Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="TxtMemberList" Style="{StaticResource InputBox}"
                           Height="140" TextWrapping="Wrap" AcceptsReturn="True"
                           VerticalScrollBarVisibility="Auto"
                           Tag="user@domain.com&#10;GroupName or GroupObjectID&#10;DeviceName or DeviceObjectID"/>
                </StackPanel>

                <Button Margin="4,4,0,0" x:Name="BtnMemberValidate" Content="Validate &amp; Preview &#x2192;"
                        Style="{StaticResource PrimaryBtn}" HorizontalAlignment="Left" Width="180"/>
              </StackPanel>
            </Border>

            <!-- === EXPORT panel === -->
            <Border x:Name="PanelExport" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Export Group Members to XLSX"
                           FontSize="15" FontWeight="SemiBold"
                           Foreground="$($DevConfig.ColorPrimary)"/>

                <!-- Group picker (same pattern) -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="TARGET GROUP" Style="{StaticResource FieldLabel}"/>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtExportGroupSearch" Grid.Column="0"
                             Style="{StaticResource InputBox}"
                             Tag="Type group name, paste Object ID, or enter display name"/>
                    <Button x:Name="BtnExportGroupSearch" Grid.Column="2" Content="Search"
                            Style="{StaticResource OutlineBtn}" Width="70" Margin="0,0,8,0"/>
                    <Button x:Name="BtnExportGroupById" Grid.Column="4" Content="Use as ID"
                            Style="{StaticResource OutlineBtn}" Width="80"/>
                  </Grid>
                  <Border Margin="0,4,0,0" BorderBrush="$($DevConfig.ColorBorder)"
                          BorderThickness="1" CornerRadius="5"
                          x:Name="ExportGroupSearchResults" Visibility="Collapsed" MaxHeight="150">
                    <ListBox x:Name="LstExportGroupResults" Background="$($DevConfig.ColorSurface)"
                             BorderThickness="0" FontSize="12"/>
                  </Border>
                  <Border x:Name="ExportSelectedGroupBadge" Margin="0,6,0,0" Visibility="Collapsed"
                          Background="#EEF5FF" BorderBrush="$($DevConfig.ColorPrimary)"
                          BorderThickness="1" CornerRadius="5" Padding="10,6">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Text="Selected:" FontSize="11" Foreground="$($DevConfig.ColorTextMuted)"/>
                      <TextBlock x:Name="TxtExportGroupName" FontSize="12" FontWeight="SemiBold"
                                 Foreground="$($DevConfig.ColorPrimary)"/>
                      <TextBlock x:Name="TxtExportGroupId" FontSize="10"
                                 Foreground="$($DevConfig.ColorTextMuted)"/>
                    </StackPanel>
                  </Border>
                </StackPanel>

                <!-- Output path -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="OUTPUT XLSX PATH" Style="{StaticResource FieldLabel}"/>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtExportPath" Grid.Column="0"
                             Style="{StaticResource InputBox}"
                             Tag="C:\Reports\GroupMembers.xlsx"/>
                    <Button x:Name="BtnExportBrowse" Grid.Column="2" Content="Browse&#x2026;"
                            Style="{StaticResource OutlineBtn}" Width="80"/>
                  </Grid>
                </StackPanel>

                <Button Margin="4,4,0,0" x:Name="BtnRunExport" Content="Export Members &#x2192;"
                        Style="{StaticResource AccentBtn}" HorizontalAlignment="Left" Width="160"/>
              </StackPanel>
            </Border>

            <!-- === DYNAMIC QUERY panel === -->
            <Border x:Name="PanelDynamic" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Set Dynamic Membership Query"
                           FontSize="15" FontWeight="SemiBold"
                           Foreground="$($DevConfig.ColorPrimary)"/>

                <!-- Group picker -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="TARGET GROUP" Style="{StaticResource FieldLabel}"/>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtDynGroupSearch" Grid.Column="0"
                             Style="{StaticResource InputBox}"
                             Tag="Type group name, paste Object ID, or enter display name"/>
                    <Button x:Name="BtnDynGroupSearch" Grid.Column="2" Content="Search"
                            Style="{StaticResource OutlineBtn}" Width="70" Margin="0,0,8,0"/>
                    <Button x:Name="BtnDynGroupById" Grid.Column="4" Content="Use as ID"
                            Style="{StaticResource OutlineBtn}" Width="80"/>
                  </Grid>
                  <Border Margin="0,4,0,0" BorderBrush="$($DevConfig.ColorBorder)"
                          BorderThickness="1" CornerRadius="5"
                          x:Name="DynGroupSearchResults" Visibility="Collapsed" MaxHeight="150">
                    <ListBox x:Name="LstDynGroupResults" Background="$($DevConfig.ColorSurface)"
                             BorderThickness="0" FontSize="12"/>
                  </Border>
                  <Border x:Name="DynSelectedGroupBadge" Margin="0,6,0,0" Visibility="Collapsed"
                          Background="#EEF5FF" BorderBrush="$($DevConfig.ColorPrimary)"
                          BorderThickness="1" CornerRadius="5" Padding="10,6">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Text="Selected:" FontSize="11" Foreground="$($DevConfig.ColorTextMuted)"/>
                      <TextBlock x:Name="TxtDynGroupName" FontSize="12" FontWeight="SemiBold"
                                 Foreground="$($DevConfig.ColorPrimary)"/>
                      <TextBlock x:Name="TxtDynGroupId" FontSize="10"
                                 Foreground="$($DevConfig.ColorTextMuted)"/>
                      <TextBlock x:Name="TxtDynCurrentRule" FontSize="10" FontStyle="Italic"
                                 Foreground="$($DevConfig.ColorTextMuted)"/>
                    </StackPanel>
                  </Border>
                </StackPanel>

                <!-- Rule input -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="MEMBERSHIP RULE" Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="TxtDynRule" Style="{StaticResource InputBox}"
                           Height="80" TextWrapping="Wrap" AcceptsReturn="True"
                           VerticalScrollBarVisibility="Auto"
                           Tag='e.g. (user.department -eq "Marketing") -and (user.accountEnabled -eq true)'/>
                </StackPanel>

                <TextBlock Margin="0,0,0,12" Text="&#x26A0; Setting a dynamic rule will convert a static group to dynamic. This removes all current members."
                           FontSize="11" Foreground="$($DevConfig.ColorWarning)"
                           TextWrapping="Wrap"/>

                <Button Margin="4,4,0,0" x:Name="BtnDynValidate" Content="Validate Rule &amp; Preview &#x2192;"
                        Style="{StaticResource PrimaryBtn}" HorizontalAlignment="Left" Width="200"/>
              </StackPanel>
            </Border>

            <!-- === RENAME panel === -->
            <Border x:Name="PanelRename" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Rename Group" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,12"
                           Foreground="$($DevConfig.ColorPrimary)"/>

                <!-- Group picker -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="TARGET GROUP" Style="{StaticResource FieldLabel}"/>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtRenGroupSearch" Grid.Column="0"
                             Style="{StaticResource InputBox}"
                             Tag="Type group name, paste Object ID, or enter display name"/>
                    <Button x:Name="BtnRenGroupSearch" Grid.Column="2" Content="Search"
                            Style="{StaticResource OutlineBtn}" Width="70" Margin="0,0,8,0"/>
                    <Button x:Name="BtnRenGroupById" Grid.Column="4" Content="Use as ID"
                            Style="{StaticResource OutlineBtn}" Width="80"/>
                  </Grid>
                  <Border Margin="0,4,0,0" BorderBrush="$($DevConfig.ColorBorder)"
                          BorderThickness="1" CornerRadius="5"
                          x:Name="RenGroupSearchResults" Visibility="Collapsed" MaxHeight="150">
                    <ListBox x:Name="LstRenGroupResults" Background="$($DevConfig.ColorSurface)"
                             BorderThickness="0" FontSize="12"/>
                  </Border>
                  <Border x:Name="RenSelectedGroupBadge" Margin="0,6,0,0" Visibility="Collapsed"
                          Background="#EEF5FF" BorderBrush="$($DevConfig.ColorPrimary)"
                          BorderThickness="1" CornerRadius="5" Padding="10,6">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Text="Current name:" FontSize="11" Foreground="$($DevConfig.ColorTextMuted)"/>
                      <TextBlock x:Name="TxtRenGroupCurrentName" FontSize="12" FontWeight="SemiBold"
                                 Foreground="$($DevConfig.ColorPrimary)"/>
                      <TextBlock x:Name="TxtRenGroupId" FontSize="10"
                                 Foreground="$($DevConfig.ColorTextMuted)"/>
                    </StackPanel>
                  </Border>
                </StackPanel>

                <!-- New name -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="NEW DISPLAY NAME *" Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="TxtNewGroupName" Style="{StaticResource InputBox}"
                           Tag="Enter new display name"/>
                </StackPanel>

                <Button Margin="4,4,0,0" x:Name="BtnRenValidate" Content="Validate &amp; Preview &#x2192;"
                        Style="{StaticResource PrimaryBtn}" HorizontalAlignment="Left" Width="180"/>
              </StackPanel>
            </Border>

            <!-- === SET OWNER panel === -->
            <Border x:Name="PanelOwner" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Set Group Owner" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,12"
                           Foreground="$($DevConfig.ColorPrimary)"/>

                <!-- Group list input -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="TARGET GROUPS  -  one per line (display name or Object ID)"
                             Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="TxtOwnerGroupList" Style="{StaticResource InputBox}"
                           Height="90" TextWrapping="Wrap" AcceptsReturn="True"
                           VerticalScrollBarVisibility="Auto"
                           Tag="Group-Finance-APAC&#10;Group-Finance-EMEA&#10;b3f1a2c0-..."/>
                </StackPanel>

                <!-- Owner UPNs -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="OWNER UPNs  -  one per line"
                             Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="TxtOwnerList" Style="{StaticResource InputBox}"
                           Height="100" TextWrapping="Wrap" AcceptsReturn="True"
                           VerticalScrollBarVisibility="Auto"
                           Tag="owner1@domain.com&#10;owner2@domain.com"/>
                </StackPanel>

                <Button Margin="4,4,0,0" x:Name="BtnOwnerValidate" Content="Validate &amp; Preview &#x2192;"
                        Style="{StaticResource PrimaryBtn}" HorizontalAlignment="Left" Width="180"/>
              </StackPanel>
            </Border>



            <!-- === ADD USER DEVICES TO GROUP panel === -->
            <Border x:Name="PanelUserDevices" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Add User Devices to Group" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,12"
                           Foreground="$($DevConfig.ColorPrimary)"/>

                <!-- Target group picker -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="TARGET GROUP" Style="{StaticResource FieldLabel}"/>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtUDGroupSearch" Grid.Column="0"
                             Style="{StaticResource InputBox}"
                             Tag="Type group name, paste Object ID, or enter display name"/>
                    <Button x:Name="BtnUDGroupSearch" Grid.Column="2" Content="Search"
                            Style="{StaticResource OutlineBtn}" Width="70" Margin="0,0,8,0"/>
                    <Button x:Name="BtnUDGroupById" Grid.Column="4" Content="Use as ID"
                            Style="{StaticResource OutlineBtn}" Width="80"/>
                  </Grid>
                  <Border Margin="0,4,0,0" BorderBrush="$($DevConfig.ColorBorder)"
                          BorderThickness="1" CornerRadius="5"
                          x:Name="UDGroupSearchResults" Visibility="Collapsed" MaxHeight="150">
                    <ListBox x:Name="LstUDGroupResults" Background="$($DevConfig.ColorSurface)"
                             BorderThickness="0" FontSize="12"/>
                  </Border>
                  <Border x:Name="UDSelectedGroupBadge" Margin="0,6,0,0" Visibility="Collapsed"
                          Background="#EEF5FF" BorderBrush="$($DevConfig.ColorPrimary)"
                          BorderThickness="1" CornerRadius="5" Padding="10,6">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Text="Selected:" FontSize="11" Foreground="$($DevConfig.ColorTextMuted)"/>
                      <TextBlock x:Name="TxtUDGroupName" FontSize="12" FontWeight="SemiBold"
                                 Foreground="$($DevConfig.ColorPrimary)"/>
                      <TextBlock x:Name="TxtUDGroupId" FontSize="10"
                                 Foreground="$($DevConfig.ColorTextMuted)"/>
                    </StackPanel>
                  </Border>
                </StackPanel>

                <!-- User UPNs -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="USER UPNs  -  one per line (devices owned by these users will be evaluated)"
                             Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="TxtUDUpnList" Style="{StaticResource InputBox}"
                           Height="100" TextWrapping="Wrap" AcceptsReturn="True"
                           VerticalScrollBarVisibility="Auto"
                           Tag="user1@domain.com&#10;user2@domain.com"/>
                </StackPanel>

                <!-- Filters -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="DEVICE PLATFORM FILTER  -  only devices matching a checked platform will be added"
                             Style="{StaticResource FieldLabel}"/>
                  <WrapPanel Margin="0,4,0,0">
                    <CheckBox x:Name="ChkUDWindows" Content="Windows"  IsChecked="True"  Margin="0,0,16,4" FontSize="12"/>
                    <CheckBox x:Name="ChkUDAndroid" Content="Android"  IsChecked="True"  Margin="0,0,16,4" FontSize="12"/>
                    <CheckBox x:Name="ChkUDiOS"     Content="iOS"      IsChecked="True"  Margin="0,0,16,4" FontSize="12"/>
                    <CheckBox x:Name="ChkUDMacOS"   Content="macOS"    IsChecked="True"  Margin="0,0,0,4"  FontSize="12"/>
                  </WrapPanel>
                </StackPanel>

                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="DEVICE OWNERSHIP FILTER"
                             Style="{StaticResource FieldLabel}"/>
                  <ComboBox x:Name="CmbUDOwnership" Style="{StaticResource StyledCombo}"
                            Width="200" HorizontalAlignment="Left" Margin="0,4,0,0">
                    <ComboBoxItem Content="All ownership types" IsSelected="True"/>
                    <ComboBoxItem Content="Company only"/>
                    <ComboBoxItem Content="Personal only"/>
                  </ComboBox>
                </StackPanel>

                <StackPanel Margin="0,0,0,16">
                  <TextBlock Text="INTUNE ENROLLMENT FILTER"
                             Style="{StaticResource FieldLabel}"/>
                  <CheckBox x:Name="ChkUDIntuneOnly" Content="Intune enrolled devices only"
                            IsChecked="False" FontSize="12" Margin="0,4,0,0"
                            ToolTip="When checked, only devices managed by Intune (managementType = MDM) are included"/>
                </StackPanel>

                <Button Margin="4,4,0,0" x:Name="BtnUDValidate" Content="Validate &amp; Preview &#x2192;"
                        Style="{StaticResource PrimaryBtn}" HorizontalAlignment="Left" Width="180"/>
              </StackPanel>
            </Border>

            <!-- === FIND COMMON GROUPS panel === -->
            <Border x:Name="PanelFindCommon" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Find Common Groups" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,4"
                           Foreground="$($DevConfig.ColorPrimary)"/>
                <TextBlock Text="Finds groups that every input object is a member of."
                           FontSize="11" Foreground="$($DevConfig.ColorTextMuted)" Margin="0,0,0,14"
                           TextWrapping="Wrap"/>

                <!-- Input type selector -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="INPUT TYPE" Style="{StaticResource FieldLabel}"/>
                  <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
                    <RadioButton x:Name="RbFCUsers"   Content="Users"   GroupName="FCType"
                                 IsChecked="True" FontSize="12" Margin="0,0,20,0" VerticalContentAlignment="Center"/>
                    <RadioButton x:Name="RbFCDevices" Content="Devices" GroupName="FCType"
                                 FontSize="12" Margin="0,0,20,0" VerticalContentAlignment="Center"/>
                    <RadioButton x:Name="RbFCGroups"  Content="Groups"  GroupName="FCType"
                                 FontSize="12" VerticalContentAlignment="Center"/>
                  </StackPanel>
                </StackPanel>

                <!-- Input list + inline Find button -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock x:Name="TxtFCInputLabel"
                             Text="USER UPNs  -  one per line"
                             Style="{StaticResource FieldLabel}"/>
                  <Grid Margin="0,4,0,0">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtFCInputList" Grid.Column="0" Style="{StaticResource InputBox}"
                             Height="130" TextWrapping="Wrap" AcceptsReturn="True"
                             VerticalScrollBarVisibility="Auto"
                             Tag="Enter one identifier per line"/>
                    <Button Grid.Column="2" x:Name="BtnFCFind" Content="&#x1F50D;  Find Common Groups"
                            Style="{StaticResource PrimaryBtn}" Width="160" VerticalAlignment="Top"/>
                  </Grid>
                </StackPanel>

                <!-- Progress overlay -->
                <Border x:Name="PnlFCProgress" Visibility="Collapsed"
                        Background="#1A2E1A" CornerRadius="6" Padding="16,12" Margin="0,0,0,12">
                  <StackPanel>
                    <TextBlock x:Name="TxtFCProgressMsg" Text="Resolving objects..."
                               Foreground="#A3E4C0" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <ProgressBar x:Name="PbFCProgress" Height="6" IsIndeterminate="True"
                                 Background="#2A4A2A" Foreground="#22C55E" BorderThickness="0"/>
                    <TextBlock x:Name="TxtFCProgressDetail" Text=""
                               Foreground="#6BB88B" FontSize="11" Margin="0,5,0,0"/>
                  </StackPanel>
                </Border>

                <!-- Results area -->
                <Border x:Name="PnlFCResults" Visibility="Collapsed">
                  <StackPanel>
                    <Grid Margin="0,0,0,8">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock x:Name="TxtFCCount" Grid.Column="0"
                                 FontSize="11" Foreground="#6B7280" VerticalAlignment="Center"/>
                      <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <Button x:Name="BtnFCFilterClear" Content="&#x2715;"
                                Style="{StaticResource PrimaryBtn}" ToolTip="Clear filter"
                                IsEnabled="False" Height="30" Width="30" FontSize="12"
                                Padding="0" Margin="0,0,4,0"/>
                        <TextBox x:Name="TxtFCFilter" Width="150" Height="30" FontSize="12"
                                VerticalContentAlignment="Center" Padding="6,0" Margin="0,0,6,0"
                                IsEnabled="False" Tag="keyword filter..."/>
                        <Button x:Name="BtnFCFilter" Content="&#x1F50D; Filter"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,12,0"/>
                        <Button x:Name="BtnFCCopyValue" Content="Copy Value"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnFCCopyRow" Content="Copy Row"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnFCCopyAll" Content="Copy All"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnFCExportXlsx" Content="Export XLSX"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4"/>
                      </StackPanel>
                    </Grid>
                    <DataGrid x:Name="DgFCResults"
                              AutoGenerateColumns="False" IsReadOnly="True"
                              SelectionMode="Extended" SelectionUnit="CellOrRowHeader"
                              HeadersVisibility="Column"
                              GridLinesVisibility="All"
                              VerticalGridLinesBrush="#555555"
                              HorizontalGridLinesBrush="#D0D0D0"
                              RowBackground="White" AlternatingRowBackground="#F5FBF5"
                              BorderBrush="#B0D0B0" BorderThickness="1"
                              CanUserReorderColumns="False" CanUserResizeRows="False" CanUserResizeColumns="True"
                              CanUserSortColumns="True" MaxHeight="420" FontSize="12"
                              ScrollViewer.HorizontalScrollBarVisibility="Auto">
                      <DataGrid.ColumnHeaderStyle>
                        <Style TargetType="DataGridColumnHeader">
                          <Setter Property="FontWeight" Value="SemiBold"/>
                          <Setter Property="Background" Value="#E8F5E9"/>
                          <Setter Property="Foreground" Value="#14532D"/>
                          <Setter Property="Padding"    Value="8,5"/>
                          <Setter Property="BorderBrush" Value="#555555"/>
                          <Setter Property="BorderThickness" Value="0,0,1,1"/>
                        </Style>
                      </DataGrid.ColumnHeaderStyle>
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Group Name"      Binding="{Binding GroupName}"      Width="2*" MinWidth="160"/>
                        <DataGridTextColumn Header="Group ID"        Binding="{Binding GroupID}"        Width="2*" MinWidth="120"/>
                        <DataGridTextColumn Header="Group Type"      Binding="{Binding GroupType}"      Width="*"  MinWidth="100"/>
                        <DataGridTextColumn Header="Membership Type" Binding="{Binding MembershipType}" Width="*"  MinWidth="120"/>
                      </DataGrid.Columns>
                    </DataGrid>
                  </StackPanel>
                </Border>
                <TextBlock x:Name="TxtFCNoResults" Visibility="Collapsed"
                           Text="No common groups found for the specified objects."
                           FontSize="12" Foreground="#6B7280" Margin="0,10,0,0"/>
              </StackPanel>
            </Border>

            <!-- === FIND DISTINCT GROUPS panel === -->
            <Border x:Name="PanelFindDistinct" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Find Distinct Groups" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,4"
                           Foreground="$($DevConfig.ColorPrimary)"/>
                <TextBlock Text="Finds groups that are NOT shared by all input objects (union minus intersection)."
                           FontSize="11" Foreground="$($DevConfig.ColorTextMuted)" Margin="0,0,0,14"
                           TextWrapping="Wrap"/>

                <!-- Input type selector -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="INPUT TYPE" Style="{StaticResource FieldLabel}"/>
                  <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
                    <RadioButton x:Name="RbFDUsers"   Content="Users"   GroupName="FDType"
                                 IsChecked="True" FontSize="12" Margin="0,0,20,0" VerticalContentAlignment="Center"/>
                    <RadioButton x:Name="RbFDDevices" Content="Devices" GroupName="FDType"
                                 FontSize="12" Margin="0,0,20,0" VerticalContentAlignment="Center"/>
                    <RadioButton x:Name="RbFDGroups"  Content="Groups"  GroupName="FDType"
                                 FontSize="12" VerticalContentAlignment="Center"/>
                  </StackPanel>
                </StackPanel>

                <!-- Input list + inline Find button -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock x:Name="TxtFDInputLabel"
                             Text="USER UPNs  -  one per line"
                             Style="{StaticResource FieldLabel}"/>
                  <Grid Margin="0,4,0,0">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtFDInputList" Grid.Column="0" Style="{StaticResource InputBox}"
                             Height="130" TextWrapping="Wrap" AcceptsReturn="True"
                             VerticalScrollBarVisibility="Auto"
                             Tag="Enter one identifier per line"/>
                    <Button Grid.Column="2" x:Name="BtnFDFind" Content="&#x1F50D;  Find Distinct Groups"
                            Style="{StaticResource PrimaryBtn}" Width="160" VerticalAlignment="Top"/>
                  </Grid>
                </StackPanel>

                <!-- Progress overlay -->
                <Border x:Name="PnlFDProgress" Visibility="Collapsed"
                        Background="#1A2E1A" CornerRadius="6" Padding="16,12" Margin="0,0,0,12">
                  <StackPanel>
                    <TextBlock x:Name="TxtFDProgressMsg" Text="Resolving objects..."
                               Foreground="#A3E4C0" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <ProgressBar x:Name="PbFDProgress" Height="6" IsIndeterminate="True"
                                 Background="#2A4A2A" Foreground="#22C55E" BorderThickness="0"/>
                    <TextBlock x:Name="TxtFDProgressDetail" Text=""
                               Foreground="#6BB88B" FontSize="11" Margin="0,5,0,0"/>
                  </StackPanel>
                </Border>

                <!-- Results area -->
                <Border x:Name="PnlFDResults" Visibility="Collapsed">
                  <StackPanel>
                    <Grid Margin="0,0,0,8">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock x:Name="TxtFDCount" Grid.Column="0"
                                 FontSize="11" Foreground="#6B7280" VerticalAlignment="Center"/>
                      <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <Button x:Name="BtnFDFilterClear" Content="&#x2715;"
                                Style="{StaticResource PrimaryBtn}" ToolTip="Clear filter"
                                IsEnabled="False" Height="30" Width="30" FontSize="12"
                                Padding="0" Margin="0,0,4,0"/>
                        <TextBox x:Name="TxtFDFilter" Width="150" Height="30" FontSize="12"
                                VerticalContentAlignment="Center" Padding="6,0" Margin="0,0,6,0"
                                IsEnabled="False" Tag="keyword filter..."/>
                        <Button x:Name="BtnFDFilter" Content="&#x1F50D; Filter"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,12,0"/>
                        <Button x:Name="BtnFDCopyValue" Content="Copy Value"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnFDCopyRow" Content="Copy Row"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnFDCopyAll" Content="Copy All"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnFDExportXlsx" Content="Export XLSX"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4"/>
                      </StackPanel>
                    </Grid>
                    <DataGrid x:Name="DgFDResults"
                              AutoGenerateColumns="False" IsReadOnly="True"
                              SelectionMode="Extended" SelectionUnit="CellOrRowHeader"
                              HeadersVisibility="Column"
                              GridLinesVisibility="All"
                              VerticalGridLinesBrush="#555555"
                              HorizontalGridLinesBrush="#D0D0D0"
                              RowBackground="White" AlternatingRowBackground="#EDF5FF"
                              BorderBrush="#A0C0E0" BorderThickness="1"
                              CanUserReorderColumns="False" CanUserResizeRows="False" CanUserResizeColumns="True"
                              CanUserSortColumns="True" MaxHeight="420" FontSize="12"
                              ScrollViewer.HorizontalScrollBarVisibility="Auto">
                      <DataGrid.ColumnHeaderStyle>
                        <Style TargetType="DataGridColumnHeader">
                          <Setter Property="FontWeight" Value="SemiBold"/>
                          <Setter Property="Background" Value="#DBEAFE"/>
                          <Setter Property="Foreground" Value="#1E3A5F"/>
                          <Setter Property="Padding"    Value="8,5"/>
                          <Setter Property="BorderBrush" Value="#555555"/>
                          <Setter Property="BorderThickness" Value="0,0,1,1"/>
                        </Style>
                      </DataGrid.ColumnHeaderStyle>
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Group Name"      Binding="{Binding GroupName}"      Width="2*" MinWidth="160"/>
                        <DataGridTextColumn Header="Group ID"        Binding="{Binding GroupID}"        Width="2*" MinWidth="120"/>
                        <DataGridTextColumn Header="Group Type"      Binding="{Binding GroupType}"      Width="*"  MinWidth="100"/>
                        <DataGridTextColumn Header="Membership Type" Binding="{Binding MembershipType}" Width="*"  MinWidth="120"/>
                        <DataGridTextColumn Header="Member"        Binding="{Binding MemberOf}"       Width="2*" MinWidth="160"/>
                      </DataGrid.Columns>
                    </DataGrid>
                  </StackPanel>
                </Border>
                <TextBlock x:Name="TxtFDNoResults" Visibility="Collapsed"
                           Text="No distinct groups found for the specified objects."
                           FontSize="12" Foreground="#6B7280" Margin="0,10,0,0"/>
              </StackPanel>
            </Border>

            <!-- === GET DEVICE INFO panel === -->
            <Border x:Name="PanelGetDeviceInfo" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Get Device Info" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,4"
                           Foreground="$($DevConfig.ColorPrimary)"/>
                <TextBlock Text="Fetches combined Entra and Intune device data and exports to XLSX."
                           FontSize="11" Foreground="$($DevConfig.ColorTextMuted)" Margin="0,0,0,14" TextWrapping="Wrap"/>

                <!-- Input type -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="INPUT TYPE" Style="{StaticResource FieldLabel}"/>
                  <ComboBox x:Name="CmbGDIInputType" Style="{StaticResource StyledCombo}"
                            Width="260" HorizontalAlignment="Left" Margin="0,4,0,0">
                    <ComboBoxItem Content="User UPNs (resolve owned devices)" IsSelected="True"/>
                    <ComboBoxItem Content="Device names"/>
                    <ComboBoxItem Content="Device serial numbers"/>
                    <ComboBoxItem Content="Entra Device Object IDs"/>
                    <ComboBoxItem Content="Intune Device IDs"/>
                    <ComboBoxItem Content="Groups (Users and Devices)"/>
                  </ComboBox>
                </StackPanel>

                <!-- Input list -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock x:Name="TxtGDIInputLabel" Text="USER UPNs  -  one per line"
                             Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="TxtGDIInputList" Style="{StaticResource InputBox}"
                           Height="110" TextWrapping="Wrap" AcceptsReturn="True"
                           VerticalScrollBarVisibility="Auto" Tag="user@domain.com"/>
                </StackPanel>

                <!-- Platform filter -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="PLATFORM FILTER (leave all unchecked for all platforms)"
                             Style="{StaticResource FieldLabel}"/>
                  <WrapPanel Margin="0,4,0,0">
                    <CheckBox x:Name="ChkGDIWindows" Content="Windows" Margin="0,0,16,4" FontSize="12"/>
                    <CheckBox x:Name="ChkGDIAndroid" Content="Android" Margin="0,0,16,4" FontSize="12"/>
                    <CheckBox x:Name="ChkGDIiOS"     Content="iOS"     Margin="0,0,16,4" FontSize="12"/>
                    <CheckBox x:Name="ChkGDIMacOS"   Content="macOS"   Margin="0,0,0,4"  FontSize="12"/>
                  </WrapPanel>
                </StackPanel>

                <!-- Ownership filter -->
                <StackPanel Margin="0,0,0,8">
                  <TextBlock Text="OWNERSHIP FILTER" Style="{StaticResource FieldLabel}"/>
                  <ComboBox x:Name="CmbGDIOwnership" Style="{StaticResource StyledCombo}"
                            Width="200" HorizontalAlignment="Left" Margin="0,4,0,0">
                    <ComboBoxItem Content="All ownership types" IsSelected="True"/>
                    <ComboBoxItem Content="Company only"/>
                    <ComboBoxItem Content="Personal only"/>
                  </ComboBox>
                </StackPanel>

                <!-- Run buttons -->
                <StackPanel Orientation="Horizontal" Margin="0,0,0,16">
                  <Button x:Name="BtnGDIRunAll" Content="&#x25B6;  Get All Device Info"
                          Style="{StaticResource PrimaryBtn}" Width="200" HorizontalAlignment="Left" Margin="0,4,0,0"/>
                  <Button x:Name="BtnGDIRun" Content="&#x25B6;  Get Device Info"
                          Style="{StaticResource PrimaryBtn}" Width="180" HorizontalAlignment="Left" Margin="12,4,0,0"/>
                </StackPanel>

                <!-- Progress -->
                <Border x:Name="PnlGDIProgress" Visibility="Collapsed"
                        Background="#1A2E1A" CornerRadius="6" Padding="16,12" Margin="0,0,0,12">
                  <StackPanel>
                    <TextBlock x:Name="TxtGDIProgressMsg" Text="Resolving devices..."
                               Foreground="#A3E4C0" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <ProgressBar x:Name="PbGDIProgress" Height="6" IsIndeterminate="True"
                                 Background="#2A4A2A" Foreground="#22C55E" BorderThickness="0"/>
                    <TextBlock x:Name="TxtGDIProgressDetail" Text=""
                               Foreground="#6BB88B" FontSize="11" Margin="0,5,0,0"/>
                  </StackPanel>
                </Border>

                <!-- Results -->
                <Border x:Name="PnlGDIResults" Visibility="Collapsed">
                  <StackPanel>
                    <Grid Margin="0,0,0,8">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock x:Name="TxtGDICount" Grid.Column="0"
                                 FontSize="11" Foreground="#6B7280" VerticalAlignment="Center"/>
                      <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <Button x:Name="BtnGDIFilterClear" Content="&#x2715;"
                                Style="{StaticResource PrimaryBtn}" ToolTip="Clear filter"
                                IsEnabled="False" Height="30" Width="30" FontSize="12"
                                Padding="0" Margin="0,0,4,0"/>
                        <TextBox x:Name="TxtGDIFilter" Width="150" Height="30" FontSize="12"
                                VerticalContentAlignment="Center" Padding="6,0" Margin="0,0,6,0"
                                IsEnabled="False" Tag="keyword filter..."/>
                        <Button x:Name="BtnGDIFilter" Content="&#x1F50D; Filter"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,12,0"/>
                        <Button x:Name="BtnGDICopyValue" Content="Copy Value"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnGDICopyRow" Content="Copy Row"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnGDICopyAll" Content="Copy All"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnGDIExportXlsx" Content="Export XLSX"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4"/>
                      </StackPanel>
                    </Grid>
                    <DataGrid x:Name="DgGDIResults"
                              AutoGenerateColumns="False" IsReadOnly="True"
                              SelectionMode="Extended" SelectionUnit="CellOrRowHeader"
                              HeadersVisibility="Column"
                              GridLinesVisibility="All"
                              VerticalGridLinesBrush="#555555"
                              HorizontalGridLinesBrush="#D0D0D0"
                              RowBackground="White" AlternatingRowBackground="#F5F5FF"
                              BorderBrush="#B0B0D0" BorderThickness="1"
                              CanUserReorderColumns="False" CanUserResizeRows="False" CanUserResizeColumns="True"
                              CanUserSortColumns="True" MaxHeight="420" FontSize="12"
                              ScrollViewer.HorizontalScrollBarVisibility="Auto">
                      <DataGrid.ColumnHeaderStyle>
                        <Style TargetType="DataGridColumnHeader">
                          <Setter Property="FontWeight" Value="SemiBold"/>
                          <Setter Property="Background" Value="#E8EAF6"/>
                          <Setter Property="Foreground" Value="#1A237E"/>
                          <Setter Property="Padding"    Value="8,5"/>
                          <Setter Property="BorderBrush" Value="#555555"/>
                          <Setter Property="BorderThickness" Value="0,0,1,1"/>
                        </Style>
                      </DataGrid.ColumnHeaderStyle>
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Device Name"         Binding="{Binding Entra_DeviceName}"          Width="2*" MinWidth="140"/>
                        <DataGridTextColumn Header="Device Type"         Binding="{Binding DeviceType}"                Width="*"  MinWidth="100"/>
                        <DataGridTextColumn Header="Join Type"           Binding="{Binding Entra_JoinType}"            Width="*"  MinWidth="110"/>
                        <DataGridTextColumn Header="Entra Ownership"     Binding="{Binding Entra_Ownership}"           Width="*"  MinWidth="100"/>
                        <DataGridTextColumn Header="Intune Ownership"    Binding="{Binding Intune_Ownership}"          Width="*"  MinWidth="100"/>
                        <DataGridTextColumn Header="Entra OS"            Binding="{Binding Entra_OSPlatform}"          Width="*"  MinWidth="80"/>
                        <DataGridTextColumn Header="Intune OS"           Binding="{Binding Intune_OSPlatform}"         Width="*"  MinWidth="80"/>
                        <DataGridTextColumn Header="Entra OS Version"    Binding="{Binding Entra_OSVersion}"           Width="*"  MinWidth="100"/>
                        <DataGridTextColumn Header="Intune OS Version"   Binding="{Binding Intune_OSVersion}"          Width="*"  MinWidth="100"/>
                        <DataGridTextColumn Header="Windows Release"     Binding="{Binding Windows_Release}"           Width="*"  MinWidth="100"/>
                        <DataGridTextColumn Header="Device Enabled"      Binding="{Binding Entra_DeviceEnabled}"       Width="*"  MinWidth="90"/>
                        <DataGridTextColumn Header="Entra Last Sign-In"  Binding="{Binding Entra_LastSignIn}"          Width="*"  MinWidth="110"/>
                        <DataGridTextColumn Header="Entra Activity"      Binding="{Binding Entra_ActivityRange}"       Width="*"  MinWidth="100"/>
                        <DataGridTextColumn Header="Intune Check-In"     Binding="{Binding Intune_LastCheckIn}"        Width="*"  MinWidth="110"/>
                        <DataGridTextColumn Header="Intune Activity"     Binding="{Binding Intune_ActivityRange}"      Width="*"  MinWidth="100"/>
                        <DataGridTextColumn Header="OEM"                 Binding="{Binding Intune_OEM}"                Width="*"  MinWidth="80"/>
                        <DataGridTextColumn Header="Model"               Binding="{Binding Intune_Model}"              Width="*"  MinWidth="100"/>
                        <DataGridTextColumn Header="Serial Number"       Binding="{Binding Intune_SerialNumber}"       Width="*"  MinWidth="120"/>
                        <DataGridTextColumn Header="Enrolled Date"       Binding="{Binding Intune_EnrolledDate}"       Width="*"  MinWidth="100"/>
                        <DataGridTextColumn Header="Compliance State"   Binding="{Binding Intune_ComplianceState}"    Width="*"  MinWidth="110"/>
                        <DataGridTextColumn Header="Entra Device ID"     Binding="{Binding Entra_DeviceID}"            Width="2*" MinWidth="130"/>
                        <DataGridTextColumn Header="Entra Object ID"     Binding="{Binding Entra_ObjectID}"            Width="2*" MinWidth="130"/>
                        <DataGridTextColumn Header="Intune Device ID"    Binding="{Binding Intune_DeviceID}"           Width="2*" MinWidth="130"/>
                        <DataGridTextColumn Header="Device Owner UPN"    Binding="{Binding Entra_DeviceOwner}"         Width="2*" MinWidth="160"/>
                        <DataGridTextColumn Header="Primary User UPN"    Binding="{Binding Entra_PrimaryUserUPN}"      Width="2*" MinWidth="160"/>
                        <DataGridTextColumn Header="User Enabled"        Binding="{Binding Entra_UserAccountEnabled}"  Width="*"  MinWidth="90"/>
                        <DataGridTextColumn Header="User Country"        Binding="{Binding Entra_UserCountry}"         Width="*"  MinWidth="90"/>
                        <DataGridTextColumn Header="User City"           Binding="{Binding Entra_UserCity}"            Width="*"  MinWidth="80"/>
                        <DataGridTextColumn Header="User Job Title"      Binding="{Binding Entra_UserJobTitle}"        Width="*"  MinWidth="110"/>
                        <DataGridTextColumn Header="AP Group Tag"        Binding="{Binding Autopilot_GroupTag}"        Width="*"  MinWidth="100"/>
                        <DataGridTextColumn Header="AP Profile Status"   Binding="{Binding Autopilot_ProfileStatus}"   Width="*"  MinWidth="120"/>
                      </DataGrid.Columns>
                    </DataGrid>
                  </StackPanel>
                </Border>
                <TextBlock x:Name="TxtGDINoResults" Visibility="Collapsed"
                           Text="No devices matched the input and filters."
                           FontSize="12" Foreground="#6B7280" Margin="0,12,0,0" TextWrapping="Wrap"/>
              </StackPanel>
            </Border>




            <!-- === GET APP ASSIGNMENTS panel === -->
            <Border x:Name="PanelGetDiscoveredApps" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Get App Info" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,4"
                           Foreground="$($DevConfig.ColorPrimary)"/>
                <TextBlock Text="Reports discovered and managed apps from Intune. Supports device-level discovery and tenant-wide managed app queries with assignment details."
                           FontSize="11" Foreground="$($DevConfig.ColorTextMuted)" Margin="0,0,0,14" TextWrapping="Wrap"/>


                <StackPanel x:Name="PnlDAInputArea">
                  <!-- Input type -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="INPUT TYPE" Style="{StaticResource FieldLabel}"/>
                  <ComboBox x:Name="CmbDAInputType" Style="{StaticResource StyledCombo}"
                            Width="300" HorizontalAlignment="Left" Margin="0,4,0,0">
                    <ComboBoxItem Content="User UPNs (resolve owned devices)" IsSelected="True"/>
                    <ComboBoxItem Content="Device names"/>
                    <ComboBoxItem Content="Device IDs (Entra)"/>
                    <ComboBoxItem Content="Serial numbers"/>
                    <ComboBoxItem Content="Groups (names or IDs)"/>
                  </ComboBox>
                </StackPanel>

                <!-- Input list -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock x:Name="TxtDAInputLabel" Text="USER UPNs  -  one per line"
                             Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="TxtDAInputList" Style="{StaticResource InputBox}"
                           Height="110" TextWrapping="Wrap" AcceptsReturn="True"
                           VerticalScrollBarVisibility="Auto" Tag="user@domain.com"/>
                </StackPanel>
                </StackPanel>

                <!-- Platform filter + Keyword filter (side by side) -->
                <Grid Margin="0,0,0,2">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="40"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0">
                    <TextBlock Text="PLATFORM FILTER (leave all unchecked for all platforms)"
                               Style="{StaticResource FieldLabel}"/>
                    <WrapPanel Margin="0,4,0,0">
                      <CheckBox x:Name="ChkDAWindows" Content="Windows" Margin="0,0,16,4" FontSize="12"/>
                      <CheckBox x:Name="ChkDAAndroid" Content="Android" Margin="0,0,16,4" FontSize="12"/>
                      <CheckBox x:Name="ChkDAiOS"     Content="iOS"     Margin="0,0,16,4" FontSize="12"/>
                      <CheckBox x:Name="ChkDAMacOS"   Content="macOS"   Margin="0,0,0,4"  FontSize="12"/>
                    </WrapPanel>
                  </StackPanel>
                  <StackPanel Grid.Column="2" VerticalAlignment="Top">
                    <TextBlock Text="APP KEYWORD FILTER (optional)" Style="{StaticResource FieldLabel}"/>
                    <TextBox x:Name="TxtDAKeyword" Style="{StaticResource InputBox}"
                             Width="250" HorizontalAlignment="Left" Margin="0,4,0,0"
                             Tag="e.g. Chrome, Zoom, Office"/>
                  </StackPanel>
                </Grid>

                <!-- Ownership filter -->
                <StackPanel x:Name="PnlDAOwnership" Margin="0,0,0,4">
                  <TextBlock Text="OWNERSHIP FILTER" Style="{StaticResource FieldLabel}"/>
                  <ComboBox x:Name="CmbDAOwnership" Style="{StaticResource StyledCombo}"
                            Width="200" HorizontalAlignment="Left" Margin="0,4,0,0">
                    <ComboBoxItem Content="All ownership types" IsSelected="True"/>
                    <ComboBoxItem Content="Company only"/>
                    <ComboBoxItem Content="Personal only"/>
                  </ComboBox>
                </StackPanel>

                <!-- App Source filter -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="APP SOURCE (select one or both)" Style="{StaticResource FieldLabel}"/>
                  <WrapPanel Margin="0,4,0,0">
                    <CheckBox x:Name="ChkDADiscoveredApps" Content="Discovered Apps" IsChecked="True" Margin="0,0,16,4" FontSize="12"/>
                    <CheckBox x:Name="ChkDAManagedApps"    Content="Managed Apps"    IsChecked="True" Margin="0,0,0,4"  FontSize="12"/>
                  </WrapPanel>
                </StackPanel>

                <!-- Action buttons -->
                <WrapPanel Margin="0,0,0,16">
                  <Button x:Name="BtnDARun" Content="&#x25B6;  Get Selected Apps"
                          Style="{StaticResource PrimaryBtn}"
                          Height="34" FontSize="13" FontWeight="Bold" Padding="18,6" Margin="0,0,10,0"/>
                  <Button x:Name="BtnDARunAll" Content="&#x25B6;  All Discovered Apps - All Devices"
                          Style="{StaticResource PrimaryBtn}"
                          Height="34" FontSize="13" FontWeight="Bold" Padding="18,6" Margin="0,0,10,0"/>
                  <Button x:Name="BtnDAGetManagedAssignments" Content="&#x25B6;  All Managed App Assignments"
                          Style="{StaticResource PrimaryBtn}"
                          Height="34" FontSize="13" FontWeight="Bold" Padding="18,6"/>
                </WrapPanel>


                <!-- Progress -->
                <Border x:Name="PnlDAProgress" Visibility="Collapsed"
                        Background="#1A2E1A" CornerRadius="6" Padding="16,12" Margin="0,0,0,12">
                  <StackPanel>
                    <TextBlock x:Name="TxtDAProgressMsg" Text="Resolving devices..."
                               Foreground="#A3E4C0" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <ProgressBar x:Name="PbDAProgress" Height="6" IsIndeterminate="True"
                                 Background="#2A4A2A" Foreground="#22C55E" BorderThickness="0"/>
                    <TextBlock x:Name="TxtDAProgressDetail" Text=""
                               Foreground="#6BB88B" FontSize="11" Margin="0,5,0,0"/>
                  </StackPanel>
                </Border>

                <!-- Results -->
                <Border x:Name="PnlDAResults" Visibility="Collapsed">
                  <StackPanel>
                    <Grid Margin="0,0,0,8">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock x:Name="TxtDACount" Grid.Column="0"
                                 FontSize="11" Foreground="#6B7280" VerticalAlignment="Center"/>
                      <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <Button x:Name="BtnDAFilterClear" Content="&#x2715;"
                                Style="{StaticResource PrimaryBtn}" ToolTip="Clear filter"
                                IsEnabled="False" Height="30" Width="30" FontSize="12"
                                Padding="0" Margin="0,0,4,0"/>
                        <TextBox x:Name="TxtDAFilter" Width="150" Height="30" FontSize="12"
                                VerticalContentAlignment="Center" Padding="6,0" Margin="0,0,6,0"
                                IsEnabled="False" Tag="keyword filter..."/>
                        <Button x:Name="BtnDAFilter" Content="&#x1F50D; Filter"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,12,0"/>
                        <Button x:Name="BtnDACopyValue" Content="Copy Value"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnDACopyRow" Content="Copy Row"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnDACopyAll" Content="Copy All"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnDAExportXlsx" Content="Export XLSX"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4"/>
                      </StackPanel>
                    </Grid>
                    <DataGrid x:Name="DgDAResults"
                              AutoGenerateColumns="False" IsReadOnly="True"
                              SelectionMode="Extended" SelectionUnit="CellOrRowHeader"
                              HeadersVisibility="Column"
                              GridLinesVisibility="All"
                              VerticalGridLinesBrush="#555555"
                              HorizontalGridLinesBrush="#D0D0D0"
                              RowBackground="White" AlternatingRowBackground="#F5F5FF"
                              BorderBrush="#B0B0D0" BorderThickness="1"
                              CanUserReorderColumns="False" CanUserResizeRows="False" CanUserResizeColumns="True"
                              CanUserSortColumns="True" MaxHeight="420" FontSize="12"
                              ScrollViewer.HorizontalScrollBarVisibility="Auto">
                      <DataGrid.ColumnHeaderStyle>
                        <Style TargetType="DataGridColumnHeader">
                          <Setter Property="FontWeight" Value="SemiBold"/>
                          <Setter Property="Background" Value="#E8EAF6"/>
                          <Setter Property="Foreground" Value="#1A237E"/>
                          <Setter Property="Padding"    Value="8,5"/>
                          <Setter Property="BorderBrush" Value="#555555"/>
                          <Setter Property="BorderThickness" Value="0,0,1,1"/>
                        </Style>
                      </DataGrid.ColumnHeaderStyle>
                    </DataGrid>
                  </StackPanel>
                </Border>
                <TextBlock x:Name="TxtDANoResults" Visibility="Collapsed"
                           Text="No apps matched the input and filters."
                           FontSize="12" Foreground="#6B7280" Margin="0,12,0,0" TextWrapping="Wrap"/>
              </StackPanel>
            </Border>


            <!-- === GET POLICY ASSIGNMENTS panel === -->
            <Border x:Name="PanelGetPolicyAssignments" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Get Policy Info" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,4"
                           Foreground="$($DevConfig.ColorPrimary)"/>
                <TextBlock Text="Query Intune policy assignments across all policy types. Use filters to narrow results, or Get All Assignments to retrieve everything."
                           FontSize="11" Foreground="$($DevConfig.ColorTextMuted)" Margin="0,0,0,14" TextWrapping="Wrap"/>


                <!-- Filters + action buttons -->
                <Grid Margin="0,0,0,14">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="16"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>

                  <!-- Filter stack (vertical) -->
                  <StackPanel Grid.Column="0">

                    <!-- Policy Name keyword -->
                    <StackPanel Margin="0,0,0,10">
                      <TextBlock Text="POLICY NAME  (keyword &#x2014; blank for all)" Style="{StaticResource FieldLabel}"/>
                      <Border Background="#FFFFFF" BorderBrush="#B0D0B0" BorderThickness="1"
                              CornerRadius="8" Height="36" Margin="0,4,0,0">
                        <TextBox x:Name="TxtGPANameFilter"
                                 Background="Transparent" Foreground="#374151" CaretBrush="#22C55E"
                                 BorderThickness="0" Padding="10,0" FontSize="12"
                                 VerticalContentAlignment="Center"/>
                      </Border>
                    </StackPanel>

                    <!-- Policy Type -->
                    <StackPanel Margin="0,0,0,10">
                      <TextBlock Text="POLICY TYPE" Style="{StaticResource FieldLabel}"/>
                      <ComboBox x:Name="CmbGPAType" Style="{StaticResource StyledCombo}" Margin="0,4,0,0">
                        <ComboBoxItem Content="(All types)" IsSelected="True"/>
                        <ComboBoxItem Content="Device Configuration"/>
                        <ComboBoxItem Content="Configuration Policy"/>
                        <ComboBoxItem Content="Autopilot Profile"/>
                        <ComboBoxItem Content="Platform Script (Windows PowerShell)"/>
                        <ComboBoxItem Content="Detection/Remediation Script"/>
                        <ComboBoxItem Content="Administrative Templates"/>
                        <ComboBoxItem Content="Compliance Policy"/>
                      </ComboBox>
                    </StackPanel>

                    <!-- Policy Sub-Type -->
                    <StackPanel Margin="0,0,0,10">
                      <TextBlock Text="POLICY SUB-TYPE" Style="{StaticResource FieldLabel}"/>
                      <ComboBox x:Name="CmbGPASubType" Style="{StaticResource StyledCombo}" Margin="0,4,0,0">
                        <ComboBoxItem Content="(All sub-types)" IsSelected="True"/>
                        <ComboBoxItem Content="Administrative Templates"/>
                        <ComboBoxItem Content="App Control for Business"/>
                        <ComboBoxItem Content="Autopilot"/>
                        <ComboBoxItem Content="Autopilot (Device Prep)"/>
                        <ComboBoxItem Content="Autopilot ESP"/>
                        <ComboBoxItem Content="Attack Surface Reduction Rules"/>
                        <ComboBoxItem Content="BitLocker"/>
                        <ComboBoxItem Content="Built-in"/>
                        <ComboBoxItem Content="Certificate"/>
                        <ComboBoxItem Content="Custom"/>
                        <ComboBoxItem Content="Defender Update Controls"/>
                        <ComboBoxItem Content="Device Features"/>
                        <ComboBoxItem Content="Edition Upgrade"/>
                        <ComboBoxItem Content="Endpoint Protection"/>
                        <ComboBoxItem Content="Exploit Protection"/>
                        <ComboBoxItem Content="Extensions"/>
                        <ComboBoxItem Content="General Device"/>
                        <ComboBoxItem Content="Health Monitoring"/>
                        <ComboBoxItem Content="Identity Protection"/>
                        <ComboBoxItem Content="Kiosk"/>
                        <ComboBoxItem Content="Local Admin Password Solution (Windows LAPS)"/>
                        <ComboBoxItem Content="Local User Group Membership"/>
                        <ComboBoxItem Content="MDE Onboarding Policy"/>
                        <ComboBoxItem Content="Microsoft Defender Antivirus"/>
                        <ComboBoxItem Content="Microsoft Defender Antivirus Exclusions"/>
                        <ComboBoxItem Content="PowerShell"/>
                        <ComboBoxItem Content="Remediation"/>
                        <ComboBoxItem Content="SCEP Certificate"/>
                        <ComboBoxItem Content="Settings Catalog"/>
                        <ComboBoxItem Content="Shared PC"/>
                        <ComboBoxItem Content="Software Update"/>
                        <ComboBoxItem Content="Trusted Certificate"/>
                        <ComboBoxItem Content="Update Rings"/>
                        <ComboBoxItem Content="VPN"/>
                        <ComboBoxItem Content="Wi-Fi"/>
                        <ComboBoxItem Content="Windows Firewall"/>
                        <ComboBoxItem Content="Windows Firewall Rules"/>
                        <ComboBoxItem Content="Windows OS Recovery"/>
                        <ComboBoxItem Content="Windows Security Experience"/>
                      </ComboBox>
                    </StackPanel>

                    <!-- OS Platform -->
                    <StackPanel>
                      <TextBlock Text="OS PLATFORM" Style="{StaticResource FieldLabel}"/>
                      <ComboBox x:Name="CmbGPAPlatform" Style="{StaticResource StyledCombo}" Margin="0,4,0,0">
                        <ComboBoxItem Content="(All platforms)" IsSelected="True"/>
                        <ComboBoxItem Content="Windows"/>
                        <ComboBoxItem Content="iOS/iPadOS"/>
                        <ComboBoxItem Content="Android"/>
                        <ComboBoxItem Content="AOSP"/>
                        <ComboBoxItem Content="macOS"/>
                        <ComboBoxItem Content="Linux"/>
                      </ComboBox>
                    </StackPanel>

                  </StackPanel>

                  <!-- Action buttons (right of filters) -->
                  <StackPanel Grid.Column="2" VerticalAlignment="Center">
                    <Button x:Name="BtnGPAGetSelected" Content="&#x25B6;  Get Selected Assignments"
                            Style="{StaticResource PrimaryBtn}" Height="32" FontSize="12" Padding="14,4"/>
                    <Button x:Name="BtnGPAGetAll" Content="&#x25B6;  Get All Assignments"
                            Style="{StaticResource PrimaryBtn}" Height="32" FontSize="12" Padding="14,4" Margin="0,8,0,0"/>
                  </StackPanel>

                </Grid>

                <!-- Progress -->
                <Border x:Name="PnlGPAProgress" Visibility="Collapsed"
                        Background="#1A2E1A" CornerRadius="6" Padding="16,12" Margin="0,0,0,12">
                  <StackPanel>
                    <TextBlock x:Name="TxtGPAProgressMsg" Text="Loading policies..."
                               Foreground="#A3E4C0" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <ProgressBar x:Name="PbGPAProgress" Height="6" IsIndeterminate="True"
                                 Background="#2A4A2A" Foreground="#22C55E" BorderThickness="0"/>
                    <TextBlock x:Name="TxtGPAProgressDetail" Text=""
                               Foreground="#6BB88B" FontSize="11" Margin="0,5,0,0"/>
                  </StackPanel>
                </Border>

                <!-- No results message -->
                <TextBlock x:Name="TxtGPANoResults" Visibility="Collapsed"
                           Text="No policy assignments matched the selected filters."
                           FontSize="12" Foreground="#6B7280" Margin="0,12,0,0" TextWrapping="Wrap"/>

                <!-- Results panel -->
                <Border x:Name="PnlGPAResults" Visibility="Collapsed">
                  <StackPanel>
                    <Grid Margin="0,0,0,8">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock x:Name="TxtGPACount" Grid.Column="0"
                                 FontSize="11" Foreground="#6B7280" VerticalAlignment="Center"/>
                      <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <Button x:Name="BtnGPAFilterClear" Content="&#x2715;"
                                Style="{StaticResource PrimaryBtn}" ToolTip="Clear filter"
                                IsEnabled="False" Height="30" Width="30" FontSize="12"
                                Padding="0" Margin="0,0,4,0"/>
                        <TextBox x:Name="TxtGPAFilter" Width="150" Height="30" FontSize="12"
                                VerticalContentAlignment="Center" Padding="6,0" Margin="0,0,6,0"
                                IsEnabled="False" Tag="keyword filter..."/>
                        <Button x:Name="BtnGPAFilter" Content="&#x1F50D; Filter"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,12,0"/>
                        <Button x:Name="BtnGPACopyValue" Content="Copy Value"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnGPACopyRow" Content="Copy Row"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnGPACopyAll" Content="Copy All"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnGPAExportXlsx" Content="Export XLSX"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4"/>
                      </StackPanel>
                    </Grid>
                    <DataGrid x:Name="DgGPAResults"
                              AutoGenerateColumns="False" IsReadOnly="True"
                              SelectionMode="Extended" SelectionUnit="CellOrRowHeader"
                              HeadersVisibility="Column"
                              GridLinesVisibility="All"
                              VerticalGridLinesBrush="#555555"
                              HorizontalGridLinesBrush="#D0D0D0"
                              RowBackground="White" AlternatingRowBackground="#F5F5FF"
                              BorderBrush="#B0B0D0" BorderThickness="1"
                              CanUserReorderColumns="False" CanUserResizeRows="False" CanUserResizeColumns="True"
                              CanUserSortColumns="True" MaxHeight="420" FontSize="12"
                              ScrollViewer.HorizontalScrollBarVisibility="Auto">
                      <DataGrid.ColumnHeaderStyle>
                        <Style TargetType="DataGridColumnHeader">
                          <Setter Property="FontWeight" Value="SemiBold"/>
                          <Setter Property="Background" Value="#E8EAF6"/>
                          <Setter Property="Foreground" Value="#1A237E"/>
                          <Setter Property="Padding"    Value="8,5"/>
                          <Setter Property="BorderBrush" Value="#555555"/>
                          <Setter Property="BorderThickness" Value="0,0,1,1"/>
                        </Style>
                      </DataGrid.ColumnHeaderStyle>
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Policy Name"     Binding="{Binding PolicyName}"    Width="3*" MinWidth="200"/>
                        <DataGridTextColumn Header="Type"            Binding="{Binding PolicyType}"    Width="2*" MinWidth="150"/>
                        <DataGridTextColumn Header="Sub-Type"        Binding="{Binding PolicySubType}" Width="2*" MinWidth="150"/>
                        <DataGridTextColumn Header="OS Platform"     Binding="{Binding OSPlatform}"    Width="*"  MinWidth="100"/>
                        <DataGridTextColumn Header="Included Group"  Binding="{Binding IncludedGroup}" Width="2*" MinWidth="160"/>
                        <DataGridTextColumn Header="Excluded Group"  Binding="{Binding ExcludedGroup}" Width="2*" MinWidth="160"/>
                        <DataGridTextColumn Header="Include Filter"  Binding="{Binding IncludeFilter}" Width="2*" MinWidth="150"/>
                        <DataGridTextColumn Header="Exclude Filter"  Binding="{Binding ExcludeFilter}" Width="2*" MinWidth="150"/>
                        <DataGridTextColumn Header="Description"     Binding="{Binding Description}"    Width="200" MaxWidth="250">
                          <DataGridTextColumn.ElementStyle>
                            <Style TargetType="TextBlock">
                              <Setter Property="TextWrapping" Value="NoWrap"/>
                              <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
                            </Style>
                          </DataGridTextColumn.ElementStyle>
                        </DataGridTextColumn>
                        <DataGridTextColumn Header="Deployment Mode" Binding="{Binding DeploymentMode}" Width="*" MinWidth="110"/>
                        <DataGridTextColumn Header="Priority" Binding="{Binding Priority}" Width="80" MinWidth="60"/>
                      </DataGrid.Columns>
                    </DataGrid>
                  </StackPanel>
                </Border>

              </StackPanel>
            </Border>


            <!-- === COMPARE GROUPS panel === -->
            <Border x:Name="PanelCompareGroups" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Compare Groups" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,4"
                           Foreground="$($DevConfig.ColorPrimary)"/>
                <TextBlock Text="Query and compare membership across multiple groups. Enter display names or Object IDs, one per line."
                           FontSize="11" Foreground="$($DevConfig.ColorTextMuted)" Margin="0,0,0,16" TextWrapping="Wrap"/>

                <!-- Group input -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="GROUPS  *  (one per line  &#x2014;  display name or Object ID)" Style="{StaticResource FieldLabel}"/>
                  <Grid Margin="0,4,0,0">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="8"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtCGGroupInput" Grid.Column="0" Style="{StaticResource InputBox}"
                             Height="80" TextWrapping="Wrap" AcceptsReturn="True"
                             VerticalScrollBarVisibility="Auto"
                             Tag="SG-Marketing-All&#10;SG-Sales-All&#10;SG-IT-All"/>
                    <StackPanel Grid.Column="2" Orientation="Vertical" VerticalAlignment="Top">
                      <Button x:Name="BtnCGCommon"   Content="Common Members"
                              Style="{StaticResource PrimaryBtn}" Width="140" Margin="0,0,0,6"/>
                      <Button x:Name="BtnCGDistinct" Content="Distinct Members"
                              Style="{StaticResource PrimaryBtn}" Width="140"/>
                    </StackPanel>
                  </Grid>
                </StackPanel>

                <!-- Progress overlay -->
                <Border x:Name="PnlCGProgress" Visibility="Collapsed"
                        Background="#1A2E1A" CornerRadius="6" Padding="16,12" Margin="0,0,0,12">
                  <StackPanel>
                    <TextBlock x:Name="TxtCGProgressMsg" Text="Querying groups..."
                               Foreground="#A3E4C0" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <ProgressBar x:Name="PbCGProgress" Height="6" IsIndeterminate="True"
                                 Background="#2A4A2A" Foreground="#22C55E" BorderThickness="0"/>
                    <TextBlock x:Name="TxtCGProgressDetail" Text=""
                               Foreground="#6BB88B" FontSize="11" Margin="0,5,0,0"/>
                  </StackPanel>
                </Border>

                <!-- Results area -->
                <Border x:Name="PnlCGResults" Visibility="Collapsed">
                  <StackPanel>
                    <Grid Margin="0,0,0,8">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock x:Name="TxtCGCount" Grid.Column="0"
                                 FontSize="11" Foreground="#6B7280" VerticalAlignment="Center"/>
                      <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <Button x:Name="BtnCGFilterClear" Content="&#x2715;"
                                Style="{StaticResource PrimaryBtn}" ToolTip="Clear filter"
                                IsEnabled="False" Height="30" Width="30" FontSize="12"
                                Padding="0" Margin="0,0,4,0"/>
                        <TextBox x:Name="TxtCGFilter" Width="150" Height="30" FontSize="12"
                                VerticalContentAlignment="Center" Padding="6,0" Margin="0,0,6,0"
                                IsEnabled="False" Tag="keyword filter..."/>
                        <Button x:Name="BtnCGFilter" Content="&#x1F50D; Filter"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,12,0"/>
                        <Button x:Name="BtnCGCopyValue" Content="Copy Value"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnCGCopyRow" Content="Copy Row"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnCGCopyAll" Content="Copy All"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnCGExportXlsx" Content="Export XLSX"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4"/>
                      </StackPanel>
                    </Grid>
                    <DataGrid x:Name="DgCGResults"
                              AutoGenerateColumns="False" IsReadOnly="True"
                              SelectionMode="Extended" SelectionUnit="CellOrRowHeader"
                              HeadersVisibility="Column"
                              GridLinesVisibility="All"
                              VerticalGridLinesBrush="#555555"
                              HorizontalGridLinesBrush="#D0D0D0"
                              RowBackground="White" AlternatingRowBackground="#FFF8ED"
                              BorderBrush="#F5C96A" BorderThickness="1"
                              CanUserReorderColumns="False" CanUserResizeRows="False" CanUserResizeColumns="True"
                              CanUserSortColumns="True" MaxHeight="420" FontSize="12"
                              ScrollViewer.HorizontalScrollBarVisibility="Auto">
                      <DataGrid.ColumnHeaderStyle>
                        <Style TargetType="DataGridColumnHeader">
                          <Setter Property="FontWeight" Value="SemiBold"/>
                          <Setter Property="Background" Value="#FEF3C7"/>
                          <Setter Property="Foreground" Value="#92400E"/>
                          <Setter Property="Padding"    Value="8,5"/>
                          <Setter Property="BorderBrush" Value="#555555"/>
                          <Setter Property="BorderThickness" Value="0,0,1,1"/>
                        </Style>
                      </DataGrid.ColumnHeaderStyle>
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Object"       Binding="{Binding DisplayName}" Width="2*" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Object ID"    Binding="{Binding ObjectID}"    Width="2*" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Object Type"  Binding="{Binding ObjectType}"  Width="*"  IsReadOnly="True"/>
                        <DataGridTextColumn Header="Source Group" Binding="{Binding SourceGroup}" Width="*"  IsReadOnly="True"/>
                      </DataGrid.Columns>
                    </DataGrid>
                  </StackPanel>
                </Border>
                <TextBlock x:Name="TxtCGNoResults" Visibility="Collapsed"
                           Text="No members found for the specified group(s)."
                           FontSize="12" Foreground="#6B7280" Margin="0,10,0,0"/>
              </StackPanel>
            </Border>

            <!-- === REMOVE GROUPS panel === -->
            <Border x:Name="PanelRemoveGroups" Style="{StaticResource Card}" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Remove Groups" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,4"
                           Foreground="$($DevConfig.ColorPrimary)"/>
                <TextBlock FontSize="11" Foreground="$($DevConfig.ColorTextMuted)" Margin="0,0,0,14" TextWrapping="Wrap"
                           Text="&#x26A0; Remove one or more Groups. Enter one group display name or Object ID per line."/>

                <!-- Input -->
                <StackPanel Margin="0,0,0,12">
                  <TextBlock Text="GROUP NAMES / OBJECT IDs  &#x2014;  one per line" Style="{StaticResource FieldLabel}"/>
                  <TextBox x:Name="TxtRGGroupList" Style="{StaticResource InputBox}"
                           Height="100" TextWrapping="Wrap" AcceptsReturn="True"
                           VerticalScrollBarVisibility="Auto"
                           Tag="GroupDisplayName&#10;00000000-0000-0000-0000-000000000001&#10;AnotherGroup"/>
                </StackPanel>

                <!-- Validate button -->
                <Button x:Name="BtnRGValidate" Content="&#x2714;  Validate Groups"
                        Style="{StaticResource PrimaryBtn}" HorizontalAlignment="Left" Width="160" Margin="4,4,0,12"/>

                <!-- Validate progress -->
                <Border x:Name="PnlRGProgress" Visibility="Collapsed"
                        Background="#1A2E1A" CornerRadius="6" Padding="16,12" Margin="0,0,0,12">
                  <StackPanel>
                    <TextBlock x:Name="TxtRGProgressMsg" Text="Validating groups..."
                               Foreground="#A3E4C0" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <ProgressBar x:Name="PbRGProgress" Height="6" IsIndeterminate="True"
                                 Background="#2A4A2A" Foreground="#22C55E" BorderThickness="0"/>
                    <TextBlock x:Name="TxtRGProgressDetail" Text=""
                               Foreground="#6BB88B" FontSize="11" Margin="0,5,0,0"/>
                  </StackPanel>
                </Border>

                <!-- Warning banner -->
                <Border x:Name="PnlRGWarning" Visibility="Collapsed"
                        Background="#FFF8E1" CornerRadius="6" Padding="12,10" Margin="0,0,0,12"
                        BorderBrush="#F59E0B" BorderThickness="1">
                  <StackPanel Orientation="Horizontal">
                    <TextBlock Text="&#x26A0;" FontSize="15" Foreground="#B45309" Margin="0,0,8,0" VerticalAlignment="Center"/>
                    <TextBlock x:Name="TxtRGWarning" FontSize="12" Foreground="#92400E" TextWrapping="Wrap"
                               VerticalAlignment="Center"/>
                  </StackPanel>
                </Border>

                <!-- Preview table + Execute button -->
                <Border x:Name="PnlRGPreview" Visibility="Collapsed">
                  <StackPanel>
                    <Grid Margin="0,0,0,8">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock x:Name="TxtRGPreviewCount" Grid.Column="0"
                                 FontSize="11" Foreground="#6B7280" VerticalAlignment="Center"/>
                      <Button x:Name="BtnRGExecute" Grid.Column="1"
                              Content="&#x1F5D1;  Remove Groups" IsEnabled="False"
                              Padding="14,8" FontSize="12" FontWeight="Bold" Cursor="Hand"
                              Foreground="White" BorderThickness="0">
                        <Button.Background>
                          <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                            <GradientStop Color="#DC2626" Offset="0"/>
                            <GradientStop Color="#B91C1C" Offset="1"/>
                          </LinearGradientBrush>
                        </Button.Background>
                        <Button.Template>
                          <ControlTemplate TargetType="Button">
                            <Border Background="{TemplateBinding Background}" CornerRadius="8"
                                    Padding="{TemplateBinding Padding}" x:Name="RGExecBg">
                              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                            <ControlTemplate.Triggers>
                              <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="RGExecBg" Property="Opacity" Value="0.45"/>
                              </Trigger>
                              <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="RGExecBg" Property="Opacity" Value="0.82"/>
                              </Trigger>
                            </ControlTemplate.Triggers>
                          </ControlTemplate>
                        </Button.Template>
                      </Button>
                    </Grid>
                    <DataGrid x:Name="DgRGPreview"
                              AutoGenerateColumns="False" IsReadOnly="True"
                              SelectionMode="Single" SelectionUnit="FullRow"
                              HeadersVisibility="Column"
                              GridLinesVisibility="All"
                              VerticalGridLinesBrush="#555555"
                              HorizontalGridLinesBrush="#D0D0D0"
                              RowBackground="White" AlternatingRowBackground="#F5FBF5"
                              BorderBrush="#B0D0B0" BorderThickness="1"
                              CanUserReorderColumns="False" CanUserResizeRows="False" CanUserResizeColumns="True"
                              CanUserSortColumns="True" MaxHeight="280" FontSize="12"
                              ScrollViewer.HorizontalScrollBarVisibility="Auto">
                      <DataGrid.RowStyle>
                        <Style TargetType="DataGridRow">
                          <Style.Triggers>
                            <DataTrigger Binding="{Binding HasMembers}" Value="True">
                              <Setter Property="Background" Value="#FFF8E1"/>
                              <Setter Property="FontWeight" Value="SemiBold"/>
                            </DataTrigger>
                          </Style.Triggers>
                        </Style>
                      </DataGrid.RowStyle>
                      <DataGrid.ColumnHeaderStyle>
                        <Style TargetType="DataGridColumnHeader">
                          <Setter Property="FontWeight" Value="SemiBold"/>
                          <Setter Property="Background" Value="#E8F5E9"/>
                          <Setter Property="Foreground" Value="#14532D"/>
                          <Setter Property="Padding"    Value="8,5"/>
                          <Setter Property="BorderBrush" Value="#555555"/>
                          <Setter Property="BorderThickness" Value="0,0,1,1"/>
                        </Style>
                      </DataGrid.ColumnHeaderStyle>
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="3*" MinWidth="160"/>
                        <DataGridTextColumn Header="Group ID"     Binding="{Binding GroupID}"     Width="3*" MinWidth="130"/>
                        <DataGridTextColumn Header="Group Type"   Binding="{Binding GroupType}"   Width="*"  MinWidth="120"/>
                        <DataGridTextColumn Header="Members"      Binding="{Binding MemberCount}" Width="*"  MinWidth="80"/>
                      </DataGrid.Columns>
                    </DataGrid>
                  </StackPanel>
                </Border>

                <!-- Execution progress -->
                <Border x:Name="PnlRGExecProgress" Visibility="Collapsed"
                        Background="#2A1A1A" CornerRadius="6" Padding="16,12" Margin="0,12,0,12">
                  <StackPanel>
                    <TextBlock x:Name="TxtRGExecProgressMsg" Text="Removing groups..."
                               Foreground="#FCA5A5" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <ProgressBar x:Name="PbRGExecProgress" Height="6" IsIndeterminate="True"
                                 Background="#4A2A2A" Foreground="#EF4444" BorderThickness="0"/>
                    <TextBlock x:Name="TxtRGExecProgressDetail" Text=""
                               Foreground="#F87171" FontSize="11" Margin="0,5,0,0"/>
                  </StackPanel>
                </Border>

                <!-- Results table -->
                <Border x:Name="PnlRGResults" Visibility="Collapsed">
                  <StackPanel>
                    <Grid Margin="0,0,0,8">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock x:Name="TxtRGResultCount" Grid.Column="0"
                                 FontSize="11" Foreground="#6B7280" VerticalAlignment="Center"/>
                      <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <Button x:Name="BtnRGCopyAll" Content="Copy All"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4" Margin="0,0,6,0"/>
                        <Button x:Name="BtnRGExportXlsx" Content="Export XLSX"
                                Style="{StaticResource PrimaryBtn}"
                                IsEnabled="False" Height="30" FontSize="12" FontWeight="SemiBold"
                                Padding="12,4"/>
                      </StackPanel>
                    </Grid>
                    <DataGrid x:Name="DgRGResults"
                              AutoGenerateColumns="False" IsReadOnly="True"
                              SelectionMode="Single" SelectionUnit="FullRow"
                              HeadersVisibility="Column"
                              GridLinesVisibility="All"
                              VerticalGridLinesBrush="#555555"
                              HorizontalGridLinesBrush="#D0D0D0"
                              RowBackground="White" AlternatingRowBackground="#F5F5F5"
                              BorderBrush="#B0D0B0" BorderThickness="1"
                              CanUserReorderColumns="False" CanUserResizeRows="False" CanUserResizeColumns="True"
                              CanUserSortColumns="True" MaxHeight="280" FontSize="12"
                              ScrollViewer.HorizontalScrollBarVisibility="Auto">
                      <DataGrid.RowStyle>
                        <Style TargetType="DataGridRow">
                          <Style.Triggers>
                            <DataTrigger Binding="{Binding IsSuccess}" Value="True">
                              <Setter Property="Background" Value="#F0FDF4"/>
                            </DataTrigger>
                            <DataTrigger Binding="{Binding IsSuccess}" Value="False">
                              <Setter Property="Background" Value="#FEF2F2"/>
                            </DataTrigger>
                          </Style.Triggers>
                        </Style>
                      </DataGrid.RowStyle>
                      <DataGrid.ColumnHeaderStyle>
                        <Style TargetType="DataGridColumnHeader">
                          <Setter Property="FontWeight" Value="SemiBold"/>
                          <Setter Property="Background" Value="#E8F5E9"/>
                          <Setter Property="Foreground" Value="#14532D"/>
                          <Setter Property="Padding"    Value="8,5"/>
                          <Setter Property="BorderBrush" Value="#555555"/>
                          <Setter Property="BorderThickness" Value="0,0,1,1"/>
                        </Style>
                      </DataGrid.ColumnHeaderStyle>
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="3*" MinWidth="160"/>
                        <DataGridTextColumn Header="Group ID"     Binding="{Binding GroupID}"     Width="3*" MinWidth="130"/>
                        <DataGridTextColumn Header="Status"       Binding="{Binding Status}"      Width="*"  MinWidth="80"/>
                        <DataGridTextColumn Header="Message"      Binding="{Binding Message}"     Width="4*" MinWidth="200"/>
                      </DataGrid.Columns>
                    </DataGrid>
                  </StackPanel>
                </Border>

                <TextBlock x:Name="TxtRGNoResults" Visibility="Collapsed"
                           Text="No groups could be resolved from the provided input. Check the names / IDs and try again."
                           FontSize="12" Foreground="#6B7280" Margin="0,12,0,0" TextWrapping="Wrap"/>
              </StackPanel>
            </Border>
          </Grid>
        </ScrollViewer>

      </Grid>
    </DockPanel>

    <!-- ================================================================== -->
    <!-- Session Notes floating overlay  (in-memory only, not persisted)    -->
    <!-- ================================================================== -->
    <Grid x:Name="OverlaySessionNotes" Visibility="Collapsed">
      <Rectangle x:Name="SnDimmer" Fill="#80000000" IsHitTestVisible="True"/>
      <Border Background="White" CornerRadius="12"
              BorderBrush="#E5E7EB" BorderThickness="1"
              Width="520" MaxHeight="480"
              HorizontalAlignment="Center" VerticalAlignment="Center">
        <Border.Effect>
          <DropShadowEffect BlurRadius="24" ShadowDepth="6" Direction="270"
                            Color="#000000" Opacity="0.28"/>
        </Border.Effect>
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <!-- Card header -->
          <Border Grid.Row="0" Background="#F9FAFB" CornerRadius="12,12,0,0" Padding="18,12">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                <TextBlock Text="&#x1F5D2;" FontSize="17" Margin="0,0,8,0" VerticalAlignment="Center"/>
                <TextBlock Text="Session Notes" FontSize="14" FontWeight="SemiBold"
                           Foreground="#111827" VerticalAlignment="Center"/>
                <TextBlock Text="  not saved to disk" FontSize="11"
                           Foreground="#9CA3AF" VerticalAlignment="Center" Margin="4,0,0,0"/>
              </StackPanel>
              <Button x:Name="BtnSessionNotesClose" Grid.Column="1"
                      Content="&#x2715;" FontSize="13" FontWeight="Bold"
                      Background="Transparent" BorderThickness="0"
                      Foreground="#6B7280" Cursor="Hand"
                      Width="30" Height="30" Padding="0"
                      VerticalContentAlignment="Center" HorizontalContentAlignment="Center"/>
            </Grid>
          </Border>
          <!-- Note text area -->
          <TextBox x:Name="TxtSessionNotes" Grid.Row="1"
                   TextWrapping="Wrap" AcceptsReturn="True" AcceptsTab="True"
                   VerticalScrollBarVisibility="Auto"
                   BorderThickness="0,1,0,1" BorderBrush="#E5E7EB"
                   Padding="18,12" FontSize="12" FontFamily="Consolas"
                   Foreground="#1F2937" Background="White"
                   MinHeight="260" MaxHeight="360"
                   Tag="Type your session notes here&#x2026;"/>
          <!-- Card footer -->
          <Border Grid.Row="2" Background="#F9FAFB" CornerRadius="0,0,12,12" Padding="18,8">
            <TextBlock Text="Notes are discarded when the tool is closed."
                       FontSize="11" Foreground="#9CA3AF" HorizontalAlignment="Center"/>
          </Border>
        </Grid>
      </Border>
    </Grid>
  </Grid>
</Window>
"@

#endregion XAML
Write-ErrorLog "Checkpoint 2: XAML defined OK."
Write-EarlyLog "Checkpoint 2: XAML defined OK."


# ============================================================
#region UI HELPER FUNCTIONS
# ============================================================

function Get-Timestamp { (Get-Date).ToString("HH:mm:ss") }

function Write-VerboseLog {
    param(
        [string]$Message,
        [ValidateSet('Info','Success','Warning','Error','Action')]
        [string]$Level = 'Info'
    )

    $color = switch ($Level) {
        'Info'    { $DevConfig.LogColorInfo }
        'Success' { $DevConfig.LogColorSuccess }
        'Warning' { $DevConfig.LogColorWarning }
        'Error'   { $DevConfig.LogColorError }
        'Action'  { $DevConfig.LogColorAction }
    }

    $prefix = switch ($Level) {
        'Info'    { '  ' }
        'Success' { "$([char]0x2713) " }
        'Warning' { '[!] ' }
        'Error'   { "$([char]0x2717) " }
        'Action'  { '> ' }
    }

    $ts   = Get-Timestamp
    $line = "[$ts] $prefix$Message"

    # Write to file
    if ($script:LogFile) {
        try { Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue }
        catch {}
    }

    # Write to RichTextBox on UI thread
    $script:Window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Normal, [Action]{
        try {
            $para = [System.Windows.Documents.Paragraph]::new()
            $run  = [System.Windows.Documents.Run]::new($line)
            $run.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                [System.Windows.Media.ColorConverter]::ConvertFromString($color)
            )
            $para.Inlines.Add($run)
            $para.Margin = [System.Windows.Thickness]::new(0,0,0,0)
            $script:RtbLog.Document.Blocks.Add($para)
            $script:RtbLog.ScrollToEnd()
        } catch {}
    })
}

function Show-Panel {
    param([string]$Name)
    $panels = @('PanelSearch','PanelListMembers','PanelObjectMembership','PanelFindGroupsByOwners','PanelWelcome','PanelCreate','PanelMembership','PanelExport',
                 'PanelDynamic','PanelRename','PanelOwner','PanelUserDevices','PanelFindCommon','PanelFindDistinct','PanelGetDeviceInfo','PanelGetDiscoveredApps','PanelGetPolicyAssignments','PanelCompareGroups','PanelRemoveGroups')
    foreach ($p in $panels) {
        $el = $script:Window.FindName($p)
        if ($el) { $el.Visibility = if ($p -eq $Name) { 'Visible' } else { 'Collapsed' } }
    }
    # Show Clear Inputs button on any panel except Welcome
    $clearBtn = $script:Window.FindName('BtnClearInputs')
    if ($clearBtn) {
        $clearBtn.Visibility = if ($Name -ne 'PanelWelcome') { 'Visible' } else { 'Collapsed' }
    }
}

function Show-Notification {
    param([string]$Message, [string]$BgColor = '#FFF3CD', [string]$FgColor = '#7A4800')
    $script:Window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Normal, [Action]{
        $script:NotifStrip.Background = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString($BgColor))
        $script:TxtNotif.Foreground = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.ColorConverter]::ConvertFromString($FgColor))
        $script:TxtNotif.Text    = $Message
        $script:NotifStrip.Visibility = 'Visible'
    })
}

function Hide-Notification {
    $script:Window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Normal, [Action]{
        $script:NotifStrip.Visibility = 'Collapsed'
    })
}

function Update-StatusBar {
    param([string]$Text)
    $script:Window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Normal, [Action]{
        $script:TxtStatusBar.Text = $Text
    })
}

function Show-ConfirmationDialog {
    param(
        [string]$Title,
        [string]$OperationLabel,
        [string]$TargetGroup,
        [string]$TargetGroupId,
        [object[]]$ValidObjects,
        [object[]]$InvalidEntries,
        [string]$ExtraInfo = ""
    )

    $validCount   = $ValidObjects.Count
    $invalidCount = $InvalidEntries.Count

    # Build summary text
    $sb = [System.Text.StringBuilder]::new()
    $null = $sb.AppendLine("OPERATION : $OperationLabel")
    $null = $sb.AppendLine("GROUP     : $TargetGroup")
    $null = $sb.AppendLine("GROUP ID  : $TargetGroupId")
    if ($ExtraInfo) { $null = $sb.AppendLine($ExtraInfo) }
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("Valid objects  : $validCount")
    if ($invalidCount -gt 0) {
        $null = $sb.AppendLine("Invalid entries: $invalidCount  (will be SKIPPED)")
        $null = $sb.AppendLine("")
        $null = $sb.AppendLine("-- INVALID (not found in Entra) --")
        foreach ($inv in $InvalidEntries) { $null = $sb.AppendLine("  $([char]0x2717) $inv") }
    }
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("-- VALID (will be processed) --")
    foreach ($obj in $ValidObjects) {
        $dispLabel = if ($obj.Type -eq 'User') { $obj.Original } else { $obj.DisplayName }
        $null = $sb.AppendLine("  $([char]0x2713) $dispLabel  [$($obj.Type)]  ($($obj.Id))")
    }

    [xml]$dlgXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Confirm Action" Width="620" Height="480"
        WindowStartupLocation="CenterOwner"
        Background="$($DevConfig.ColorBackground)"
        FontFamily="$($DevConfig.FontFamily)" FontSize="13"
        ResizeMode="CanResize">
  <Grid Margin="16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <StackPanel Grid.Row="0" Margin="0,0,0,10">
      <TextBlock Text="$Title" FontSize="15" FontWeight="SemiBold"
                 Foreground="$($DevConfig.ColorPrimary)"/>
      <TextBlock Text="Review the summary below before proceeding."
                 FontSize="11" Foreground="$($DevConfig.ColorTextMuted)" Margin="0,4,0,0"/>
    </StackPanel>

    <Border Grid.Row="1" BorderBrush="$($DevConfig.ColorBorder)" BorderThickness="1"
            CornerRadius="6" Background="$($DevConfig.ColorSurface)">
      <TextBox x:Name="SummaryBox" IsReadOnly="True" TextWrapping="Wrap"
               Background="Transparent" BorderThickness="0"
               FontFamily="Consolas" FontSize="11"
               Foreground="$($DevConfig.ColorText)"
               VerticalScrollBarVisibility="Auto"
               HorizontalScrollBarVisibility="Disabled"
               Padding="10,8"/>
    </Border>

    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right"
                Margin="0,12,0,0">
      <Button x:Name="BtnDlgCancel" Content="Cancel" Width="90" Margin="0,0,10,0"
              Background="Transparent"
              Foreground="$($DevConfig.ColorText)"
              BorderBrush="$($DevConfig.ColorBorder)" BorderThickness="1"
              Padding="8,6" Cursor="Hand"/>
      <Button x:Name="BtnDlgConfirm" Content="&#x2713; Confirm &amp; Execute" Width="180"
              Background="$($DevConfig.ColorAccent)" Foreground="White"
              BorderThickness="0" Padding="10,7" FontWeight="SemiBold" Cursor="Hand"/>
    </StackPanel>
  </Grid>
</Window>
"@

    $dlgReader = [System.Xml.XmlNodeReader]::new($dlgXaml)
    $dlg = [Windows.Markup.XamlReader]::Load($dlgReader)
    $dlg.Owner = $script:Window

    $dlg.FindName('SummaryBox').Text = $sb.ToString()

    $script:DlgResult = $false
    $dlg.FindName('BtnDlgCancel').Add_Click({ $script:DlgResult = $false; $dlg.Close() })
    $dlg.FindName('BtnDlgConfirm').Add_Click({ $script:DlgResult = $true; $dlg.Close() })

    $dlg.ShowDialog() | Out-Null
    return $script:DlgResult
}

#endregion
Write-ErrorLog "Checkpoint 3: UI helpers defined OK."


# ============================================================
#region GRAPH HELPER FUNCTIONS
# ============================================================

function Invoke-GraphGet {
    param([string]$Uri)
    $results = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    do {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next
        if ($resp['value']) { $results.AddRange([object[]]$resp['value']) }
        $next = if ($resp.ContainsKey('@odata.nextLink')) { $resp['@odata.nextLink'] } else { $null }
    } while ($next)
    return ,$results
}

function Resolve-GroupByNameOrId {
    param([string]$GroupEntry)
    $trimmed = $GroupEntry.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) { return $null }

    # GUID pattern → look up by ID
    if ($trimmed -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
        try {
            $g = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$trimmed`?`$select=id,displayName,groupTypes,membershipRule,mailEnabled,securityEnabled"
            return [PSCustomObject]@{ Id = $g['id']; DisplayName = $g['displayName']; GroupTypes = $g['groupTypes']; MembershipRule = $g['membershipRule'] }
        } catch { return $null }
    }

    # Name search
    $enc = [Uri]::EscapeDataString($trimmed)
    try {
        $safe_trimmed = $trimmed -replace "'","''"  # PS5-safe: extracted to avoid nested double-quotes in string
        $resp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe_trimmed'`&`$select=id,displayName,groupTypes,membershipRule,mailEnabled,securityEnabled`&`$top=10"
        if ($resp['value'] -and $resp['value'].Count -gt 0) {
            $g = $resp['value'][0]
            return [PSCustomObject]@{ Id = $g['id']; DisplayName = $g['displayName']; GroupTypes = $g['groupTypes']; MembershipRule = $g['membershipRule'] }
        }
    } catch {}
    return $null
}

function Search-Groups {
    param([string]$Query)
    if ([string]::IsNullOrWhiteSpace($Query)) { return @() }
    $enc = $Query.Trim() -replace "'","''"
    try {
        $resp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=startswith(displayName,'$enc')`&`$select=id,displayName,groupTypes`&`$top=15"
        return $resp['value']
    } catch { return @() }
}

function Resolve-MemberEntry {
    param([string]$Entry)

    # Returns [PSCustomObject]@{ Id; DisplayName; Type; Found }
    # Type = 'User' | 'Group' | 'Device' | 'Unknown'

    $trimmed = $Entry.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) { return $null }

    $isGuid = $trimmed -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
    $isUpn  = $trimmed -match '^[^@\s]+@[^@\s]+\.[^@\s]+$'

    # ---- User (UPN only) ----
    if ($isUpn) {
        try {
            $u = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$([Uri]::EscapeDataString($trimmed))?`$select=id,displayName,userPrincipalName"
            return [PSCustomObject]@{ Id = $u['id']; DisplayName = $u['displayName']; Type = 'User'; Found = $true; Original = $trimmed }
        } catch {
            return [PSCustomObject]@{ Id = $null; DisplayName = $trimmed; Type = 'User'; Found = $false; Original = $trimmed }
        }
    }

    # ---- GUID  -  try group, then device ----
    if ($isGuid) {
        try {
            $g = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$trimmed`?`$select=id,displayName"
            return [PSCustomObject]@{ Id = $g['id']; DisplayName = $g['displayName']; Type = 'Group'; Found = $true; Original = $trimmed }
        } catch {}
        try {
            $d = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/devices/$trimmed`?`$select=id,displayName"
            return [PSCustomObject]@{ Id = $d['id']; DisplayName = $d['displayName']; Type = 'Device'; Found = $true; Original = $trimmed }
        } catch {}
        return [PSCustomObject]@{ Id = $null; DisplayName = $trimmed; Type = 'Unknown'; Found = $false; Original = $trimmed }
    }

    # ---- Display name  -  try group, then device ----
    $safeEntry = $trimmed -replace "'","''"
    try {
        $resp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safeEntry'`&`$select=id,displayName`&`$top=1"
        if ($resp['value'] -and $resp['value'].Count -gt 0) {
            $g = $resp['value'][0]
            return [PSCustomObject]@{ Id = $g['id']; DisplayName = $g['displayName']; Type = 'Group'; Found = $true; Original = $trimmed }
        }
    } catch {}
    try {
        $resp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=displayName eq '$safeEntry'`&`$select=id,displayName`&`$top=1"
        if ($resp['value'] -and $resp['value'].Count -gt 0) {
            $d = $resp['value'][0]
            return [PSCustomObject]@{ Id = $d['id']; DisplayName = $d['displayName']; Type = 'Device'; Found = $true; Original = $trimmed }
        }
    } catch {}

    return [PSCustomObject]@{ Id = $null; DisplayName = $trimmed; Type = 'Unknown'; Found = $false; Original = $trimmed }
}

function Validate-InputList {
    param([string[]]$Entries)
    # Use plain arrays  -  Generic.List serialises to Object[] across runspace
    # boundaries which causes AddRange type errors.
    $valid   = @()
    $invalid = @()

    foreach ($entry in $Entries) {
        $e = $entry.Trim()
        if ([string]::IsNullOrWhiteSpace($e)) { continue }
        Write-VerboseLog "Resolving: $e" -Level Info
        $result = Resolve-MemberEntry -Entry $e
        if ($null -eq $result) { continue }
        if ($result.Found) {
            Write-VerboseLog "  $([char]0x2713) $($result.DisplayName) [$($result.Type)] ($($result.Id))" -Level Success
            $valid += $result
        } else {
            Write-VerboseLog "  $([char]0x2717) Not found: $e" -Level Warning
            $invalid += $e
        }
    }
    return @{ Valid = $valid; Invalid = $invalid }
}

function Get-GroupMemberCount {
    param([string]$GroupId)
    # Requires ConsistencyLevel:eventual AND $count=true in URL
    # Silently returns $null on any error (e.g. missing permission)  -  never blocks UI
    try {
        $resp = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members/`$count?`$count=true" `
            -Headers @{ ConsistencyLevel = 'eventual' }
        return [int]$resp
    } catch { return $null }
}

function Get-GroupMembers {
    param([string]$GroupId)
    # Note: @odata.type must NOT be in $select  -  Graph returns it automatically.
    # Including it causes a BadRequest error.
    return Invoke-GraphGet -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members?`$select=id,displayName,userPrincipalName`&`$top=999"
}

function Add-MembersToGroup {
    param(
        [string]$GroupId,
        [object[]]$Members
    )
    $ok = 0; $fail = 0; $skipped = 0
    foreach ($m in $Members) {
        if ($null -ne $Shared -and $Shared['StopRequested']) {
            Write-VerboseLog "Stop requested  -  halting after $ok added." -Level Warning
            break
        }
        $body = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($m.Id)" }
        Write-VerboseLog "Adding $($m.DisplayName) [$($m.Type)]..." -Level Action
        try {
            $null = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members/`$ref" -Body ($body | ConvertTo-Json) -ContentType "application/json"
            Write-VerboseLog "Added: $($m.DisplayName)" -Level Success
            $ok++
        } catch {
            $errMsg    = $_.Exception.Message
            $errDetail = $_.ErrorDetails.Message
            if ($errMsg -match 'already exist' -or $errDetail -match 'already exist') {
                Write-VerboseLog "Skipping $($m.DisplayName) - already a member of this group" -Level Warning
                $skipped++
            } else {
                Write-VerboseLog "Failed to add $($m.DisplayName): $errMsg" -Level Error
                $fail++
            }
        }
    }
    return @{ Ok = $ok; Fail = $fail; Skipped = $skipped }
}

function Remove-MembersFromGroup {
    param(
        [string]$GroupId,
        [object[]]$Members
    )
    $ok = 0; $fail = 0
    foreach ($m in $Members) {
        if ($null -ne $Shared -and $Shared['StopRequested']) {
            Write-VerboseLog "Stop requested  -  halting after $ok removed." -Level Warning
            break
        }
        Write-VerboseLog "Removing $($m.DisplayName) [$($m.Type)]..." -Level Action
        try {
            $null = Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members/$($m.Id)/`$ref"
            Write-VerboseLog "Removed: $($m.DisplayName)" -Level Success
            $ok++
        } catch {
            Write-VerboseLog "Failed to remove $($m.DisplayName): $($_.Exception.Message)" -Level Error
            $fail++
        }
    }
    return @{ Ok = $ok; Fail = $fail }
}

function Export-GroupMembersToXlsx {
    param([string]$GroupId, [string]$GroupName, [string]$OutputPath)

    Write-VerboseLog "Fetching members for group: $GroupName" -Level Action
    $members = Get-GroupMembers -GroupId $GroupId
    Write-VerboseLog "Retrieved $($members.Count) members" -Level Info

    # Build rows
    $rows = foreach ($m in $members) {
        $odataType = if ($m.ContainsKey('@odata.type')) { $m['@odata.type'] } else { '' }
        $type = switch ($odataType) {
            '#microsoft.graph.user'   { 'User' }
            '#microsoft.graph.group'  { 'Group' }
            '#microsoft.graph.device' { 'Device' }
            default                   { if ($odataType) { $odataType } else { 'Unknown' } }
        }
        [PSCustomObject]@{
            DisplayName       = $m['displayName']
            UserPrincipalName = if ($m['userPrincipalName']) { $m['userPrincipalName'] } else { '-' }
            ObjectId          = $m['id']
            ObjectType        = $type
            GroupName         = $GroupName
            GroupId           = $GroupId
            ExportedAt        = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }
    }

    # Ensure folder exists
    $folder = Split-Path $OutputPath -Parent
    if ($folder -and -not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }

    # Try ImportExcel if available, else CSV fallback
    if (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue) {
        Import-Module ImportExcel -ErrorAction SilentlyContinue
        $rows | Export-Excel -Path $OutputPath -WorksheetName "Members" -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium2
        Write-VerboseLog "Exported $($rows.Count) members to XLSX: $OutputPath" -Level Success
    } else {
        # Fallback: save as CSV with .xlsx extension (user informed)
        $rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-VerboseLog "ImportExcel module not found  -  exported as CSV to: $OutputPath" -Level Warning
        Write-VerboseLog "Install ImportExcel: Install-Module ImportExcel -Scope CurrentUser" -Level Info
    }
    return $rows.Count
}

#endregion
Write-ErrorLog "Checkpoint 4: Graph helpers defined OK."


# ============================================================
#region BUILD & WIRE UI
# ============================================================

Write-EarlyLog "Loading XAML..."
Write-ErrorLog "Loading XAML..."
try {
    $reader = [System.Xml.XmlNodeReader]::new($Xaml)
    $script:Window = [Windows.Markup.XamlReader]::Load($reader)
    Write-ErrorLog "XAML loaded OK."
    Write-EarlyLog "XAML loaded OK."
} catch {
    Write-ErrorLog "XAML load failed: $($_.Exception.Message)"
    Write-EarlyLog "XAML load FAILED: $($_.Exception.Message)"
    try {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("XAML load failed:`n`n$($_.Exception.Message)","Error") | Out-Null
    } catch {}
    exit 1
}

# Pre-initialise all script-scope state variables (required by Set-StrictMode -Version Latest)
$script:SelectedGroup       = $null
$script:ExportSelectedGroup = $null
$script:DynSelectedGroup    = $null
$script:RenSelectedGroup     = $null
$script:OwnerGroupsTxt       = ''
$script:OwnerGroupsValidated = $null
$script:CurrentMemberOp      = 'Add'
$script:DlgResult            = $false
$script:SensitivityLabels    = @()   # sensitivity label fetch removed (API permission not available)
$script:DynGroupState        = $null   # populated by BtnDynValidate Work block; checked in Done block
$script:SearchKeyword        = ''     # keyword for BtnEntraSearch
$script:SearchTypes          = @{}    # type filter hashtable
$script:SearchResults        = $null  # List[PSCustomObject] from BtnEntraSearch
$script:SearchExactMatch      = $false  # exact-match-only filter for BtnEntraSearch
$script:LMAllResults         = $null  # List[PSCustomObject] from BtnLMQuery (all, pre-filter)
$script:LMGroupInput         = ''     # raw text from TxtLMGroupInput
$script:LMProgressTimer      = $null  # DispatcherTimer for progress bar updates
$script:CGAllResults         = $null  # List[PSCustomObject] from BtnCGCommon / BtnCGDistinct
$script:CGGroupInput         = ''     # raw text from TxtCGGroupInput
$script:CGProgressTimer      = $null  # DispatcherTimer for CG progress bar
$script:CreateSensLabelId    = ''
$script:AdminUnits            = @()   # Administrative Units fetched from Graph
$script:CreateSelectedAU      = $null # Selected AU for group creation (hashtable with Id, DisplayName)
$script:CreateSensLabelName  = ''
$script:PickerCallbacks     = @{}
$script:LogPaneVisible      = $true
$script:AuthTimer           = $null
$script:AuthPs              = $null
$script:AuthRs              = $null
$script:AuthHandle          = $null
$script:AuthDlgResult       = $false
$script:AuthDlgTenant       = ''
$script:AuthDlgClient       = ''
$script:BgPs                = $null
$script:BgRs                = $null
$script:BgHandle            = $null
$script:BgTimer             = $null
$script:BgDone              = $null
$script:BgBtn               = $null
$script:SharedBg            = $null
$script:BgStopped           = $false
$script:PimRoles            = @()
$script:PimQueue            = $null
$script:PimRs               = $null
$script:PimPs               = $null
$script:PimHandle           = $null
$script:PimTimer            = $null
$script:PimAutoTimer        = $null
$script:BladeExpanded       = $false
$script:CreateParams        = $null
$script:CreateValidated     = $null
$script:CreateNewGroupId    = $null
$script:MemberParams        = $null
$script:MemberValidated     = $null
$script:MemberExecResult    = $null
$script:ExportOutPath       = ''
$script:ExportCount         = 0
$script:DynRule             = ''
$script:RenNewName          = ''
$script:RenExistingId       = $null
$script:OwnerUpnsTxt        = ''
$script:OwnerValidated      = $null
$script:OwnerExecResult     = $null
$script:UDSelectedGroup     = $null
$script:UDParams            = $null
$script:UDValidated         = $null
$script:UDExecResult        = $null
$script:FCParams            = $null
$script:FCResult            = $null
$script:FDParams            = $null
$script:FDResult            = $null
$script:GPAParams         = $null  # GPA: query params passed to background work block
$script:GPAResult         = $null  # GPA: result rows from background query
$script:RGGroupList          = $null  # RG: raw group input text
$script:RGValidated          = $null  # RG: validated group objects for preview
$script:RGExecResult         = $null  # RG: per-group deletion results
$script:AuthCancelled       = $false
$script:AuthStartTime       = [datetime]::UtcNow
$script:AuthTimeoutSec      = 120
$script:BladeDevExpanded    = $false
$script:GDIParams           = $null
$script:FGONoOwnerMode      = $false
$script:GdiQueue           = $null
$script:GdiStop            = $null
$script:OmStop             = $null
$script:DaStop             = $null
$script:GdiRs              = $null
$script:OmQueue            = $null
$script:OmRs               = $null
$script:OmPs               = $null
$script:OmHandle           = $null
$script:OmTimer            = $null
$script:OMParams           = $null
$script:DaQueue            = $null
$script:DaRs               = $null
$script:DaPs               = $null
$script:DaHandle           = $null
$script:DaTimer            = $null
$script:SearchKeyword      = ''
$script:SearchEntries      = @()
$script:SearchExactMatch   = $false
$script:SearchGetManager   = $false
$script:SearchTypes        = $null
$script:SearchResults      = $null
$script:SearchAllData      = $null
$script:LMAllData          = $null
$script:OMAllData          = $null
$script:FCAllData          = $null
$script:FDAllData          = $null
$script:GDIAllData         = $null
$script:DAAllData          = $null
$script:GPAAllData         = $null
$script:CGAllData          = $null
$script:DAParams           = $null
$script:GdiPs              = $null
$script:GdiHandle          = $null
$script:GdiTimer           = $null
$script:FGOStop            = $null
$script:FGOQueue           = $null
$script:FGORs              = $null
$script:FGOPs              = $null
$script:FGOHandle          = $null
$script:FGOTimer           = $null
$script:FGOAllData         = $null

# Cache named elements
$script:TxtStatusBar       = $script:Window.FindName('TxtStatusBar')
$script:NotifStrip         = $script:Window.FindName('NotifStrip')
$script:TxtNotif           = $script:Window.FindName('TxtNotif')
$script:PimStrip           = $script:Window.FindName('PimStrip')
$script:SpPimRoles         = $script:Window.FindName('SpPimRoles')
$script:BgImage            = $script:Window.FindName('BgImage')
$script:LogoImg            = $script:Window.FindName('LogoImg')
$script:RtbLog             = $script:Window.FindName('RtbLog')
$script:LogPane            = $script:Window.FindName('LogPane')


# ── Settings panel toggle ─────────────────────────────────────────────────────


# ── Log pane min/max height enforcement via GridSplitter drag ─────────────────
$script:LogPaneMinH = 44
$script:LogPaneMaxH = 260
$script:LogPane.Add_SizeChanged({
    $h = $script:LogPane.ActualHeight
    if ($h -lt $script:LogPaneMinH -and $script:LogPaneVisible) {
        $script:LogPane.Height = $script:LogPaneMinH
    } elseif ($h -gt $script:LogPaneMaxH) {
        $script:LogPane.Height = $script:LogPaneMaxH
    }
})


# ── Helper: load a bitmap from a local path or web URL ──
# BitmapFrame.Create(Uri) handles both file:// and http(s):// without extra modules.
# Local path takes priority over URL.
# ÄÄ SVG -> RasterBitmap  (no external libraries required) ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ
# Renders SVG via a hidden WPF WebBrowser (IE/Trident engine), captures with
# RenderTargetBitmap, then closes the window.  Handles simple/flat logo SVGs.
function Convert-SvgToBitmap {
    param([string]$SvgContent, [int]$Width = 120, [int]$Height = 40)
    try {
        # Add Win32 interop helpers (once per session)
        if (-not ([System.Management.Automation.PSTypeName]'SvgCapture').Type) {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class SvgCapture {
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, int flags);
    [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern int  ReleaseDC(IntPtr hwnd, IntPtr hdc);
    [DllImport("gdi32.dll")]  public static extern IntPtr CreateCompatibleDC(IntPtr hdc);
    [DllImport("gdi32.dll")]  public static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int w, int h);
    [DllImport("gdi32.dll")]  public static extern IntPtr SelectObject(IntPtr hdc, IntPtr obj);
    [DllImport("gdi32.dll")]  public static extern bool DeleteDC(IntPtr hdc);
    [DllImport("gdi32.dll")]  public static extern bool DeleteObject(IntPtr obj);
}
"@
        }

        $svgB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($SvgContent))
        $html   = "<!DOCTYPE html><html><body style='margin:0;padding:0;overflow:hidden;" +
                  "background:transparent;'><img src='data:image/svg+xml;base64,$svgB64' " +
                  "width='$Width' height='$Height'/></body></html>"

        $wb = [System.Windows.Controls.WebBrowser]::new()
        $wb.Width  = $Width
        $wb.Height = $Height
        $null = $wb.Add_LoadCompleted({ $script:_SvgLoadDone = $true })

        $hostWin = [System.Windows.Window]::new()
        $hostWin.WindowStyle   = 'None'
        $hostWin.ShowInTaskbar = $false
        $hostWin.AllowsTransparency = $false
        $hostWin.Left = -32000; $hostWin.Top = -32000
        $hostWin.Width = $Width; $hostWin.Height = $Height
        $hostWin.Content = $wb
        $script:_SvgLoadDone = $false
        $null = $hostWin.Show()
        $null = $wb.NavigateToString($html)

        # Wait up to 4 s for LoadCompleted, pumping dispatcher each cycle
        $deadline = (Get-Date).AddSeconds(4)
        while (-not $script:_SvgLoadDone -and (Get-Date) -lt $deadline) {
            $null = [System.Windows.Application]::Current.Dispatcher.Invoke(
                [Action]{}, [System.Windows.Threading.DispatcherPriority]::ApplicationIdle)
            Start-Sleep -Milliseconds 80
        }
        # Extra render pump
        $null = [System.Windows.Application]::Current.Dispatcher.Invoke(
            [Action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
        Start-Sleep -Milliseconds 100

        # Get the HWND of the host window (not the WebBrowser — window HWND captures its children too)
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($hostWin)
        $hwnd   = $helper.Handle

        # Capture via PrintWindow (works on HWND-hosted controls; RenderTargetBitmap does NOT)
        $screenDC = [SvgCapture]::GetDC([IntPtr]::Zero)
        $memDC    = [SvgCapture]::CreateCompatibleDC($screenDC)
        $hBmp     = [SvgCapture]::CreateCompatibleBitmap($screenDC, $Width, $Height)
        $oldObj   = [SvgCapture]::SelectObject($memDC, $hBmp)
        # PW_RENDERFULLCONTENT = 2 (forces re-render, needed for HwndHost controls)
        $null = [SvgCapture]::PrintWindow($hwnd, $memDC, 2)

        # Convert GDI HBITMAP -> WPF BitmapSource
        $bmpSrc = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHBitmap(
                      $hBmp, [IntPtr]::Zero,
                      [System.Windows.Int32Rect]::Empty,
                      [System.Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions())
        $null = $bmpSrc.Freeze()

        # GDI cleanup
        $null = [SvgCapture]::SelectObject($memDC, $oldObj)
        $null = [SvgCapture]::DeleteDC($memDC)
        $null = [SvgCapture]::DeleteObject($hBmp)
        $null = [SvgCapture]::ReleaseDC([IntPtr]::Zero, $screenDC)
        $null = $hostWin.Close()

        $null = Write-VerboseLog "SVG logo captured at ${Width}x${Height}." -Level Success
        return $bmpSrc
    } catch {
        $null = Write-VerboseLog "SVG render failed: $($_.Exception.Message)" -Level Warning
        return $null
    }
}

function Load-BitmapSource {
    param([string]$LocalPath, [string]$Url, [string]$Base64,
          [int]$SvgW = 120, [int]$SvgH = 40)

    # Robust SVG detector: handles BOM, <?xml?> preamble, leading whitespace
    function Test-IsSvg { param([string]$s)
        if (-not $s) { return $false }
        $sample = $s.Substring(0, [Math]::Min(600, $s.Length))
        return $sample -match '(?is)<svg[\s>]'
    }

    $result = $null

    # Priority 1: Base64-encoded image embedded in DevConfig
    if (-not [string]::IsNullOrWhiteSpace($Base64)) {
        try {
            $bytes   = [Convert]::FromBase64String($Base64.Trim())
            $decoded = [System.Text.Encoding]::UTF8.GetString($bytes, 0, [Math]::Min(600, $bytes.Length))
            if (Test-IsSvg $decoded) {
                $svgText = [System.Text.Encoding]::UTF8.GetString($bytes)
                $result  = Convert-SvgToBitmap -SvgContent $svgText -Width $SvgW -Height $SvgH
            } else {
                # Raster bytes (PNG / JPG / ICO / BMP)
                $stream = [System.IO.MemoryStream]::new($bytes)
                $bmp    = [System.Windows.Media.Imaging.BitmapImage]::new()
                $bmp.BeginInit()
                $bmp.StreamSource = $stream
                $bmp.CacheOption  = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                $bmp.EndInit()
                $null = $bmp.Freeze()
                $result = $bmp
            }
        } catch { Write-VerboseLog "Base64 logo decode failed: $($_.Exception.Message)" -Level Warning }
        # Ensure we return only a BitmapSource (discard any leaked pipeline objects)
        if ($result -isnot [System.Windows.Media.Imaging.BitmapSource]) { $result = $null }
        return $result
    }

    # Priority 2: Local file path
    if (-not [string]::IsNullOrWhiteSpace($LocalPath) -and (Test-Path $LocalPath)) {
        if ($LocalPath -match '\.svg$') {
            $result = Convert-SvgToBitmap -SvgContent ([IO.File]::ReadAllText($LocalPath)) -Width $SvgW -Height $SvgH
        } else {
            $result = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]::new((Resolve-Path $LocalPath).Path))
        }
        if ($result -isnot [System.Windows.Media.Imaging.BitmapSource]) { $result = $null }
        return $result
    }

    # Priority 3: Web URL
    if (-not [string]::IsNullOrWhiteSpace($Url)) {
        if ($Url -match '\.svg(\?|$)') {
            try {
                $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -ErrorAction Stop
                $body = if ($resp.Content -is [byte[]]) {
                            [System.Text.Encoding]::UTF8.GetString($resp.Content)
                        } else { [string]$resp.Content }
                $result = Convert-SvgToBitmap -SvgContent $body -Width $SvgW -Height $SvgH
            } catch { Write-VerboseLog "SVG URL download failed: $($_.Exception.Message)" -Level Warning }
            if ($result -isnot [System.Windows.Media.Imaging.BitmapSource]) { $result = $null }
            return $result
        }
        try {
            $bmp = [System.Windows.Media.Imaging.BitmapImage]::new()
            $bmp.BeginInit()
            $bmp.UriSource     = [Uri]::new($Url)
            $bmp.CacheOption   = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $bmp.CreateOptions = [System.Windows.Media.Imaging.BitmapCreateOptions]::None
            $bmp.EndInit()
            $result = $bmp
        } catch { Write-VerboseLog "URL logo load failed: $($_.Exception.Message)" -Level Warning }
    }
    return $result
}

# Apply background image (local path or URL)
try {
    $bgSrc = Load-BitmapSource -LocalPath $DevConfig.BackgroundImagePath -Url $DevConfig.BackgroundImageUrl
    if ($bgSrc) { $script:BgImage.Source = $bgSrc }
} catch { Write-VerboseLog "Could not load background image: $($_.Exception.Message)" -Level Warning }

# Apply logo (local path or URL)
try {
    $logoSrc = Load-BitmapSource -LocalPath $DevConfig.LogoImagePath -Url $DevConfig.LogoImageUrl `
                                 -Base64 $DevConfig.LogoBase64 -SvgW $DevConfig.LogoWidth -SvgH $DevConfig.LogoHeight
    if ($logoSrc) {
        $script:LogoImg.Source     = $logoSrc
        $script:LogoImg.Visibility = 'Visible'
    }
} catch { Write-VerboseLog "Could not load logo: $($_.Exception.Message)" -Level Warning }

# Log pane toggle
$script:Window.FindName('BtnLogToggle').Add_Click({
    if ($script:LogPaneVisible) {
        $script:LogPane.Height    = 0
        $script:LogPane.Visibility = 'Collapsed'
        $script:Window.FindName('BtnLogToggle').Content = '^ Show'
        $script:LogPaneVisible = $false
    } else {
        $script:LogPane.Visibility = 'Visible'
        $script:LogPane.Height     = $DevConfig.VerbosePaneDefaultHeight
        $script:Window.FindName('BtnLogToggle').Content = 'v Hide'
        $script:LogPaneVisible = $true
    }
})

$script:Window.FindName('BtnLogClear').Add_Click({
    $script:RtbLog.Document.Blocks.Clear()
})

# ÄÄ Collapse Group Mgmt blade on startup ÄÄ
$script:Window.FindName('BladeGroupMgmtContent').Visibility = 'Collapsed'
$script:Window.FindName('BladeDivider').Visibility = 'Collapsed'
$script:Window.FindName('TxtBladeArrow').Text = '+'

# ── Group Management blade toggle ──
$script:BladeExpanded = $false
$script:Window.FindName('BtnBladeGroupMgmt').Add_Click({
    $script:BladeExpanded = -not $script:BladeExpanded
    $content = $script:Window.FindName('BladeGroupMgmtContent')
    $divider = $script:Window.FindName('BladeDivider')
    $arrow   = $script:Window.FindName('TxtBladeArrow')
    if ($script:BladeExpanded) {
        $content.Visibility = 'Visible'
        $divider.Visibility = 'Visible'
        $arrow.Text = [char]0x2212
    } else {
        $content.Visibility = 'Collapsed'
        $divider.Visibility = 'Collapsed'
        $arrow.Text = '+'
    }
})

# ── Clear Inputs button ──
# Show when any panel other than Welcome is active; clears all input fields in the current panel
$script:Window.FindName('BtnClearInputs').Add_Click({
    $panels = @{
        'PanelCreate'     = @('TxtCreateName','TxtCreateMailNick','TxtCreateDesc','TxtCreateOwner','TxtCreateMembers','TxtCreateDynamic')
        'PanelMembership' = @('TxtGroupSearch','TxtMemberList')
        'PanelExport'     = @('TxtExportGroupSearch','TxtExportPath')
        'PanelDynamic'    = @('TxtDynGroupSearch','TxtDynRule')
        'PanelRename'     = @('TxtRenGroupSearch','TxtNewGroupName')
        'PanelOwner'      = @('TxtOwnerGroupList','TxtOwnerList')
        'PanelUserDevices'= @('TxtUDGroupSearch','TxtUDUpnList')
        'PanelFindCommon' = @('TxtFCInputList')
        'PanelFindDistinct' = @('TxtFDInputList')
        'PanelGetDeviceInfo' = @('TxtGDIInputList')
        'PanelGetDiscoveredApps' = @('TxtDAInputList','TxtDAKeyword')
        'PanelSearch'        = @('TxtSearchKeyword')
        'PanelListMembers'   = @('TxtLMGroupInput')
        'PanelObjectMembership' = @('TxtOMInputList')
        'PanelFindGroupsByOwners' = @('TxtFGOInputList')
        'PanelCompareGroups' = @('TxtCGGroupInput')
        'PanelGetPolicyAssignments' = @('TxtGPANameFilter')
        'PanelRemoveGroups'          = @('TxtRGGroupList')
    }
    $badgesToHide = @(
        'SelectedGroupBadge','GroupSearchResults',
        'ExportSelectedGroupBadge','ExportGroupSearchResults',
        'DynSelectedGroupBadge','DynGroupSearchResults',
        'RenSelectedGroupBadge','RenGroupSearchResults',
        'UDSelectedGroupBadge','UDGroupSearchResults'
    )
    foreach ($panelName in $panels.Keys) {
        $panel = $script:Window.FindName($panelName)
        if ($panel -and $panel.Visibility -eq 'Visible') {
            foreach ($fieldName in $panels[$panelName]) {
                $field = $script:Window.FindName($fieldName)
                if ($field) { $field.Text = '' }
            }
            # Reset combos (Create panel)
            $cmb = $script:Window.FindName('CmbCreateType')
            if ($cmb) { $cmb.SelectedIndex = 0 }
            $cmbLbl = $script:Window.FindName('CmbSensitivityLabel')
            if ($cmbLbl) { $cmbLbl.SelectedIndex = 0 }
            # Hide group badges and result lists
            foreach ($b in $badgesToHide) {
                $el = $script:Window.FindName($b)
                if ($el) { $el.Visibility = 'Collapsed' }
            }
            # Reset selected group state for this panel
            $script:SelectedGroup       = $null
            $script:ExportSelectedGroup = $null
            $script:DynSelectedGroup     = $null
            $script:RenSelectedGroup     = $null
            $script:OwnerGroupsTxt       = ''
            $script:OwnerGroupsValidated = $null
            $script:UDSelectedGroup      = $null
            # Special reset for Search panel (clears results + buttons + checkboxes)
            if ($panelName -eq 'PanelSearch') {
                $dg = $script:Window.FindName('DgSearchResults')
                if ($dg) { $dg.ItemsSource = $null }
                $script:Window.FindName('PnlSearchResults').Visibility   = 'Collapsed'
                $script:Window.FindName('TxtSearchNoResults').Visibility = 'Collapsed'
                $script:Window.FindName('BtnSearchCopyValue').IsEnabled  = $false
                $script:Window.FindName('BtnSearchCopyRow').IsEnabled    = $false
                $script:Window.FindName('BtnSearchCopyAll').IsEnabled    = $false
                $script:Window.FindName('TxtSearchFilter').IsEnabled  = $false
                $script:Window.FindName('BtnSearchFilter').IsEnabled  = $false
                $script:Window.FindName('BtnSearchFilterClear').IsEnabled = $false
                $script:Window.FindName('TxtSearchFilter').Text       = ''
                $script:SearchAllData = $null
                $script:Window.FindName('BtnSearchExportXlsx').IsEnabled = $false
                foreach ($cb in @('ChkSrchUsers','ChkSrchGetManager','ChkSrchDevices','ChkSrchSG','ChkSrchM365')) {
                    $cbx = $script:Window.FindName($cb)
                    if ($cbx) { $cbx.IsChecked = $true }
                }
                $script:SearchResults = $null
            }
            # Special reset for List Group Members panel
            if ($panelName -eq 'PanelListMembers') {
                $dg = $script:Window.FindName('DgLMResults')
                if ($dg) { $dg.ItemsSource = $null; $dg.Columns.Clear() }
                $script:Window.FindName('PnlLMResults').Visibility   = 'Collapsed'
                $script:Window.FindName('TxtLMNoResults').Visibility = 'Collapsed'
                $script:Window.FindName('PnlLMProgress').Visibility  = 'Collapsed'
                $script:Window.FindName('BtnLMCopyValue').IsEnabled  = $false
                $script:Window.FindName('BtnLMCopyRow').IsEnabled    = $false
                $script:Window.FindName('BtnLMCopyAll').IsEnabled    = $false
                $script:Window.FindName('TxtLMFilter').IsEnabled  = $false
                $script:Window.FindName('BtnLMFilter').IsEnabled  = $false
                $script:Window.FindName('BtnLMFilterClear').IsEnabled = $false
                $script:Window.FindName('TxtLMFilter').Text       = ''
                $script:LMAllData = $null
                $script:Window.FindName('BtnLMExportXlsx').IsEnabled = $false
                foreach ($cb in @('ChkLMUsers','ChkLMDevices','ChkLMGroups')) {
                    $cbx = $script:Window.FindName($cb)
                    if ($cbx) { $cbx.IsChecked = $true }
                }
                $script:LMAllResults = $null
                if ($script:LMProgressTimer) { $script:LMProgressTimer.Stop() }
            }
            # Special reset for Object Membership panel
            if ($panelName -eq 'PanelObjectMembership') {
                $dg = $script:Window.FindName('DgOMResults')
                if ($dg) { $dg.ItemsSource = $null }
                $script:Window.FindName('PnlOMResults').Visibility   = 'Collapsed'
                $script:Window.FindName('TxtOMNoResults').Visibility = 'Collapsed'
                $script:Window.FindName('PnlOMProgress').Visibility  = 'Collapsed'
                $script:Window.FindName('BtnOMCopyValue').IsEnabled  = $false
                $script:Window.FindName('BtnOMCopyRow').IsEnabled    = $false
                $script:Window.FindName('BtnOMCopyAll').IsEnabled    = $false
                $script:Window.FindName('BtnOMFilter').IsEnabled  = $false
                $script:Window.FindName('BtnOMFilterClear').IsEnabled = $false
                $script:OMAllData = $null
                $script:Window.FindName('TxtOMFilter').Text       = ''
                $script:Window.FindName('TxtOMFilter').IsEnabled  = $false
                $script:Window.FindName('BtnOMExportXlsx').IsEnabled = $false
                $script:OMAllResults = $null
                if ($script:OmTimer) { $script:OmTimer.Stop() }
            }
            # Special reset for Find Groups by Owners panel
            if ($panelName -eq 'PanelFindGroupsByOwners') {
                $dg = $script:Window.FindName('DgFGOResults')
                if ($dg) { $dg.ItemsSource = $null }
                $script:Window.FindName('PnlFGOResults').Visibility   = 'Collapsed'
                $script:Window.FindName('TxtFGONoResults').Visibility = 'Collapsed'
                $script:Window.FindName('PnlFGOProgress').Visibility  = 'Collapsed'
                $script:Window.FindName('BtnFGOCopyValue').IsEnabled  = $false
                $script:Window.FindName('BtnFGOCopyRow').IsEnabled    = $false
                $script:Window.FindName('BtnFGOCopyAll').IsEnabled    = $false
                $script:Window.FindName('BtnFGOFilter').IsEnabled  = $false
                $script:Window.FindName('BtnFGOFilterClear').IsEnabled = $false
                $script:FGOAllData = $null
                $script:Window.FindName('TxtFGOFilter').Text       = ''
                $script:Window.FindName('TxtFGOFilter').IsEnabled  = $false
                $script:Window.FindName('BtnFGOExportXlsx').IsEnabled = $false
                if ($script:FGOTimer) { $script:FGOTimer.Stop() }
            }
            # Special reset for Get App Info panel
            if ($panelName -eq 'PanelGetDiscoveredApps') {
                $dg = $script:Window.FindName('DgDAResults')
                if ($dg) { $dg.ItemsSource = $null; $dg.Columns.Clear() }
                $script:Window.FindName('PnlDAResults').Visibility   = 'Collapsed'
                $script:Window.FindName('TxtDANoResults').Visibility = 'Collapsed'
                $script:Window.FindName('PnlDAProgress').Visibility  = 'Collapsed'
                $script:Window.FindName('BtnDACopyValue').IsEnabled  = $false
                $script:Window.FindName('BtnDACopyRow').IsEnabled    = $false
                $script:Window.FindName('BtnDACopyAll').IsEnabled    = $false
                $script:Window.FindName('TxtDAFilter').IsEnabled  = $false
                $script:Window.FindName('BtnDAFilter').IsEnabled  = $false
                $script:Window.FindName('BtnDAFilterClear').IsEnabled = $false
                $script:Window.FindName('TxtDAFilter').Text       = ''
                $script:DAAllData = $null
                $script:Window.FindName('BtnDAExportXlsx').IsEnabled = $false
                if ($script:DaTimer) { $script:DaTimer.Stop() }
                foreach ($chk in @('ChkDAWindows','ChkDAAndroid','ChkDAiOS','ChkDAMacOS')) {
                    $cbx = $script:Window.FindName($chk)
                    if ($cbx) { $cbx.IsChecked = $false }
                }
                $script:Window.FindName('ChkDADiscoveredApps').IsChecked = $true
                $script:Window.FindName('ChkDAManagedApps').IsChecked    = $true
                $script:Window.FindName('BtnDARun').IsEnabled             = $true
                $script:Window.FindName('BtnDARunAll').IsEnabled          = $true
                $script:Window.FindName('BtnDAGetManagedAssignments').IsEnabled = $true
            }
            # Special reset for Compare Groups panel
            if ($panelName -eq 'PanelCompareGroups') {
                $dg = $script:Window.FindName('DgCGResults')
                if ($dg) { $dg.ItemsSource = $null }
                $script:Window.FindName('PnlCGResults').Visibility   = 'Collapsed'
                $script:Window.FindName('TxtCGNoResults').Visibility = 'Collapsed'
                $script:Window.FindName('PnlCGProgress').Visibility  = 'Collapsed'
                $script:Window.FindName('BtnCGCopyValue').IsEnabled  = $false
                $script:Window.FindName('BtnCGCopyRow').IsEnabled    = $false
                $script:Window.FindName('BtnCGCopyAll').IsEnabled    = $false
                $script:Window.FindName('TxtCGFilter').IsEnabled  = $false
                $script:Window.FindName('BtnCGFilter').IsEnabled  = $false
                $script:Window.FindName('BtnCGFilterClear').IsEnabled = $false
                $script:Window.FindName('TxtCGFilter').Text       = ''
                $script:CGAllData = $null
                $script:Window.FindName('BtnCGExportXlsx').IsEnabled = $false
                $script:CGAllResults = $null
                if ($script:CGProgressTimer) { $script:CGProgressTimer.Stop() }
                $script:Window.FindName('BtnCGCommon').IsEnabled   = $true
                $script:Window.FindName('BtnCGDistinct').IsEnabled = $true
            }
            # Special reset for Find Common Groups panel
            if ($panelName -eq 'PanelFindCommon') {
                $dg = $script:Window.FindName('DgFCResults')
                if ($dg) { $dg.ItemsSource = $null }
                $script:Window.FindName('PnlFCResults').Visibility   = 'Collapsed'
                $script:Window.FindName('TxtFCNoResults').Visibility = 'Collapsed'
                $script:Window.FindName('PnlFCProgress').Visibility  = 'Collapsed'
                $script:Window.FindName('BtnFCCopyValue').IsEnabled  = $false
                $script:Window.FindName('BtnFCCopyRow').IsEnabled    = $false
                $script:Window.FindName('BtnFCCopyAll').IsEnabled    = $false
                $script:Window.FindName('TxtFCFilter').Text       = ''
                $script:FCAllData = $null
                $script:Window.FindName('BtnFCFilter').IsEnabled  = $false
                $script:Window.FindName('BtnFCFilterClear').IsEnabled = $false
                $script:Window.FindName('TxtFCFilter').IsEnabled  = $false
                $script:Window.FindName('BtnFCExportXlsx').IsEnabled = $false
                $script:FCResult = $null
            }
            # Special reset for Find Distinct Groups panel
            if ($panelName -eq 'PanelFindDistinct') {
                $dg = $script:Window.FindName('DgFDResults')
                if ($dg) { $dg.ItemsSource = $null }
                $script:Window.FindName('PnlFDResults').Visibility   = 'Collapsed'
                $script:Window.FindName('TxtFDNoResults').Visibility = 'Collapsed'
                $script:Window.FindName('PnlFDProgress').Visibility  = 'Collapsed'
                $script:Window.FindName('BtnFDCopyValue').IsEnabled  = $false
                $script:Window.FindName('BtnFDCopyRow').IsEnabled    = $false
                $script:Window.FindName('BtnFDCopyAll').IsEnabled    = $false
                $script:Window.FindName('TxtFDFilter').IsEnabled  = $false
                $script:Window.FindName('BtnFDFilter').IsEnabled  = $false
                $script:Window.FindName('BtnFDFilterClear').IsEnabled = $false
                $script:Window.FindName('TxtFDFilter').Text       = ''
                $script:FDAllData = $null
                $script:Window.FindName('BtnFDExportXlsx').IsEnabled = $false
                $script:FDResult = $null
            }
            # Special reset for Get Policy Info panel
            if ($panelName -eq 'PanelGetPolicyAssignments') {
                $script:Window.FindName('CmbGPAType').SelectedIndex     = 0
                $script:Window.FindName('CmbGPAType').IsEnabled         = $true
                $script:Window.FindName('CmbGPASubType').SelectedIndex  = 0
                $script:Window.FindName('CmbGPAPlatform').SelectedIndex = 0
                $script:Window.FindName('PnlGPAProgress').Visibility    = 'Collapsed'
                $script:Window.FindName('PnlGPAResults').Visibility     = 'Collapsed'
                $script:Window.FindName('TxtGPANoResults').Visibility   = 'Collapsed'
                $dg = $script:Window.FindName('DgGPAResults')
                if ($dg) { $dg.ItemsSource = $null }
                foreach ($b in @('BtnGPACopyValue','BtnGPACopyRow','BtnGPACopyAll','BtnGPAExportXlsx')) {
                    $script:Window.FindName($b).IsEnabled = $false
                }
                foreach ($b in @('BtnGPAGetSelected','BtnGPAGetAll')) {
                    $script:Window.FindName($b).IsEnabled = $true
                }
                $script:GPAResult = $null
            }
            # Special reset for Remove Groups panel
            if ($panelName -eq 'PanelRemoveGroups') {
                $dg = $script:Window.FindName('DgRGPreview')
                if ($dg) { $dg.ItemsSource = $null }
                $dg2 = $script:Window.FindName('DgRGResults')
                if ($dg2) { $dg2.ItemsSource = $null }
                $script:Window.FindName('PnlRGProgress').Visibility      = 'Collapsed'
                $script:Window.FindName('PnlRGWarning').Visibility        = 'Collapsed'
                $script:Window.FindName('PnlRGPreview').Visibility        = 'Collapsed'
                $script:Window.FindName('PnlRGExecProgress').Visibility   = 'Collapsed'
                $script:Window.FindName('PnlRGResults').Visibility        = 'Collapsed'
                $script:Window.FindName('TxtRGNoResults').Visibility      = 'Collapsed'
                $script:Window.FindName('BtnRGExecute').IsEnabled         = $false
                $script:Window.FindName('BtnRGCopyAll').IsEnabled         = $false
                $script:Window.FindName('BtnRGExportXlsx').IsEnabled      = $false
                $script:RGGroupList  = $null
                $script:RGValidated  = $null
                $script:RGExecResult = $null
            }
            Hide-Notification
            Write-VerboseLog "Inputs cleared." -Level Info
            break
        }
    }
})

# ── Reset-AuthUI  -  defined at script scope so the DispatcherTimer tick can call it ──
function Reset-AuthUI {
    param([string]$StatusText = 'Not connected')
    try { $script:AuthTimer.Stop() }                              catch {}
    try { $script:AuthPs.Stop(1000, $null) | Out-Null }          catch {}
    try { $script:AuthPs.EndInvoke($script:AuthHandle) }         catch {}
    try { $script:AuthRs.Close() }                               catch {}
    try { $script:AuthPs.Dispose() }                             catch {}
    $script:Window.Tag = $null
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    $script:Window.FindName('BtnConnect').IsEnabled     = $true
    $script:Window.FindName('BtnAuthCancel').Visibility = 'Collapsed'
    Update-StatusBar -Text $StatusText
}

# ── CONNECT button ──
$script:Window.FindName('BtnConnect').Add_Click({
    Hide-Notification

    # ── Auth input dialog  -  pre-populated from DevConfig, user can override Client ID ──
    [xml]$authDlgXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Connect to Microsoft Graph" Width="460" SizeToContent="Height"
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize"
        Background="$($DevConfig.ColorBackground)"
        FontFamily="$($DevConfig.FontFamily)" FontSize="13">
  <Grid Margin="20">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="10"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="10"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="16"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <StackPanel Grid.Row="0">
      <TextBlock Text="TENANT ID" FontSize="11" FontWeight="SemiBold"
                 Foreground="$($DevConfig.ColorTextMuted)" Margin="0,0,0,3"/>
      <TextBox x:Name="TxtAuthTenantId" FontSize="12" Padding="8,6"
               Background="$($DevConfig.ColorSurface)" Foreground="$($DevConfig.ColorText)"
               BorderBrush="$($DevConfig.ColorBorder)" BorderThickness="1"/>
    </StackPanel>
    <StackPanel Grid.Row="2">
      <TextBlock Text="CLIENT ID  (override or leave as configured)" FontSize="11" FontWeight="SemiBold"
                 Foreground="$($DevConfig.ColorTextMuted)" Margin="0,0,0,3"/>
      <TextBox x:Name="TxtAuthClientId" FontSize="12" Padding="8,6"
               Background="$($DevConfig.ColorSurface)" Foreground="$($DevConfig.ColorText)"
               BorderBrush="$($DevConfig.ColorBorder)" BorderThickness="1"/>
    </StackPanel>
    <StackPanel Grid.Row="4">
      <TextBlock x:Name="TxtAuthError" FontSize="11" Foreground="$($DevConfig.ColorError)"
                 Visibility="Collapsed" TextWrapping="Wrap"/>
    </StackPanel>
    <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="BtnAuthCancel" Content="Cancel" Width="80" Margin="0,0,10,0"
              Background="Transparent" Foreground="$($DevConfig.ColorText)"
              BorderBrush="$($DevConfig.ColorBorder)" BorderThickness="1"
              Padding="8,6" Cursor="Hand"/>
      <Button x:Name="BtnAuthConnect" Content="Connect" Width="100"
              Background="$($DevConfig.ColorPrimary)" Foreground="White"
              BorderThickness="0" Padding="8,6" FontWeight="SemiBold" Cursor="Hand"/>
    </StackPanel>
  </Grid>
</Window>
"@
    $authReader = [System.Xml.XmlNodeReader]::new($authDlgXaml)
    $authDlg    = [Windows.Markup.XamlReader]::Load($authReader)
    $authDlg.Owner = $script:Window

    # Pre-populate fields
    $authDlg.FindName('TxtAuthTenantId').Text = $DevConfig.TenantId
    $authDlg.FindName('TxtAuthClientId').Text = $DevConfig.ClientId

    $script:AuthDlgResult  = $false
    $script:AuthDlgTenant  = ''
    $script:AuthDlgClient  = ''

    $authDlg.FindName('BtnAuthCancel').Add_Click({
        $script:AuthDlgResult = $false
        $authDlg.Close()
    })

    $authDlg.FindName('BtnAuthConnect').Add_Click({
        $tid = $authDlg.FindName('TxtAuthTenantId').Text.Trim()
        $cid = $authDlg.FindName('TxtAuthClientId').Text.Trim()
        $errBlk = $authDlg.FindName('TxtAuthError')
        if ([string]::IsNullOrWhiteSpace($tid)) {
            $errBlk.Text = "Tenant ID is required."
            $errBlk.Visibility = 'Visible'
            return
        }
        if ([string]::IsNullOrWhiteSpace($cid)) {
            $errBlk.Text = "Client ID is required."
            $errBlk.Visibility = 'Visible'
            return
        }
        $script:AuthDlgResult = $true
        $script:AuthDlgTenant = $tid
        $script:AuthDlgClient = $cid
        $authDlg.Close()
    })

    $authDlg.ShowDialog() | Out-Null

    if (-not $script:AuthDlgResult) {
        Write-VerboseLog "Connection cancelled by user." -Level Warning
        return
    }

    $effectiveTenantId = $script:AuthDlgTenant
    $effectiveClientId = $script:AuthDlgClient

    # Auth timeout  -  seconds to wait for browser authentication before aborting
    $script:AuthTimeoutSec = 120

    $script:Window.FindName('BtnConnect').IsEnabled     = $false
    $script:Window.FindName('BtnAuthCancel').Visibility = 'Visible'
    $script:Window.FindName('BtnAuthCancel').IsEnabled  = $true
    $script:Window.FindName('BtnAuthCancel').Content = (New-StopBtnContent 'Cancel Auth')
    $script:AuthCancelled = $false
    $script:AuthStartTime = [datetime]::UtcNow
    Write-VerboseLog 'Initiating Microsoft Graph connection...' -Level Action
    Write-VerboseLog "Tenant ID : $effectiveTenantId" -Level Info
    Write-VerboseLog "Client ID : $effectiveClientId" -Level Info
    Write-VerboseLog "Browser authentication window should open. Sign in within $script:AuthTimeoutSec seconds, or click [Stop] Cancel Auth to abort." -Level Info
    Update-StatusBar -Text 'Waiting for browser authentication... (Cancel Auth to abort)'

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('TenantId', $effectiveTenantId)
    $rs.SessionStateProxy.SetVariable('ClientId', $effectiveClientId)
    $rs.SessionStateProxy.SetVariable('Window',   $script:Window)

    $ps = [powershell]::Create()
    $ps.Runspace = $rs

    $null = $ps.AddScript({
        param()

        # ── Helper: clear MSAL token cache to remove stale/corrupt entries ──
        function Clear-MsalCache {
            try {
                $cachePath = Join-Path $env:LOCALAPPDATA 'Microsoft\TokenCache'
                if (Test-Path $cachePath) {
                    Get-ChildItem $cachePath -File -ErrorAction SilentlyContinue |
                        Remove-Item -Force -ErrorAction SilentlyContinue
                }
                # Also clear the Graph SDK's own cache
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            } catch {}
        }

        # ── Helper: try browser auth, return error message or $null on success ──
        function Try-BrowserAuth {
            try {
                try {
                    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -NoWelcome `
                        -AuthType Browser -ErrorAction Stop
                } catch [System.Management.Automation.ParameterBindingException] {
                    # -AuthType not supported in this SDK version
                    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -NoWelcome -ErrorAction Stop
                }
                return $null   # success
            } catch {
                return $_.Exception.Message
            }
        }

        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

            # ── Attempt 1: browser auth ────────────────────────────────────
             # Force browser re-authentication - clear any cached MSAL tokens first
             Clear-MsalCache
            $err1 = Try-BrowserAuth

            if ($err1) {
                # Signal UI to log the first failure
                $Window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Normal, [Action]{
                    $Window.Tag = "warn:Browser auth failed: $err1  -  clearing cache and retrying..."
                })
                Start-Sleep -Milliseconds 500

                # ── Attempt 2: clear MSAL cache + retry browser ────────────
                Clear-MsalCache
                $err2 = Try-BrowserAuth

                if ($err2) {
                    # Signal UI to show device code instructions
                    $Window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Normal, [Action]{
                        $Window.Tag = "warn:Browser auth failed again: $err2  -  switching to device code flow..."
                    })
                    Start-Sleep -Milliseconds 500

                    # ── Attempt 3: TLS callback fix + retry browser ────────
                    # Zscaler and other SSL-inspection proxies re-sign TLS
                    # connections. The runspace may not inherit the Windows
                    # certificate store correctly, causing MSAL to reject the
                    # Zscaler root CA and produce an invalid JWT error.
                    # Setting ServerCertificateValidationCallback to null forces
                    # .NET to use the OS certificate store (which includes the
                    # Zscaler root CA pushed by IT policy). Scoped to this
                    # runspace only  -  disposed immediately after auth.
                    [Net.ServicePointManager]::SecurityProtocol = `
                        [Net.SecurityProtocolType]::Tls12
                    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
                    Clear-MsalCache
                    $err3 = Try-BrowserAuth
                    if ($err3) { throw $err3 }
                }
            }

            $ctx = Get-MgContext
            $Window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Normal, [Action]{
                $Window.Tag = "connected:$($ctx.Account)|$($ctx.TenantId)|$(($ctx.Scopes -join ','))"
            })
        } catch {
            $err = $_.Exception.Message
            $Window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Normal, [Action]{
                $Window.Tag = "error:$err"
            })
        }
    })

    $handle = $ps.BeginInvoke()

    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(500)
    $script:AuthTimer  = $timer
    $script:AuthPs     = $ps
    $script:AuthRs     = $rs
    $script:AuthHandle = $handle

    $timer.Add_Tick({
        # ── User clicked Cancel Auth ────────────────────────────────
        if ($script:AuthCancelled) {
            Write-VerboseLog 'Authentication cancelled by user.' -Level Warning
            Reset-AuthUI -StatusText 'Not connected'
            Show-Notification 'Authentication cancelled.' -BgColor '#FFF3CD' -FgColor '#7A4800'
            return
        }

        # ── Timeout ─────────────────────────────────────────────────
        $elapsed = ([datetime]::UtcNow - $script:AuthStartTime).TotalSeconds
        if ($elapsed -gt $script:AuthTimeoutSec) {
            Write-VerboseLog "Authentication timed out after $script:AuthTimeoutSec seconds." -Level Warning
            Reset-AuthUI -StatusText 'Not connected  -  timed out'
            Show-Notification "Authentication timed out after $script:AuthTimeoutSec seconds. Please try again." `
                -BgColor '#FFF3CD' -FgColor '#7A4800'
            return
        }

        $tag = $script:Window.Tag -as [string]
        if (-not $tag) { return }

        # ── Intermediate warning from runspace (cache clear / fallback notice) ──
        if ($tag.StartsWith('warn:')) {
            $warnMsg = $tag.Substring(5)
            $script:Window.Tag = $null
            Write-VerboseLog $warnMsg -Level Warning
            # If switching to device code, update status bar to guide the user
            if ($warnMsg -like '*TLS*' -or $warnMsg -like '*switching*') {
                Update-StatusBar -Text 'Retrying with TLS fix  -  please complete browser authentication again'
            }
            return
        }

        if ($tag.StartsWith('connected:')) {
            $script:AuthTimer.Stop()
            $info   = $tag.Substring(10) -split '\|'
            $acct   = $info[0]
            $tenant = $info[1]
            try { $script:AuthPs.EndInvoke($script:AuthHandle) } catch {}
            $script:AuthRs.Close()
            $script:AuthPs.Dispose()
            $script:Window.Tag = $null
            $script:Window.FindName('BtnAuthCancel').Visibility = 'Collapsed'

            Write-VerboseLog 'Connected successfully' -Level Success
            Write-VerboseLog "Account : $acct" -Level Info
            Write-VerboseLog "Tenant  : $tenant" -Level Info

            Update-StatusBar -Text "Connected as: $acct | Tenant: $tenant"
            $script:Window.FindName('BtnConnect').IsEnabled    = $false
            $script:Window.FindName('BtnDisconnect').IsEnabled = $true
            Show-Notification -Message "Connected as $acct" -BgColor '#D4EDDA' -FgColor '#155724'

            foreach ($b in @('BtnOpSearch','BtnOpListMembers','BtnOpObjectMembership','BtnOpFindGroupsByOwners','BtnOpCreate','BtnOpAdd','BtnOpRemove','BtnOpExport','BtnOpDynamic','BtnOpRename','BtnOpOwner','BtnOpGetPolicyAssignments','BtnOpUserDevices','BtnOpFindCommon','BtnOpFindDistinct','BtnOpGetDeviceInfo','BtnOpGetDiscoveredApps','BtnOpCompareGroups','BtnOpRemoveGroups')) {
                $script:Window.FindName($b).IsEnabled = $true
            }
            Invoke-PimCheck
            Start-PimAutoRefresh

        } elseif ($tag.StartsWith('error:')) {
            $err = $tag.Substring(6)
            Write-VerboseLog "Connection failed: $err" -Level Error
            Reset-AuthUI -StatusText 'Not connected'
            Show-Notification "Connection failed: $err" -BgColor '#F8D7DA' -FgColor '#721C24'
        }
    })
    $timer.Start()
})

# ── AUTH CANCEL button ──
$script:Window.FindName('BtnAuthCancel').Add_Click({
    $script:AuthCancelled = $true
    $script:Window.FindName('BtnAuthCancel').IsEnabled = $false
    $script:Window.FindName('BtnAuthCancel').Content   = 'Cancelling...'
    Write-VerboseLog 'Cancelling authentication...' -Level Warning
})

# ── DISCONNECT button ──
$script:Window.FindName('BtnDisconnect').Add_Click({
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Write-VerboseLog "Disconnected from Microsoft Graph." -Level Info
    } catch {}
    Update-StatusBar -Text "Not connected"
    $script:Window.FindName('BtnConnect').IsEnabled    = $true
    $script:Window.FindName('BtnDisconnect').IsEnabled = $false
    $script:PimStrip.Visibility = 'Collapsed'
    $script:SpPimRoles.Children.Clear()
    Stop-PimAutoRefresh
    Show-Panel 'PanelWelcome'
    Hide-Notification
    foreach ($b in @('BtnOpSearch','BtnOpListMembers','BtnOpObjectMembership','BtnOpFindGroupsByOwners','BtnOpCreate','BtnOpAdd','BtnOpRemove','BtnOpExport','BtnOpDynamic','BtnOpRename','BtnOpOwner','BtnOpGetPolicyAssignments','BtnOpUserDevices','BtnOpFindCommon','BtnOpFindDistinct','BtnOpGetDeviceInfo','BtnOpGetDiscoveredApps','BtnOpCompareGroups','BtnOpRemoveGroups')) {
        $btn = $script:Window.FindName($b)
        if ($btn) { $btn.IsEnabled = $false }
    }
})

# ÄÄ FEEDBACK button ÄÄ
# Hide the button if no FeedbackUrl is configured
$btnFeedback = $script:Window.FindName('BtnFeedback')
if ($btnFeedback) {
    # Always show the button; click handler opens URL or logs a hint if not configured
    $btnFeedback.Visibility = 'Visible'
    $btnFeedback.Add_Click({
        $url = $DevConfig.FeedbackUrl
        if (-not [string]::IsNullOrWhiteSpace($url)) {
            try { Start-Process $url } catch {
                Write-VerboseLog "Could not open feedback URL: $url" -Level Warning
            }
        } else {
            Write-VerboseLog "Feedback URL not configured. Set FeedbackUrl in DevConfig." -Level Info
        }
    })
}

# ── AUTHOR attribution link (blog version) ──
$authorLink = $script:Window.FindName('TxtAuthorLink')
if ($authorLink) {
    $authorLink.Add_MouseLeftButtonDown({
        try { Start-Process "https://www.linkedin.com/in/satish-singhi-791163167/" } catch {}
    })
}

# ── PIM Role Status Check ────────────────────────────────────────────────────
function Invoke-PimCheck {

    # Queries the current user's active PIM role assignments via Graph.
    # Runs on a background runspace, renders results as coloured badges
    # in the PIM strip. Colour: green = >60min, amber = <=60min, red = <=15min.

    Write-VerboseLog "Checking active PIM roles..." -Level Info
    $script:PimStrip.Visibility = 'Visible'
    $script:SpPimRoles.Children.Clear()

    # Placeholder while loading
    $loadTb = [System.Windows.Controls.TextBlock]::new()
    $loadTb.Text = "Checking..."; $loadTb.FontSize = 11; $loadTb.VerticalAlignment = 'Center'
    $loadTb.Foreground = [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.ColorConverter]::ConvertFromString('#A0C0E8'))
    $loadTb.Margin = [System.Windows.Thickness]::new(0,0,0,0)
    $script:SpPimRoles.Children.Add($loadTb) | Out-Null

    $script:PimQueue = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()

    $script:PimRs = [runspacefactory]::CreateRunspace()
    $script:PimRs.ApartmentState = 'MTA'
    $script:PimRs.ThreadOptions  = 'ReuseThread'
    $script:PimRs.Open()
    $script:PimRs.SessionStateProxy.SetVariable('PimQueue', $script:PimQueue)

    $script:PimPs = [powershell]::Create()
    $script:PimPs.Runspace = $script:PimRs
    $null = $script:PimPs.AddScript({
        function PQ { param([hashtable]$M) $PimQueue.Enqueue($M) }

        function Send-RoleResult {
            param([string]$RoleName, [string]$AssignType)
            # Expiry detection removed  -  show active role name only.
            PQ @{ Type='role'; RoleName=$RoleName; AssignType=$AssignType }
        }

        Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
        $ctx = Get-MgContext
        if (-not $ctx) { PQ @{ Type='error'; Msg='No active Graph context' }; PQ @{ Type='done' }; return }

        $me     = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me?$select=id,userPrincipalName'
        $userId = $me['id']
        PQ @{ Type='log'; Msg="Checking PIM roles for: $($me['userPrincipalName'])" }

        # Track role names for deduplication between queries
        $builtInNames = @{}
        $customCount  = 0

        # ── transitiveMemberOf: active directory roles for the current user ──
        # Uses only Directory.Read.All (already required). Returns currently active roles.
        try {
            $uri2  = "https://graph.microsoft.com/v1.0/me/transitiveMemberOf/microsoft.graph.directoryRole" +
                     "?`$select=id,displayName,roleTemplateId"
            $resp2 = Invoke-MgGraphRequest -Method GET -Uri $uri2
            $roles = $resp2['value']
            PQ @{ Type='log'; Msg=('PIM: transitiveMemberOf (' + $roles.Count + ' active role(s))') }
        
            if (-not $roles -or $roles.Count -eq 0) {
                PQ @{ Type='log'; Msg='PIM: No built-in directory roles found via transitiveMemberOf' }
            } else {
                foreach ($role in $roles) {
                    $rName = $role['displayName']
                    $builtInNames[$rName] = $true
                    Send-RoleResult -RoleName $rName -AssignType 'Active'
                }
            }
        } catch {
            $pimErr = $_.Exception.Message
            if ($pimErr -match '403|Forbidden|401|Unauthorized') {
                PQ @{ Type='log'; Msg='PIM: transitiveMemberOf returned 403 - skipping built-in role check' }
            } else {
                PQ @{ Type='error'; Msg=$pimErr }
            }
        }

        # ── roleAssignments: custom + direct role assignments (includes custom roles) ──
        # Uses Directory.Read.All (already required). Catches custom Entra roles not visible
        # via transitiveMemberOf which only returns built-in directoryRole objects.
        try {
            $raUri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments" +
                     "?`$filter=principalId eq '$userId'" +
                     "&`$expand=roleDefinition(`$select=displayName,isBuiltIn)"
            $raResp = Invoke-MgGraphRequest -Method GET -Uri $raUri
            $assignments = $raResp['value']
            if ($assignments -and $assignments.Count -gt 0) {
                foreach ($ra in $assignments) {
                    $rd = $ra['roleDefinition']
                    if (-not $rd) { continue }
                    $rdName    = $rd['displayName']
                    $isBuiltIn = $rd['isBuiltIn']
                    # Skip built-in roles already reported by transitiveMemberOf
                    if ($isBuiltIn -and $builtInNames.ContainsKey($rdName)) { continue }
                    if (-not $isBuiltIn) {
                        $customCount++
                        Send-RoleResult -RoleName $rdName -AssignType 'Custom Role'
                    } else {
                        # Built-in role found via direct assignment but not transitiveMemberOf
                        $builtInNames[$rdName] = $true
                        Send-RoleResult -RoleName $rdName -AssignType 'Active'
                    }
                }
            }
            PQ @{ Type='log'; Msg="PIM: roleAssignments ($($assignments.Count) total, $customCount custom role(s))" }
        } catch {
            $raErr = $_.Exception.Message
            PQ @{ Type='log'; Msg="PIM: Could not query roleAssignments for custom roles: $raErr" }
        }

        # If no roles found from either query, report none
        if ($builtInNames.Count -eq 0 -and $customCount -eq 0) {
            PQ @{ Type='none' }
        }
        PQ @{ Type='done' }
    })

    $script:PimHandle = $script:PimPs.BeginInvoke()

    $script:PimTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:PimTimer.Interval = [TimeSpan]::FromMilliseconds(200)
    $script:PimTimer.Add_Tick({
        $m = $null
        $gotResults = $false
        while ($script:PimQueue.TryDequeue([ref]$m)) {
            switch ($m['Type']) {
                'log' {
                    # Verbose log message from runspace
                    Write-VerboseLog "PIM: $($m['Msg'])" -Level Info
                }
                'role' {
                    if (-not $gotResults) {
                        $script:SpPimRoles.Children.Clear()
                        $gotResults = $true
                    }
                    # Build badge  -  role name only, no expiry
                    $badge = [System.Windows.Controls.Border]::new()
                    $badge.CornerRadius = [System.Windows.CornerRadius]::new(4)
                    $badge.Padding      = [System.Windows.Thickness]::new(8,3,8,3)
                    $badge.Margin       = [System.Windows.Thickness]::new(0,3,6,3)
                    $isCustom = $m['AssignType'] -eq 'Custom Role'
                    $badgeBg = if ($isCustom) { '#3B3080' } else { '#1A5E3A' }
                    $badge.Background   = [System.Windows.Media.SolidColorBrush]::new(
                        [System.Windows.Media.ColorConverter]::ConvertFromString($badgeBg))
                    $badge.ToolTip      = "$($m['RoleName'])  ($($m['AssignType']))"

                    $badgeTb = [System.Windows.Controls.TextBlock]::new()
                    $badgeTb.FontSize   = 10.5
                    $badgeTb.FontWeight = 'SemiBold'
                    $badgeFg = if ($isCustom) { '#C4B5FD' } else { '#90EEC0' }
                    $badgeTb.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                        [System.Windows.Media.ColorConverter]::ConvertFromString($badgeFg))
                    $badgeTb.Text       = "$([char]0x2713) $($m['RoleName'])"
                    $badge.Child        = $badgeTb
                    $script:SpPimRoles.Children.Add($badge) | Out-Null

                    Write-VerboseLog "PIM: $([char]0x2713) $($m['RoleName'])  ($($m['AssignType']))" -Level Info
                }
                'none' {
                    # API returned 200 but zero instances  -  no active roles right now
                    $script:SpPimRoles.Children.Clear()
                    $tb = [System.Windows.Controls.TextBlock]::new()
                    $tb.Text      = "No active PIM roles detected"
                    $tb.FontSize  = 11
                    $tb.VerticalAlignment = 'Center'
                    $tb.ToolTip   = "Navigate to Entra ID > Privileged Identity Management > My roles to activate"
                    $tb.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                        [System.Windows.Media.ColorConverter]::ConvertFromString('#92400E'))
                    $script:SpPimRoles.Children.Add($tb) | Out-Null
                    Write-VerboseLog "PIM: No active PIM roles detected" -Level Warning
                }
                'nopim' {
                    # 403 response  -  no active PIM roles or RoleManagement.Read.Directory not consented
                    # Treat the same as 'none': roles are not active
                    $script:SpPimRoles.Children.Clear()
                    $tb = [System.Windows.Controls.TextBlock]::new()
                    $tb.Text      = "No active PIM roles detected"
                    $tb.FontSize  = 11
                    $tb.VerticalAlignment = 'Center'
                    $tb.ToolTip   = "403 response from Graph  -  either no PIM roles are active, or RoleManagement.Read.Directory permission is not granted on the App Registration.`nNavigate to Entra ID > Privileged Identity Management > My roles to activate."
                    $tb.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                        [System.Windows.Media.ColorConverter]::ConvertFromString('#92400E'))
                    $script:SpPimRoles.Children.Add($tb) | Out-Null
                    Write-VerboseLog "PIM: No active PIM roles detected (transitiveMemberOf returned 403)" -Level Warning
                    Write-VerboseLog "PIM: To activate: Entra ID > Privileged Identity Management > My roles" -Level Info
                }
                'error' {
                    # Unexpected error  -  not a permission issue
                    $script:SpPimRoles.Children.Clear()
                    $tb = [System.Windows.Controls.TextBlock]::new()
                    $tb.Text      = "PIM check error  -  see verbose log"
                    $tb.FontSize  = 11
                    $tb.VerticalAlignment = 'Center'
                    $tb.ToolTip   = $m['Msg']
                    $tb.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                        [System.Windows.Media.ColorConverter]::ConvertFromString('#F06C6C'))
                    $script:SpPimRoles.Children.Add($tb) | Out-Null
                    Write-VerboseLog "PIM check error: $($m['Msg'])" -Level Error
                }
                'done' {
                    $script:PimTimer.Stop()
                    $script:PimPs.EndInvoke($script:PimHandle)
                    $script:PimRs.Close(); $script:PimPs.Dispose()
                    $script:PimPs = $null; $script:PimRs = $null
                }
            }
        }
    })
    $script:PimTimer.Start()
}

# ── PIM refresh button ──
$script:Window.FindName('BtnPimRefresh').Add_Click({
    Write-VerboseLog "PIM roles  -  manual refresh." -Level Info
    Invoke-PimCheck
})

# ── PIM auto-refresh timer ────────────────────────────────────────────────────
# Fires every PimRefreshIntervalMinutes. Skipped silently if a background
# operation is running ($script:BgPs is not null = job in progress).
function Start-PimAutoRefresh {
    if ($DevConfig.PimRefreshIntervalMinutes -le 0) { return }
    if ($null -ne $script:PimAutoTimer) {
        $script:PimAutoTimer.Stop()
    }
    $script:PimAutoTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:PimAutoTimer.Interval = [TimeSpan]::FromMinutes($DevConfig.PimRefreshIntervalMinutes)
    $script:PimAutoTimer.Add_Tick({
        # Skip if a background Graph operation is currently running
        if ($null -ne $script:BgPs) {
            Write-VerboseLog "PIM auto-refresh skipped  -  background operation in progress." -Level Info
            return
        }
        # Skip if PIM check itself is already running
        if ($null -ne $script:PimPs) {
            return
        }
        Write-VerboseLog "PIM auto-refresh (every $($DevConfig.PimRefreshIntervalMinutes) min)." -Level Info
        Invoke-PimCheck
    })
    $script:PimAutoTimer.Start()
    Write-VerboseLog "PIM auto-refresh enabled  -  interval: $($DevConfig.PimRefreshIntervalMinutes) min." -Level Info
}

function Stop-PimAutoRefresh {
    if ($null -ne $script:PimAutoTimer) {
        $script:PimAutoTimer.Stop()
        $script:PimAutoTimer = $null
    }
}

# Helper: build StackPanel content for Stop-style buttons (icon + label)
function New-StopBtnContent {
    param([string]$Text)
    $sp   = [System.Windows.Controls.StackPanel]::new()
    $sp.Orientation = [System.Windows.Controls.Orientation]::Horizontal
    $ico  = [System.Windows.Controls.TextBlock]::new()
    $ico.Text        = [char]0xE71A
    $ico.FontFamily  = [System.Windows.Media.FontFamily]::new("Segoe MDL2 Assets")
    $ico.FontSize    = 12
    $ico.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    $ico.Margin      = [System.Windows.Thickness]::new(0,0,5,0)
    $ico.Foreground  = [System.Windows.Media.Brushes]::White
    $lbl  = [System.Windows.Controls.TextBlock]::new()
    $lbl.Text              = $Text
    $lbl.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    $null = $sp.Children.Add($ico)
    $null = $sp.Children.Add($lbl)
    return $sp
}

# ── STOP button ──
$script:Window.FindName('BtnStop').Add_Click({
    Write-VerboseLog "Stop requested by user  -  stopping after current object..." -Level Warning
    $script:BgStopped = $true
    if ($null -ne $script:SharedBg) { $script:SharedBg['StopRequested'] = $true }
    if ($null -ne $script:GdiStop)  { $script:GdiStop['Stop'] = $true }
    if ($null -ne $script:OmStop)  { $script:OmStop['Stop'] = $true }
    if ($null -ne $script:DaStop)  { $script:DaStop['Stop'] = $true }
    if ($null -ne $script:FGOStop)  { $script:FGOStop['Stop'] = $true }
    $script:Window.FindName('BtnStop').IsEnabled = $false
    $script:Window.FindName('BtnStop').Content = (New-StopBtnContent "Stopping...")
})

# Disable operation buttons until connected
foreach ($b in @('BtnOpSearch','BtnOpListMembers','BtnOpObjectMembership','BtnOpFindGroupsByOwners','BtnOpCreate','BtnOpAdd','BtnOpRemove','BtnOpExport','BtnOpDynamic','BtnOpRename','BtnOpOwner','BtnOpGetPolicyAssignments','BtnOpUserDevices','BtnOpFindCommon','BtnOpFindDistinct','BtnOpGetDeviceInfo','BtnOpGetDiscoveredApps','BtnOpCompareGroups','BtnOpRemoveGroups')) {
    $btn = $script:Window.FindName($b)
    if ($btn) { $btn.IsEnabled = $false }
}

# ── Op tile buttons ──
$script:Window.FindName('BtnOpCreate').Add_Click({
    Show-Panel 'PanelCreate'
    Hide-Notification
    Write-VerboseLog "Panel: Create Group" -Level Info

    # ── Fetch Administrative Units on wizard open ──
    $cmbAU = $script:Window.FindName('CmbCreateAU')
    $txtAULoad = $script:Window.FindName('TxtAULoading')

    # Reset AU ComboBox to (None) + show loading indicator
    $cmbAU.Items.Clear()
    $noneItem = [System.Windows.Controls.ComboBoxItem]::new()
    $noneItem.Content = '(None)'
    $noneItem.Tag     = ''
    $noneItem.IsSelected = $true
    $cmbAU.Items.Add($noneItem) | Out-Null
    $cmbAU.SelectedIndex = 0
    $txtAULoad.Visibility = 'Visible'

    Invoke-OnBackground -DisableButton $null -BusyMessage $null -Work {
        try {
            $auList = [System.Collections.Generic.List[PSCustomObject]]::new()
            $uri = "https://graph.microsoft.com/v1.0/directory/administrativeUnits?`$select=id,displayName,description&`$top=999"
            while ($uri) {
                $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
                foreach ($au in $resp['value']) {
                    $auList.Add([PSCustomObject]@{
                        Id          = $au['id']
                        DisplayName = $au['displayName']
                        Description = if ($au['description']) { $au['description'] } else { '' }
                    })
                }
                $uri = $resp['@odata.nextLink']
            }
            $script:AdminUnits = @($auList | Sort-Object DisplayName)
            Write-VerboseLog "Fetched $($script:AdminUnits.Count) Administrative Unit(s)" -Level Info
        } catch {
            $script:AdminUnits = @()
            Write-VerboseLog "Could not fetch Administrative Units: $($_.Exception.Message)" -Level Warning
        }
    } -Done {
        $cmbAU     = $script:Window.FindName('CmbCreateAU')
        $txtAULoad = $script:Window.FindName('TxtAULoading')
        $txtAULoad.Visibility = 'Collapsed'

        foreach ($au in $script:AdminUnits) {
            $item = [System.Windows.Controls.ComboBoxItem]::new()
            $item.Content = $au.DisplayName
            $item.Tag     = $au.Id
            $item.ToolTip = if ($au.Description) { $au.Description } else { $au.Id }
            $cmbAU.Items.Add($item) | Out-Null
        }
        if ($script:AdminUnits.Count -eq 0) {
            Write-VerboseLog "No Administrative Units found (or permission denied)." -Level Warning
        }
    }
})

$script:Window.FindName('BtnOpAdd').Add_Click({
    $script:CurrentMemberOp = 'Add'
    $script:Window.FindName('TxtMembershipTitle').Text       = "Add Members"
    $script:Window.FindName('TxtMemberInputLabel').Text      = "MEMBERS TO ADD  -  one per line (User UPNs / Group names or IDs / Device names or IDs)"
    $script:Window.FindName('BtnMemberValidate').Content     = "Validate & Preview ->"
    # Reset state
    $script:Window.FindName('TxtGroupSearch').Text           = ""
    $script:Window.FindName('TxtMemberList').Text            = ""
    $script:Window.FindName('SelectedGroupBadge').Visibility = 'Collapsed'
    $script:Window.FindName('GroupSearchResults').Visibility = 'Collapsed'
    $script:SelectedGroup = $null
    Show-Panel 'PanelMembership'
    Hide-Notification
    Write-VerboseLog "Panel: Add Members" -Level Info
})

$script:Window.FindName('BtnOpRemove').Add_Click({
    $script:CurrentMemberOp = 'Remove'
    $script:Window.FindName('TxtMembershipTitle').Text       = "Remove Members"
    $script:Window.FindName('TxtMemberInputLabel').Text      = "MEMBERS TO REMOVE  -  one per line (User UPNs / Group names or IDs / Device names or IDs)"
    $script:Window.FindName('BtnMemberValidate').Content     = "Validate & Preview ->"
    $script:Window.FindName('TxtGroupSearch').Text           = ""
    $script:Window.FindName('TxtMemberList').Text            = ""
    $script:Window.FindName('SelectedGroupBadge').Visibility = 'Collapsed'
    $script:Window.FindName('GroupSearchResults').Visibility = 'Collapsed'
    $script:SelectedGroup = $null
    Show-Panel 'PanelMembership'
    Hide-Notification
    Write-VerboseLog "Panel: Remove Members" -Level Info
})

$script:Window.FindName('BtnOpExport').Add_Click({
    $script:Window.FindName('TxtExportGroupSearch').Text           = ""
    $script:Window.FindName('TxtExportPath').Text                  = ""
    $script:Window.FindName('ExportSelectedGroupBadge').Visibility = 'Collapsed'
    $script:Window.FindName('ExportGroupSearchResults').Visibility = 'Collapsed'
    $script:ExportSelectedGroup = $null
    Show-Panel 'PanelExport'
    Hide-Notification
    Write-VerboseLog "Panel: Export Members" -Level Info
})

$script:Window.FindName('BtnOpDynamic').Add_Click({
    $script:Window.FindName('TxtDynGroupSearch').Text           = ""
    $script:Window.FindName('TxtDynRule').Text                  = ""
    $script:Window.FindName('DynSelectedGroupBadge').Visibility = 'Collapsed'
    $script:Window.FindName('DynGroupSearchResults').Visibility = 'Collapsed'
    $script:DynSelectedGroup = $null
    Show-Panel 'PanelDynamic'
    Hide-Notification
    Write-VerboseLog "Panel: Set Dynamic Query" -Level Info
})

$script:Window.FindName('BtnOpRename').Add_Click({
    $script:Window.FindName('TxtRenGroupSearch').Text           = ""
    $script:Window.FindName('TxtNewGroupName').Text             = ""
    $script:Window.FindName('RenSelectedGroupBadge').Visibility = 'Collapsed'
    $script:Window.FindName('RenGroupSearchResults').Visibility = 'Collapsed'
    $script:RenSelectedGroup = $null
    Show-Panel 'PanelRename'
    Hide-Notification
    Write-VerboseLog "Panel: Rename Group" -Level Info
})

$script:Window.FindName('BtnOpOwner').Add_Click({
    $script:Window.FindName('TxtOwnerGroupList').Text = ""
    $script:Window.FindName('TxtOwnerList').Text      = ""
    $script:OwnerGroupsTxt       = ''
    $script:OwnerGroupsValidated = $null
    Show-Panel 'PanelOwner'
    Hide-Notification
    Write-VerboseLog "Panel: Set Group Owner" -Level Info
})

$script:Window.FindName('BtnOpSearch').Add_Click({
    $script:Window.FindName('TxtSearchKeyword').Text         = ""
    $script:Window.FindName('PnlSearchResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtSearchNoResults').Visibility = 'Collapsed'
    $script:SearchResults = $null
    Show-Panel 'PanelSearch'
    Hide-Notification
    Write-VerboseLog "Panel: Search Entra Objects" -Level Info
})

$script:Window.FindName('BtnOpListMembers').Add_Click({
    $script:Window.FindName('PnlLMResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtLMNoResults').Visibility = 'Collapsed'
    $script:Window.FindName('PnlLMProgress').Visibility  = 'Collapsed'
    Show-Panel 'PanelListMembers'
    Hide-Notification
    Write-VerboseLog "Panel: List Group Members" -Level Info
})

$script:Window.FindName('BtnOpObjectMembership').Add_Click({
    $script:Window.FindName('PnlOMResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtOMNoResults').Visibility = 'Collapsed'
    $script:Window.FindName('PnlOMProgress').Visibility  = 'Collapsed'
    $script:Window.FindName('TxtOMInputList').Text       = ''
    $script:Window.FindName('CmbOMObjectType').SelectedIndex = 0
    $script:Window.FindName('TxtOMInputLabel').Text = 'USER UPNs  -  one per line'
    $dg = $script:Window.FindName('DgOMResults')
    if ($dg) { $dg.ItemsSource = $null }
    $script:Window.FindName('BtnOMCopyValue').IsEnabled  = $false
    $script:Window.FindName('BtnOMCopyRow').IsEnabled    = $false
    $script:Window.FindName('BtnOMCopyAll').IsEnabled    = $false
    $script:Window.FindName('TxtOMFilter').IsEnabled  = $false
    $script:Window.FindName('BtnOMFilter').IsEnabled  = $false
    $script:Window.FindName('BtnOMFilterClear').IsEnabled = $false
    $script:Window.FindName('TxtOMFilter').Text       = ''
    $script:OMAllData = $null
    $script:Window.FindName('BtnOMExportXlsx').IsEnabled = $false
    Show-Panel 'PanelObjectMembership'
    Hide-Notification
    Write-VerboseLog "Panel: Object Membership" -Level Info
})

$script:Window.FindName('CmbOMObjectType').Add_SelectionChanged({
    $sel = $script:Window.FindName('CmbOMObjectType').SelectedItem.Content
    $label = switch ($sel) {
        'Users (UPN)'        { 'USER UPNs  -  one per line' }
        'Devices (Name)'     { 'DEVICE NAMES  -  one per line' }
        'Devices (ID)'       { 'DEVICE IDs  -  one per line' }
        'Security Groups'    { 'SECURITY GROUP names or IDs  -  one per line' }
        'M365 Groups'        { 'M365 GROUP names or IDs  -  one per line' }
        default              { 'OBJECTS  -  one per line' }
    }
    $script:Window.FindName('TxtOMInputLabel').Text = $label
})


# ── OBJECT MEMBERSHIP  -  Run ─────────────────────────────────────────
$script:Window.FindName('BtnOMRun').Add_Click({
    Hide-Notification

    $inputTxt   = $script:Window.FindName('TxtOMInputList').Text
    $objTypeSel = $script:Window.FindName('CmbOMObjectType').SelectedItem.Content

    if ([string]::IsNullOrWhiteSpace($inputTxt)) {
        Show-Notification 'Please enter at least one identifier.' -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }

    $script:OMParams = @{
        InputTxt  = $inputTxt
        ObjType   = $objTypeSel
    }

    # Reset UI
    $script:Window.FindName('PnlOMResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtOMNoResults').Visibility = 'Collapsed'
    $script:Window.FindName('BtnOMCopyValue').IsEnabled  = $false
    $script:Window.FindName('BtnOMCopyRow').IsEnabled    = $false
    $script:Window.FindName('BtnOMCopyAll').IsEnabled    = $false
    $script:Window.FindName('BtnOMFilter').IsEnabled  = $false
    $script:Window.FindName('BtnOMFilterClear').IsEnabled = $false
    $script:Window.FindName('TxtOMFilter').Text       = ''
    $script:OMAllData = $null
    $script:Window.FindName('TxtOMFilter').IsEnabled  = $false
    $script:Window.FindName('BtnOMExportXlsx').IsEnabled = $false
    $dg = $script:Window.FindName('DgOMResults')
    if ($dg) { $dg.ItemsSource = $null }
    $script:Window.FindName('PnlOMProgress').Visibility      = 'Visible'
    $script:Window.FindName('TxtOMProgressMsg').Text         = 'Resolving objects...'
    $script:Window.FindName('TxtOMProgressDetail').Text      = ''

    $script:Window.FindName('BtnOMRun').IsEnabled = $false
    $script:Window.FindName('BtnStop').Visibility = 'Visible'
    $script:Window.FindName('BtnStop').IsEnabled  = $true
    $script:Window.FindName('BtnStop').Content = (New-StopBtnContent 'Stop')
    $script:BgStopped = $false
    Write-VerboseLog '--- Object Membership: starting ---' -Level Action

    $script:OmQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $script:OmStop  = [hashtable]::Synchronized(@{ Stop = $false })

    $script:OmRs = [runspacefactory]::CreateRunspace()
    $script:OmRs.ApartmentState = 'STA'
    $script:OmRs.ThreadOptions  = 'ReuseThread'
    $script:OmRs.Open()
    $script:OmRs.SessionStateProxy.SetVariable('OmQueue',  $script:OmQueue)
    $script:OmRs.SessionStateProxy.SetVariable('OmStop',   $script:OmStop)
    $script:OmRs.SessionStateProxy.SetVariable('OmParams', $script:OMParams)
    $script:OmRs.SessionStateProxy.SetVariable('LogFile',  $script:LogFile)

    $script:OmPs = [powershell]::Create()
    $script:OmPs.Runspace = $script:OmRs

    $null = $script:OmPs.AddScript({
        function OLog {
            param([string]$Msg, [string]$Level = 'Info')
            $prefix = switch ($Level) {
                "Success" { "$([char]0x2713) " } "Warning" { "[!] " } "Error" { "$([char]0x2717) " }
                'Action'  { '> ' } default   { '  ' }
            }
            $line = "[$(Get-Date -Format 'HH:mm:ss')] $prefix$Msg"
            $OmQueue.Enqueue(@{ Type='log'; Line=$line; Level=$Level })
            if ($LogFile) {
                try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
            }
        }

        function Get-GraphAll([string]$Uri) {
            $items = [System.Collections.Generic.List[object]]::new()
            $next  = $Uri
            do {
                $r = Invoke-MgGraphRequest -Method GET -Uri $next
                if ($r['value']) { foreach ($i in @($r['value'])) { $items.Add($i) } }
                $next = if ($r.ContainsKey('@odata.nextLink')) { $r['@odata.nextLink'] } else { $null }
            } while ($next)
            return ,$items
        }

        $isGuidRx = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'

        function Resolve-ObjectId {
            param([string]$Entry, [string]$ObjType)
            $Entry = $Entry.Trim()
            if ([string]::IsNullOrWhiteSpace($Entry)) { return $null }
            switch ($ObjType) {
                'Users (UPN)' {
                    try {
                        $enc = [Uri]::EscapeDataString($Entry)
                        $u = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$enc`?`$select=id,displayName"
                        return @{ Id=[string]$u['id']; DisplayName=[string]$u['displayName']; ResourceType='users' }
                    } catch { return $null }
                }
                'Devices (Name)' {
                    try {
                        $safe = $Entry -replace "'","''"
                        $r = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=displayName eq '$safe'`&`$select=id,displayName`&`$top=1"
                        if ($r['value'] -and $r['value'].Count -gt 0) {
                            $d = $r['value'][0]
                            return @{ Id=[string]$d['id']; DisplayName=[string]$d['displayName']; ResourceType='devices' }
                        }
                        return $null
                    } catch { return $null }
                }
                'Devices (ID)' {
                    try {
                        $enc = [Uri]::EscapeDataString($Entry)
                        $d = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/devices/$enc`?`$select=id,displayName"
                        return @{ Id=[string]$d['id']; DisplayName=[string]$d['displayName']; ResourceType='devices' }
                    } catch { return $null }
                }
                'Security Groups' {
                    try {
                        if ($Entry -match $isGuidRx) {
                            $g = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$Entry`?`$select=id,displayName"
                            return @{ Id=[string]$g['id']; DisplayName=[string]$g['displayName']; ResourceType='groups' }
                        } else {
                            $safe = $Entry -replace "'","''"
                            $r = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe' and securityEnabled eq true and mailEnabled eq false`&`$select=id,displayName`&`$top=1"
                            if ($r['value'] -and $r['value'].Count -gt 0) {
                                $d = $r['value'][0]
                                return @{ Id=[string]$d['id']; DisplayName=[string]$d['displayName']; ResourceType='groups' }
                            }
                            return $null
                        }
                    } catch { return $null }
                }
                'M365 Groups' {
                    try {
                        if ($Entry -match $isGuidRx) {
                            $g = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$Entry`?`$select=id,displayName"
                            return @{ Id=[string]$g['id']; DisplayName=[string]$g['displayName']; ResourceType='groups' }
                        } else {
                            $safe = $Entry -replace "'","''"
                            $r = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe' and groupTypes/any(c:c eq 'Unified')`&`$select=id,displayName`&`$top=1"
                            if ($r['value'] -and $r['value'].Count -gt 0) {
                                $d = $r['value'][0]
                                return @{ Id=[string]$d['id']; DisplayName=[string]$d['displayName']; ResourceType='groups' }
                            }
                            return $null
                        }
                    } catch { return $null }
                }
            }
            return $null
        }

        try {
            $p = $OmParams
            $entries = @($p['InputTxt'] -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
            OLog "Processing $($entries.Count) input(s) - Type: $($p['ObjType'])" 'Action'

            $allRows = [System.Collections.Generic.List[PSCustomObject]]::new()
            $idx = 0

            foreach ($entry in $entries) {
                if ($OmStop['Stop']) { OLog 'Stop requested - halting.' 'Warning'; break }
                $idx++
                $OmQueue.Enqueue(@{ Type='progress'; Msg="Resolving object $idx of $($entries.Count)..."; Detail=$entry })

                OLog "Resolving: $entry" 'Info'
                $resolved = Resolve-ObjectId -Entry $entry -ObjType $p['ObjType']
                if (-not $resolved) {
                    OLog "  Could not resolve: $entry" 'Warning'
                    continue
                }

                $resType = $resolved['ResourceType']
                $resId   = $resolved['Id']
                $resName = $resolved['DisplayName']
                OLog "  Resolved: $resName ($resType/$resId)" 'Success'

                OLog "  Fetching transitive memberships..." 'Info'
                try {
                    $memberships = Get-GraphAll "https://graph.microsoft.com/v1.0/$resType/$resId/transitiveMemberOf?`$select=id,displayName,description,groupTypes,securityEnabled,mailEnabled`&`$top=999"
                } catch {
                    OLog "  transitiveMemberOf failed: $($_.Exception.Message)" 'Warning'
                    continue
                }

                OLog "  $($memberships.Count) group membership(s) found" 'Info'

                # Fetch direct memberships to distinguish Direct vs Transitive
                $directIds = [System.Collections.Generic.HashSet[string]]::new()
                try {
                    $directMembers = Get-GraphAll "https://graph.microsoft.com/v1.0/$resType/$resId/memberOf?`$select=id`&`$top=999"
                    foreach ($dm in $directMembers) {
                        $null = $directIds.Add([string]$dm['id'])
                    }
                } catch {}

                foreach ($grp in $memberships) {
                    if ($OmStop['Stop']) { break }
                    $odataType = [string]$grp['@odata.type']
                    if ($odataType -ne '#microsoft.graph.group') { continue }

                    $gId    = [string]$grp['id']
                    $gName  = [string]$grp['displayName']
                    $gDesc  = [string]$grp['description']
                    $gTypes = @($grp['groupTypes'])
                    $secEn  = $grp['securityEnabled']
                    $mailEn = $grp['mailEnabled']

                    $groupType = if ($gTypes -contains 'Unified') { 'M365' }
                                 elseif ($secEn -eq $true -and $mailEn -eq $false) { 'Security' }
                                 elseif ($secEn -eq $true -and $mailEn -eq $true)  { 'Mail-Enabled Security' }
                                 elseif ($secEn -eq $false -and $mailEn -eq $true) { 'Distribution' }
                                 else { 'Other' }

                    $memType = if ($directIds.Contains($gId)) { 'Direct' } else { 'Transitive' }

                    $objTypeDisplay = switch ($resType) {
                        'users'   { 'User' }
                        'devices' { 'Device' }
                        'groups'  { 'Group' }
                    }

                    $allRows.Add([PSCustomObject]@{
                        InputObject    = $entry
                        ObjectType     = $objTypeDisplay
                        GroupName      = $gName
                        GroupId        = $gId
                        GroupType      = $groupType
                        Description    = $gDesc
                        MembershipType = $memType
                    })
                }
            }

            OLog "Completed: $($allRows.Count) membership record(s) found." 'Success'
            $OmQueue.Enqueue(@{ Type='done'; Rows=$allRows; Error=$null })

        } catch {
            $OmQueue.Enqueue(@{ Type='done'; Rows=@(); Error=$_.Exception.Message })
        }
    })

    $script:OmHandle = $script:OmPs.BeginInvoke()

    # Timer to drain OmQueue on the UI thread
    $script:OmTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:OmTimer.Interval = [TimeSpan]::FromMilliseconds(300)
    $script:OmTimer.Add_Tick({
        $entry = $null
        while ($script:OmQueue.TryDequeue([ref]$entry)) {
            switch ($entry['Type']) {
                'log' {
                    $lvl = $entry['Level']
                    $col = switch ($lvl) {
                        'Success' { $DevConfig.LogColorSuccess }
                        'Warning' { $DevConfig.LogColorWarning }
                        'Error'   { $DevConfig.LogColorError   }
                        'Action'  { $DevConfig.LogColorAction  }
                        default   { $DevConfig.LogColorInfo    }
                    }
                    try {
                        $para = [System.Windows.Documents.Paragraph]::new()
                        $run  = [System.Windows.Documents.Run]::new($entry['Line'])
                        $run.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                            [System.Windows.Media.ColorConverter]::ConvertFromString($col))
                        $para.Margin = [System.Windows.Thickness]::new(0)
                        $para.Inlines.Add($run)
                        $script:RtbLog.Document.Blocks.Add($para)
                        $script:RtbLog.ScrollToEnd()
                    } catch {}
                }
                'progress' {
                    $script:Window.FindName('PnlOMProgress').Visibility       = 'Visible'
                    $script:Window.FindName('TxtOMProgressMsg').Text          = $entry['Msg']
                    $script:Window.FindName('TxtOMProgressDetail').Text       = $entry['Detail']
                    Update-StatusBar -Text "Object Membership: $($entry['Msg'])"
                }
                'done' {
                    $script:OmTimer.Stop()
                    try { $script:OmPs.EndInvoke($script:OmHandle) } catch {}
                    try { $script:OmRs.Close() }                      catch {}
                    try { $script:OmPs.Dispose() }                    catch {}
                    $script:OmPs = $null; $script:OmRs = $null
                    $script:Window.FindName('BtnStop').Visibility = 'Collapsed'
                    $script:Window.FindName('BtnStop').IsEnabled  = $true
                    $script:Window.FindName('BtnStop').Content = (New-StopBtnContent 'Stop')
                    $script:Window.FindName('BtnOMRun').IsEnabled = $true
                    $script:Window.FindName('PnlOMProgress').Visibility = 'Collapsed'

                    $errMsg = $entry['Error']
                    if ($errMsg) {
                        Write-VerboseLog "Object Membership error: $errMsg" -Level Error
                        Show-Notification "Error: $errMsg" -BgColor '#F8D7DA' -FgColor '#721C24'
                        Update-StatusBar -Text 'Connected'
                        return
                    }

                    $rows = @($entry['Rows'])
                    if ($rows.Count -eq 0) {
                        Write-VerboseLog 'Object Membership: no results.' -Level Warning
                        $script:Window.FindName('TxtOMNoResults').Visibility = 'Visible'
                        Show-Notification 'No group memberships found.' -BgColor '#FFF3CD' -FgColor '#7A4800'
                        Update-StatusBar -Text 'Connected'
                        return
                    }

                    $dg = $script:Window.FindName('DgOMResults')
                    $dg.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSCustomObject]]::new($rows)
                    $script:Window.FindName('TxtOMCount').Text            = "$($rows.Count) membership(s) found"
                    $script:Window.FindName('PnlOMResults').Visibility     = 'Visible'
                    $script:Window.FindName('BtnOMCopyAll').IsEnabled      = $true
                    $script:Window.FindName('TxtOMFilter').IsEnabled  = $true
                    $script:Window.FindName('BtnOMFilter').IsEnabled  = $true
                    $script:Window.FindName('BtnOMFilterClear').IsEnabled = $true
                    $script:Window.FindName('BtnOMExportXlsx').IsEnabled   = $true
                    Write-VerboseLog "Object Membership: $($rows.Count) record(s) loaded." -Level Success
                    Show-Notification "$($rows.Count) membership(s) ready." -BgColor '#D4EDDA' -FgColor '#155724'
                    Update-StatusBar -Text 'Connected'
                }
            }
        }
    })
    $script:OmTimer.Start()
})


# ── DgOMResults: cell selection ──
$script:Window.FindName('DgOMResults').Add_SelectedCellsChanged({
    $dg       = $script:Window.FindName('DgOMResults')
    $selCells = $dg.SelectedCells.Count
    $script:Window.FindName('BtnOMCopyValue').IsEnabled = ($selCells -ge 1)
    $script:Window.FindName('BtnOMCopyRow').IsEnabled   = ($selCells -gt 0)
})

# ── BtnOMCopyValue ──
# ── BtnOMFilter ──
$script:Window.FindName('BtnOMFilter').Add_Click({
    $dg      = $script:Window.FindName('DgOMResults')
    $keyword = $script:Window.FindName('TxtOMFilter').Text.Trim()
    if ($null -eq $script:OMAllData) {
        $script:OMAllData = @($dg.ItemsSource)
    }
    if ($keyword -eq '') {
        $dg.ItemsSource = $script:OMAllData
        Show-Notification "Filter cleared - $($script:OMAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    } else {
        $filtered = @($script:OMAllData | Where-Object {
            $row = $_
            $match = $false
            foreach ($prop in $row.PSObject.Properties) {
                if ([string]$prop.Value -like "*$keyword*") { $match = $true; break }
            }
            $match
        })

# ── BtnOMFilterClear ──
$script:Window.FindName('BtnOMFilterClear').Add_Click({
    $script:Window.FindName('TxtOMFilter').Text = ''
    $dg = $script:Window.FindName('DgOMResults')
    if ($null -ne $script:OMAllData) {
        $dg.ItemsSource = $script:OMAllData
        Show-Notification "Filter cleared - $($script:OMAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    }
})
        $dg.ItemsSource = $filtered
        Show-Notification "$($filtered.Count) of $($script:OMAllData.Count) result(s) shown after filter." -BgColor '#D4EDDA' -FgColor '#155724'
    }
    Write-VerboseLog "OM: Filter applied - keyword='$keyword'" -Level Info
})

$script:Window.FindName('BtnOMCopyValue').Add_Click({
    $dg = $script:Window.FindName('DgOMResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $vals = [System.Collections.Generic.List[string]]::new()
    foreach ($cellInfo in @($dg.SelectedCells)) {
        $col = $cellInfo.Column -as [System.Windows.Controls.DataGridBoundColumn]
        if ($col) { $vals.Add(([string]$cellInfo.Item.($col.Binding.Path.Path)).Trim()) }
    }
    [System.Windows.Clipboard]::SetText($vals -join "`r`n")
    Show-Notification '$($vals.Count) value(s) copied to clipboard.' -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog 'OM: Copied cell value to clipboard.' -Level Success
})

# ── BtnOMCopyRow ──
$script:Window.FindName('BtnOMCopyRow').Add_Click({
    $dg = $script:Window.FindName('DgOMResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $seen = [System.Collections.Generic.HashSet[object]]::new()
    $rowObjs = [System.Collections.Generic.List[object]]::new()
    foreach ($cell in @($dg.SelectedCells)) { if ($seen.Add($cell.Item)) { $rowObjs.Add($cell.Item) } }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc     = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $txtLines = [System.Collections.Generic.List[string]]::new()
    $txtLines.Add(($colDefs | ForEach-Object { $_.H }) -join "`t")
    foreach ($row in $rowObjs) {
        $vals = $colDefs | ForEach-Object { [string]$row.($_.P) }
        $txtLines.Add($vals -join "`t")
    }
    [System.Windows.Clipboard]::SetText($txtLines -join "`n")
    Show-Notification "$($rowObjs.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "OM: Copied $($rowObjs.Count) row(s) to clipboard." -Level Success
    # Highlight full rows
    $dg.UnselectAllCells()
    foreach ($r in $rowObjs) { $null = $dg.SelectedItems.Add($r) }
})

# ── BtnOMCopyAll ──
$script:Window.FindName('BtnOMCopyAll').Add_Click({
    $dg       = $script:Window.FindName('DgOMResults')
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) { return }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc     = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $txtLines = [System.Collections.Generic.List[string]]::new()
    $txtLines.Add(($colDefs | ForEach-Object { $_.H }) -join "`t")
    foreach ($row in $allItems) {
        $vals = $colDefs | ForEach-Object { [string]$row.($_.P) }
        $txtLines.Add($vals -join "`t")
    }
    [System.Windows.Clipboard]::SetText($txtLines -join "`n")
    Show-Notification "$($allItems.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "OM: Copied all $($allItems.Count) row(s) to clipboard." -Level Success
})

# ── BtnOMExportXlsx ──
$script:Window.FindName('BtnOMExportXlsx').Add_Click({
    $dg       = $script:Window.FindName('DgOMResults')
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) {
        Show-Notification 'No results to export.' -BgColor '#FFF3CD' -FgColor '#856404'
        return
    }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc     = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title            = 'Export Object Membership Results'
    $dlg.Filter           = 'Excel Workbook (*.xlsx)|*.xlsx|CSV File (*.csv)|*.csv|All Files (*.*)|*.*'
    $dlg.FileName         = "ObjectMembership_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($dlg.ShowDialog() -ne $true) { return }
    $outPath = $dlg.FileName
    try {
        $folder = Split-Path $outPath -Parent
        if ($folder -and -not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        $exportRows = $allItems | ForEach-Object {
            $rowObj = $_; $ht = [ordered]@{}
            foreach ($cd in $colDefs) { $ht[$cd['H']] = [string]$rowObj.($cd['P']) }
            [PSCustomObject]$ht
        }
        if (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue) {
            Import-Module ImportExcel -ErrorAction SilentlyContinue
            $exportRows | Export-Excel -Path $outPath -WorksheetName 'ObjectMembership' -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium6
            Show-Notification "Exported $($allItems.Count) rows: $outPath" -BgColor '#D4EDDA' -FgColor '#155724'
        } else {
            $exportRows | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8
            Show-Notification "ImportExcel not found - saved as CSV: $(Split-Path $outPath -Leaf)" -BgColor '#FFF3CD' -FgColor '#856404'
        }
        Write-VerboseLog "OM: Exported $($allItems.Count) row(s) to: $outPath" -Level Success
    } catch {
        Write-VerboseLog "OM Export failed: $($_.Exception.Message)" -Level Error
        Show-Notification "Export failed: $($_.Exception.Message)" -BgColor '#F8D7DA' -FgColor '#721C24'
    }
})




# ============================================================
# BACKGROUND WORKER HELPER
# Uses a [hashtable]::Synchronized() shared between the UI thread
# and the background runspace  -  no serialisation, no write-back needed.
# $script:SharedBg is the live object both sides read/write directly.
# ============================================================
function Invoke-OnBackground {
    param(
        [scriptblock]$Work,
        [scriptblock]$Done,
        [string]$BusyMessage = "Working...",
        [System.Windows.Controls.Button]$DisableButton = $null
    )

    if ($DisableButton) { $DisableButton.IsEnabled = $false }
    $script:BgStopped = $false
    $script:Window.FindName('BtnStop').Visibility = 'Visible'
    Write-VerboseLog $BusyMessage -Level Action

    # Initialise / reset the shared state bag
    $script:SharedBg = [hashtable]::Synchronized(@{
        # Thread-safe log queue  -  runspace appends, timer drains to RichTextBox
        LogQueue            = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
        # Inputs  -  written by UI thread before runspace starts
        CreateParams        = $script:CreateParams
        CreateValidated     = $script:CreateValidated
        MemberParams        = $script:MemberParams
        MemberValidated     = $script:MemberValidated
        ExportOutPath       = $script:ExportOutPath
        ExportCount         = $script:ExportCount
        DynRule             = $script:DynRule
        RenNewName          = $script:RenNewName
        RenExistingId       = $script:RenExistingId
        OwnerUpnsTxt         = $script:OwnerUpnsTxt
        OwnerValidated       = $script:OwnerValidated
        OwnerExecResult      = $script:OwnerExecResult
        MemberExecResult    = $script:MemberExecResult
        CreateNewGroupId    = $script:CreateNewGroupId
        SelectedGroup       = $script:SelectedGroup
        ExportSelectedGroup = $script:ExportSelectedGroup
        DynSelectedGroup     = $script:DynSelectedGroup
        RenSelectedGroup     = $script:RenSelectedGroup
        OwnerGroupsTxt       = $script:OwnerGroupsTxt
        OwnerGroupsValidated = $script:OwnerGroupsValidated
        UDSelectedGroup      = $script:UDSelectedGroup
        UDParams            = $script:UDParams
        UDValidated         = $script:UDValidated
        UDExecResult        = $script:UDExecResult
        FCParams            = $script:FCParams
        FCResult            = $script:FCResult
        FDParams            = $script:FDParams
        FDResult            = $script:FDResult
        SensitivityLabels   = $script:SensitivityLabels
        AdminUnits          = $script:AdminUnits
        LogFile             = $script:LogFile
        SearchKeyword        = $script:SearchKeyword
        SearchEntries        = $script:SearchEntries
        SearchExactMatch     = $script:SearchExactMatch
        SearchTypes          = $script:SearchTypes
        SearchGetManager     = $script:SearchGetManager
        LMGroupInput         = $script:LMGroupInput
        LMAllResults         = $script:LMAllResults
        LMProgress           = $null
        CGGroupInput         = $script:CGGroupInput
        CGAllResults         = $script:CGAllResults
        CGProgress           = $null
        GPAParams            = $script:GPAParams
        GPAResult            = $null
        RGGroupList          = $script:RGGroupList
        RGValidated          = $script:RGValidated
        RGExecResult         = $null
        # Signalling
        Status              = 'running'   # 'ok' | 'error' | 'stopped'
        ErrorMessage        = ''
        StopRequested       = $false
    })

    # Build function definitions to inject
    $fnNames = @(
        'Get-Timestamp','Write-VerboseLog','Invoke-GraphGet','Resolve-GroupByNameOrId',
        'Search-Groups','Resolve-MemberEntry','Validate-InputList',
        'Get-GroupMemberCount','Get-GroupMembers','Add-MembersToGroup',
        'Remove-MembersFromGroup','Export-GroupMembersToXlsx',
        'Get-GPASubType','Get-GPAPlatform'
    )
    $fnScript = [System.Text.StringBuilder]::new()
    foreach ($name in $fnNames) {
        $sb = (Get-Command $name -ErrorAction SilentlyContinue).ScriptBlock
        if ($sb) { $null = $fnScript.AppendLine("function $name {`n$sb`n}") }
    }

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()

    # Pass the SAME synchronized hashtable object  -  no serialisation occurs
    $rs.SessionStateProxy.SetVariable('Shared',  $script:SharedBg)
    $rs.SessionStateProxy.SetVariable('WorkStr', $Work.ToString())
    $rs.SessionStateProxy.SetVariable('DevConfig', $DevConfig)

    $runScript = $fnScript.ToString() + @'
Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue

# Runspace-safe Write-VerboseLog  -  no WPF/Dispatcher access.
# Logs to file and enqueues for the UI timer to drain into the RichTextBox.
function Write-VerboseLog {
    param([string]$Message,
          [ValidateSet("Info","Success","Warning","Error","Action")][string]$Level = "Info")
    $prefix = switch ($Level) {
        "Info"    { "  " } "Success" { "✓ " } "Warning" { "⚠ " }
        "Error"   { "✗ " } "Action"  { "▶ " }
    }
    $ts   = Get-Timestamp
    $line = "[$ts] $prefix$Message"
    if ($script:LogFile) {
        try { Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
    }
    $Shared["LogQueue"].Enqueue(@{ Line = $line; Level = $Level })
}

# Map Shared keys into $script:* so the Work block can use $script:X syntax
$script:CreateParams        = $Shared['CreateParams']
$script:CreateValidated     = $Shared['CreateValidated']
$script:MemberParams        = $Shared['MemberParams']
$script:MemberValidated     = $Shared['MemberValidated']
$script:ExportOutPath       = $Shared['ExportOutPath']
$script:ExportCount         = $Shared['ExportCount']
$script:DynRule             = $Shared['DynRule']
$script:RenNewName          = $Shared['RenNewName']
$script:RenExistingId       = $Shared['RenExistingId']
$script:OwnerUpnsTxt        = $Shared['OwnerUpnsTxt']
$script:OwnerValidated      = $Shared['OwnerValidated']
$script:OwnerExecResult     = $Shared['OwnerExecResult']
$script:MemberExecResult    = $Shared['MemberExecResult']
$script:CreateNewGroupId    = $Shared['CreateNewGroupId']
$script:SelectedGroup       = $Shared['SelectedGroup']
$script:ExportSelectedGroup = $Shared['ExportSelectedGroup']
$script:DynSelectedGroup    = $Shared['DynSelectedGroup']
$script:RenSelectedGroup     = $Shared['RenSelectedGroup']
$script:OwnerGroupsTxt       = $Shared['OwnerGroupsTxt']
$script:OwnerGroupsValidated = $Shared['OwnerGroupsValidated']
$script:UDSelectedGroup      = $Shared['UDSelectedGroup']
$script:UDParams            = $Shared['UDParams']
$script:UDValidated         = $Shared['UDValidated']
$script:UDExecResult        = $Shared['UDExecResult']
$script:FCParams            = $Shared['FCParams']
$script:FCResult            = $Shared['FCResult']
$script:FDParams            = $Shared['FDParams']
$script:FDResult            = $Shared['FDResult']
$script:SensitivityLabels   = $Shared['SensitivityLabels']
$script:AdminUnits          = $Shared['AdminUnits']
$script:LogFile             = $Shared['LogFile']
$script:SearchKeyword        = $Shared['SearchKeyword']
$script:SearchEntries        = $Shared['SearchEntries']
$script:SearchExactMatch     = $Shared['SearchExactMatch']
$script:SearchTypes          = $Shared['SearchTypes']
$script:SearchGetManager     = $Shared['SearchGetManager']
$script:LMGroupInput         = $Shared['LMGroupInput']
$script:LMAllResults         = $Shared['LMAllResults']
$script:CGGroupInput         = $Shared['CGGroupInput']
$script:CGAllResults         = $Shared['CGAllResults']
$script:GPAParams            = $Shared['GPAParams']
$script:RGGroupList  = $Shared['RGGroupList']
$script:RGValidated  = $Shared['RGValidated']
$script:RGExecResult = $Shared['RGExecResult']

try {
    $workBlock = [scriptblock]::Create($WorkStr)
    & $workBlock

    # Write results back into the shared bag  -  UI thread reads these in Done
    $Shared['CreateValidated']  = $script:CreateValidated
    $Shared['MemberValidated']  = $script:MemberValidated
    $Shared['MemberExecResult'] = $script:MemberExecResult
    $Shared['ExportCount']      = $script:ExportCount
    $Shared['CreateNewGroupId'] = $script:CreateNewGroupId
    $Shared['RenExistingId']    = $script:RenExistingId
    $Shared['OwnerValidated']       = $script:OwnerValidated
    $Shared['OwnerExecResult']      = $script:OwnerExecResult
    $Shared['OwnerGroupsValidated'] = $script:OwnerGroupsValidated
    $Shared['UDValidated']      = $script:UDValidated
    $Shared['UDExecResult']     = $script:UDExecResult
    $Shared['FCResult']         = $script:FCResult
    $Shared['FDResult']         = $script:FDResult
    $Shared['LMAllResults']     = $script:LMAllResults
    $Shared['CGAllResults']     = $script:CGAllResults
    $Shared['GPAResult']        = $script:GPAResult
    $Shared['RGValidated']      = $script:RGValidated
    $Shared['RGExecResult']     = $script:RGExecResult
    $Shared['SensitivityLabels'] = $script:SensitivityLabels
    $Shared['AdminUnits']       = $script:AdminUnits
    $Shared['Status']           = 'ok'
} catch {
    $Shared['ErrorMessage'] = $_.Exception.Message
    $Shared['Status']       = 'error'
}
'@

    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    $null = $ps.AddScript($runScript)

    $script:BgPs     = $ps
    $script:BgRs     = $rs
    $script:BgHandle = $ps.BeginInvoke()
    $script:BgDone   = $Done
    $script:BgBtn    = $DisableButton

    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(300)
    $script:BgTimer  = $timer

    $timer.Add_Tick({
        # Drain log queue regardless of completion status
        $shared = $script:SharedBg
        if ($null -ne $shared -and $null -ne $shared['LogQueue']) {
            $entry = $null
            while ($shared['LogQueue'].TryDequeue([ref]$entry)) {
                $line  = $entry['Line']
                $level = $entry['Level']
                $col   = switch ($level) {
                    'Info'    { $DevConfig.LogColorInfo }
                    'Success' { $DevConfig.LogColorSuccess }
                    'Warning' { $DevConfig.LogColorWarning }
                    'Error'   { $DevConfig.LogColorError }
                    'Action'  { $DevConfig.LogColorAction }
                    default   { $DevConfig.LogColorInfo }
                }
                try {
                    $para = [System.Windows.Documents.Paragraph]::new()
                    $run  = [System.Windows.Documents.Run]::new($line)
                    $run.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                        [System.Windows.Media.ColorConverter]::ConvertFromString($col))
                    $para.Margin = [System.Windows.Thickness]::new(0,0,0,0)
                    $para.Inlines.Add($run)
                    $script:RtbLog.Document.Blocks.Add($para)
                    $script:RtbLog.ScrollToEnd()
                } catch {}
            }
        }
        if ($null -eq $shared -or $shared['Status'] -eq 'running') { return }

        $script:BgTimer.Stop()
        $script:Window.FindName('BtnStop').Visibility = 'Collapsed'
        try { $script:BgPs.EndInvoke($script:BgHandle) } catch {}
        try { $script:BgRs.Close() } catch {}
        try { $script:BgPs.Dispose() } catch {}
        if ($script:BgBtn) { $script:BgBtn.IsEnabled = $true }
        $script:Window.FindName('BtnStop').IsEnabled = $true
        $script:Window.FindName('BtnStop').Content = (New-StopBtnContent "Stop")

        if ($script:BgStopped) {
            # ── Stop operation-specific progress timers & collapse panels ──
            if ($script:LMProgressTimer) { $script:LMProgressTimer.Stop(); $script:LMProgressTimer = $null }
            if ($script:CGProgressTimer) { $script:CGProgressTimer.Stop(); $script:CGProgressTimer = $null }
            foreach ($pnl in @('PnlLMProgress','PnlOMProgress','PnlFGOProgress','PnlCGProgress',
                               'PnlFCProgress','PnlFDProgress','PnlGDIProgress','PnlDAProgress',
                               'PnlGPAProgress','PnlRGProgress')) {
                $el = $script:Window.FindName($pnl)
                if ($el) { $el.Visibility = 'Collapsed' }
            }
            Write-VerboseLog "Operation stopped by user." -Level Warning
            Show-Notification "Operation stopped by user." -BgColor '#FFF3CD' -FgColor '#7A4800'
            return
        }

        if ($shared['Status'] -eq 'error') {
            $errMsg = $shared['ErrorMessage']
            Write-VerboseLog "Background error: $errMsg" -Level Error
            Show-Notification "Error: $errMsg" -BgColor '#F8D7DA' -FgColor '#721C24'
        } else {
            # Promote results from shared bag into $script:* for Done block
            $script:CreateValidated  = $shared['CreateValidated']
            $script:MemberValidated  = $shared['MemberValidated']
            $script:MemberExecResult = $shared['MemberExecResult']
            $script:ExportCount      = $shared['ExportCount']
            $script:CreateNewGroupId = $shared['CreateNewGroupId']
            $script:RenExistingId    = $shared['RenExistingId']
            $script:OwnerValidated       = $shared['OwnerValidated']
            $script:OwnerExecResult      = $shared['OwnerExecResult']
            $script:OwnerGroupsValidated = $shared['OwnerGroupsValidated']
            $script:UDValidated          = $shared['UDValidated']
            $script:UDExecResult      = $shared['UDExecResult']
            $script:FCResult          = $shared['FCResult']
            $script:FDResult          = $shared['FDResult']
            $script:SensitivityLabels = $shared['SensitivityLabels']
            $script:AdminUnits       = $shared['AdminUnits']
            $script:DynGroupState    = $shared['DynGroupState']
            $script:SearchResults    = $shared['SearchResults']
            $script:LMAllResults     = $shared['LMAllResults']
            $script:CGAllResults     = $shared['CGAllResults']
            $script:GPAResult         = $shared['GPAResult']
            $script:RGValidated  = $shared['RGValidated']
            $script:RGExecResult = $shared['RGExecResult']
            if ($script:BgDone) { & $script:BgDone }
        }
    })
    $timer.Start()
}


# ── GROUP SEARCH helper (reused across panels) ──
# All element references are resolved by name INSIDE each handler to avoid
# closure capture failure (local variables are out of scope when events fire).
# OnSelected callbacks are stored in a script-scope hashtable keyed by picker name.
function Wire-GroupPicker {
    param(
        [string]$SearchBoxName,
        [string]$SearchBtnName,
        [string]$ByIdBtnName,
        [string]$ResultsListName,
        [string]$ResultsBorderName,
        [string]$BadgeName,
        [string]$BadgeNameTxt,
        [string]$BadgeIdTxt,
        [string]$ScriptVarName,
        [scriptblock]$OnSelected
    )

    $script:PickerCallbacks[$SearchBoxName] = $OnSelected

    # ---- Search button & Enter key ----
    $searchBtn = $script:Window.FindName($SearchBtnName)
    $searchBtn.Tag = "$SearchBoxName|$ResultsListName|$ResultsBorderName"
    $searchBtn.Add_Click({
        param($sender, $e)
        try {
        $parts        = $sender.Tag -split '\|'
        $sBoxName     = $parts[0]
        $lstName      = $parts[1]
        $borderName   = $parts[2]
        $q = $script:Window.FindName($sBoxName).Text.Trim()
        if ([string]::IsNullOrWhiteSpace($q)) { return }
        Write-VerboseLog "Searching groups: $q" -Level Info
        $lst    = $script:Window.FindName($lstName)
        $border = $script:Window.FindName($borderName)
        $lst.Items.Clear()
        $groups = Search-Groups -Query $q
        if ($groups.Count -eq 0) {
            Write-VerboseLog "No groups found matching '$q'" -Level Warning
            $border.Visibility = 'Collapsed'
            return
        }
        foreach ($g in $groups) {
            $typeLabel = if ($g['groupTypes'] -contains 'Unified') { '[M365]' } else { '[Security]' }
            $lst.Items.Add([PSCustomObject]@{
                Display = "$($g['displayName'])  $typeLabel  |  $($g['id'])"
                Id = $g['id']; Name = $g['displayName']; Raw = $g
            })
        }
        $border.Visibility = 'Visible'
        } catch { Write-VerboseLog "Group search error: $($_.Exception.Message)" -Level Error }
    })

    $searchBox = $script:Window.FindName($SearchBoxName)
    $searchBox.Tag = "$SearchBtnName"
    $searchBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq 'Return') {
            $script:Window.FindName($sender.Tag).RaiseEvent(
                [System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)
            )
        }
    })

    # ---- Use as ID button ----
    $byIdBtn = $script:Window.FindName($ByIdBtnName)
    $byIdBtn.Tag = "$SearchBoxName|$BadgeNameTxt|$BadgeIdTxt|$BadgeName|$ResultsBorderName|$ScriptVarName"
    $byIdBtn.Add_Click({
        param($sender, $e)
        $parts        = $sender.Tag -split '\|'
        $sBoxName     = $parts[0]
        $badgeNameTxt = $parts[1]
        $badgeIdTxt   = $parts[2]
        $badgeName    = $parts[3]
        $borderName   = $parts[4]
        $varName      = $parts[5]
        $idVal = $script:Window.FindName($sBoxName).Text.Trim()
        if ([string]::IsNullOrWhiteSpace($idVal)) {
            Show-Notification "Enter a Group Object ID or display name in the text box first."
            return
        }
        Write-VerboseLog "Resolving group: $idVal" -Level Action
        $g = Resolve-GroupByNameOrId -GroupEntry $idVal
        if ($null -eq $g) {
            Write-VerboseLog "Group not found: $idVal" -Level Error
            Show-Notification "Group not found: $idVal" -BgColor '#F8D7DA' -FgColor '#721C24'
            return
        }
        Set-Variable -Name $varName -Value $g -Scope Script
        $script:Window.FindName($badgeNameTxt).Text = $g.DisplayName
        $script:Window.FindName($badgeIdTxt).Text   = "  ($($g.Id))"
        $script:Window.FindName($badgeName).Visibility  = 'Visible'
        $script:Window.FindName($borderName).Visibility = 'Collapsed'
        Write-VerboseLog "Group selected: $($g.DisplayName) ($($g.Id))" -Level Success
        $cb = $script:PickerCallbacks[$sBoxName]
        if ($cb) { & $cb $g }
    })

    # ---- Results list selection ----
    $lstResults = $script:Window.FindName($ResultsListName)
    $lstResults.Tag = "$ResultsListName|$BadgeNameTxt|$BadgeIdTxt|$BadgeName|$ResultsBorderName|$ScriptVarName|$SearchBoxName"
    $lstResults.Add_SelectionChanged({
        param($sender, $e)
        try {
        $parts        = $sender.Tag -split '\|'
        $lstName      = $parts[0]
        $badgeNameTxt = $parts[1]
        $badgeIdTxt   = $parts[2]
        $badgeName    = $parts[3]
        $borderName   = $parts[4]
        $varName      = $parts[5]
        $sBoxName     = $parts[6]
        $sel = $script:Window.FindName($lstName).SelectedItem
        if ($null -eq $sel) { return }
        $g    = $sel.Raw
        $gObj = [PSCustomObject]@{
            Id = $g['id']; DisplayName = $g['displayName']
            GroupTypes = $g['groupTypes']; MembershipRule = $g['membershipRule']
        }
        Set-Variable -Name $varName -Value $gObj -Scope Script
        $script:Window.FindName($badgeNameTxt).Text = $gObj.DisplayName
        $script:Window.FindName($badgeIdTxt).Text   = "  ($($gObj.Id))"
        $script:Window.FindName($badgeName).Visibility  = 'Visible'
        $script:Window.FindName($borderName).Visibility = 'Collapsed'
        Write-VerboseLog "Group selected: $($gObj.DisplayName) ($($gObj.Id))" -Level Success
        $cb = $script:PickerCallbacks[$sBoxName]
        if ($cb) { & $cb $gObj }
        } catch { Write-VerboseLog "Group selection error: $($_.Exception.Message)" -Level Error }
    })
}

# Wire group pickers for each panel
Wire-GroupPicker `
    -SearchBoxName 'TxtGroupSearch' -SearchBtnName 'BtnGroupSearch' -ByIdBtnName 'BtnGroupById' `
    -ResultsListName 'LstGroupResults' -ResultsBorderName 'GroupSearchResults' `
    -BadgeName 'SelectedGroupBadge' -BadgeNameTxt 'TxtSelectedGroupName' -BadgeIdTxt 'TxtSelectedGroupId' `
    -ScriptVarName 'SelectedGroup' `
    -OnSelected {
        param($g)
        $cnt = Get-GroupMemberCount -GroupId $g.Id
        if ($null -ne $cnt) {
            $script:Window.FindName('TxtSelectedGroupMemCount').Text = "  Members: $cnt"
        }
    }

Wire-GroupPicker `
    -SearchBoxName 'TxtExportGroupSearch' -SearchBtnName 'BtnExportGroupSearch' -ByIdBtnName 'BtnExportGroupById' `
    -ResultsListName 'LstExportGroupResults' -ResultsBorderName 'ExportGroupSearchResults' `
    -BadgeName 'ExportSelectedGroupBadge' -BadgeNameTxt 'TxtExportGroupName' -BadgeIdTxt 'TxtExportGroupId' `
    -ScriptVarName 'ExportSelectedGroup'

Wire-GroupPicker `
    -SearchBoxName 'TxtDynGroupSearch' -SearchBtnName 'BtnDynGroupSearch' -ByIdBtnName 'BtnDynGroupById' `
    -ResultsListName 'LstDynGroupResults' -ResultsBorderName 'DynGroupSearchResults' `
    -BadgeName 'DynSelectedGroupBadge' -BadgeNameTxt 'TxtDynGroupName' -BadgeIdTxt 'TxtDynGroupId' `
    -ScriptVarName 'DynSelectedGroup' `
    -OnSelected {
        param($g)
        if ($g.MembershipRule) {
            $script:Window.FindName('TxtDynCurrentRule').Text = "  Current rule: $($g.MembershipRule)"
            $script:Window.FindName('TxtDynRule').Text        = $g.MembershipRule
        }
    }

Wire-GroupPicker `
    -SearchBoxName 'TxtRenGroupSearch' -SearchBtnName 'BtnRenGroupSearch' -ByIdBtnName 'BtnRenGroupById' `
    -ResultsListName 'LstRenGroupResults' -ResultsBorderName 'RenGroupSearchResults' `
    -BadgeName 'RenSelectedGroupBadge' -BadgeNameTxt 'TxtRenGroupCurrentName' -BadgeIdTxt 'TxtRenGroupId' `
    -ScriptVarName 'RenSelectedGroup'

# ── Export browse ──
$script:Window.FindName('BtnExportBrowse').Add_Click({
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title      = "Save Group Members XLSX"
    $dlg.Filter     = "Excel files (*.xlsx)|*.xlsx"
    $dlg.DefaultExt = ".xlsx"
    $dlg.FileName   = "GroupMembers_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    if ($dlg.ShowDialog()) {
        $script:Window.FindName('TxtExportPath').Text = $dlg.FileName
    }
})



# ── FIND GROUPS BY OWNERS  -  Tile Click ──────────────────────────────────────
$script:Window.FindName('BtnOpFindGroupsByOwners').Add_Click({
    $script:Window.FindName('PnlFGOResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtFGONoResults').Visibility = 'Collapsed'
    $script:Window.FindName('PnlFGOProgress').Visibility  = 'Collapsed'
    $script:Window.FindName('TxtFGOInputList').Text       = ''
    $dg = $script:Window.FindName('DgFGOResults')
    if ($dg) { $dg.ItemsSource = $null }
    $script:Window.FindName('BtnFGOCopyValue').IsEnabled  = $false
    $script:Window.FindName('BtnFGOCopyRow').IsEnabled    = $false
    $script:Window.FindName('BtnFGOCopyAll').IsEnabled    = $false
    $script:Window.FindName('TxtFGOFilter').IsEnabled  = $false
    $script:Window.FindName('BtnFGOFilter').IsEnabled  = $false
    $script:Window.FindName('BtnFGOFilterClear').IsEnabled = $false
    $script:Window.FindName('TxtFGOFilter').Text       = ''
    $script:FGOAllData = $null
    $script:Window.FindName('BtnFGOExportXlsx').IsEnabled = $false
    Show-Panel 'PanelFindGroupsByOwners'
    Hide-Notification
    Write-VerboseLog "Panel: Find Groups by Owners" -Level Info
})


# ── FIND GROUPS BY OWNERS  -  Run ─────────────────────────────────────────────
$script:Window.FindName('BtnFGORun').Add_Click({
    Hide-Notification

    $inputTxt = $script:Window.FindName('TxtFGOInputList').Text
    $entries  = @($inputTxt -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
    $script:FGONoOwnerMode = ($entries.Count -eq 0)

    # Reset UI
    $script:Window.FindName('PnlFGOResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtFGONoResults').Visibility = 'Collapsed'
    $script:Window.FindName('BtnFGOCopyValue').IsEnabled  = $false
    $script:Window.FindName('BtnFGOCopyRow').IsEnabled    = $false
    $script:Window.FindName('BtnFGOCopyAll').IsEnabled    = $false
    $script:Window.FindName('BtnFGOFilter').IsEnabled  = $false
    $script:Window.FindName('BtnFGOFilterClear').IsEnabled = $false
    $script:Window.FindName('TxtFGOFilter').Text       = ''
    $script:Window.FindName('TxtFGOFilter').IsEnabled  = $false
    $script:Window.FindName('BtnFGOExportXlsx').IsEnabled = $false
    $script:FGOAllData = $null
    $dg = $script:Window.FindName('DgFGOResults')
    if ($dg) { $dg.ItemsSource = $null }

    $script:Window.FindName('PnlFGOProgress').Visibility = 'Visible'
    $script:Window.FindName('TxtFGOProgressMsg').Text    = if ($script:FGONoOwnerMode) { 'Finding groups with no owners...' } else { 'Resolving owners...' }
    $script:Window.FindName('TxtFGOProgressDetail').Text = ''
    $script:Window.FindName('BtnFGORun').IsEnabled       = $false

    $script:FGOQueue = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
    $script:FGOStop  = [hashtable]::Synchronized(@{ Stop = $false })

    $script:FGORs = [runspacefactory]::CreateRunspace()
    $script:FGORs.ApartmentState = 'STA'
    $script:FGORs.Open()
    $script:FGORs.SessionStateProxy.SetVariable('entries',   $entries)
    $script:FGORs.SessionStateProxy.SetVariable('noOwnerMode', $script:FGONoOwnerMode)
    $script:FGORs.SessionStateProxy.SetVariable('queue',     $script:FGOQueue)
    $script:FGORs.SessionStateProxy.SetVariable('stopFlag',  $script:FGOStop)
    # VerboseLog not passed to runspace - uses queue pattern instead

    $script:FGOPs = [powershell]::Create()
    $script:FGOPs.Runspace = $script:FGORs
    $script:FGOPs.AddScript({
        param($entries, $queue, $stopFlag, $noOwnerMode)

        $resultList = [System.Collections.Generic.List[PSObject]]::new()
        $resolved   = 0
        $failed     = 0

        if ($noOwnerMode) {
            # ── No-owner mode: fetch all groups and find those with 0 owners ──
            $queue.Enqueue(@{ type = 'progress'; msg = 'Fetching all groups...'; detail = '' })
            $allGroups = [System.Collections.Generic.List[hashtable]]::new()
            $uri = 'https://graph.microsoft.com/v1.0/groups?$select=id,displayName,groupTypes,mailEnabled,securityEnabled,membershipRule&$top=999&$count=true'
            $hdrs = @{ ConsistencyLevel = 'eventual' }
            try {
                while ($uri) {
                    if ($stopFlag['Stop']) { break }
                    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $hdrs
                    if ($resp['value']) {
                        foreach ($g in $resp['value']) { $allGroups.Add($g) }
                    }
                    $uri = $resp['@odata.nextLink']
                }
            } catch {
                $queue.Enqueue(@{ type = 'progress'; msg = "Error fetching groups: $($_.Exception.Message)"; detail = '' })
            }
            $queue.Enqueue(@{ type = 'progress'; msg = "Checking $($allGroups.Count) group(s) for missing owners..."; detail = '' })
            $checked = 0
            foreach ($g in $allGroups) {
                if ($stopFlag['Stop']) { break }
                $checked++
                if ($checked % 50 -eq 0) {
                    $queue.Enqueue(@{ type = 'progress'; msg = "Checking owners: $checked of $($allGroups.Count)"; detail = '' })
                }
                try {
                    $ownersResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($g['id'])/owners?`$select=id&`$top=1"
                    $ownerCount = @($ownersResp['value']).Count
                    if ($ownerCount -eq 0) {
                        $gTypes = $g['groupTypes']
                        $mailEnabled = $g['mailEnabled']
                        $secEnabled  = $g['securityEnabled']
                        $gType = 'Other'
                        if ($gTypes -and $gTypes -contains 'Unified') { $gType = 'M365' }
                        elseif ($secEnabled -and -not $mailEnabled)   { $gType = 'Security' }
                        elseif ($secEnabled -and $mailEnabled)        { $gType = 'Mail-Enabled Security' }
                        elseif (-not $secEnabled -and $mailEnabled)   { $gType = 'Distribution' }
                        $mType = 'Assigned'
                        if ($gTypes -and $gTypes -contains 'DynamicMembership') { $mType = 'Dynamic' }
                        $memberCount = '...'
                        try {
                            $cntResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($g['id'])/members/`$count" -Headers @{ 'ConsistencyLevel' = 'eventual' }
                            $memberCount = [string]$cntResp
                        } catch { $memberCount = 'N/A' }
                        $resultList.Add([PSCustomObject]@{
                            Owner          = '(none)'
                            GroupName      = [string]$g['displayName']
                            GroupId        = [string]$g['id']
                            GroupType      = $gType
                            MembershipType = $mType
                            MemberCount    = $memberCount
                        })
                    }
                } catch {
                    # Skip groups where owner check fails
                }
            }
            $resolved = $allGroups.Count
            $failed   = 0
        } else {
        foreach ($upn in $entries) {
            if ($stopFlag['Stop']) {
                $queue.Enqueue(@{ type = 'progress'; msg = 'Stopped by user.'; detail = '' })
                break
            }

            $queue.Enqueue(@{ type = 'progress'; msg = "Resolving owner: $upn ..."; detail = "$($resolved + $failed + 1) of $($entries.Count)" })

            try {
                # Resolve user
                $u = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/users/$([Uri]::EscapeDataString($upn))?`$select=id,userPrincipalName"
                $userId = $u['id']
                $ownerUpn = $u['userPrincipalName']

                # Get owned objects (groups only)
                $ownedGroups = [System.Collections.Generic.List[hashtable]]::new()
                $uri = "https://graph.microsoft.com/v1.0/users/$userId/ownedObjects/microsoft.graph.group?`$select=id,displayName,groupTypes,mailEnabled,securityEnabled,membershipRule&`$top=999"
                while ($uri) {
                    if ($stopFlag['Stop']) { break }
                    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
                    if ($resp['value']) {
                        foreach ($g in $resp['value']) { $ownedGroups.Add($g) }
                    }
                    $uri = $resp['@odata.nextLink']
                }

                if ($stopFlag['Stop']) { break }

                foreach ($g in $ownedGroups) {
                    if ($stopFlag['Stop']) { break }

                    # Determine group type
                    $gTypes = $g['groupTypes']
                    $mailEnabled = $g['mailEnabled']
                    $secEnabled  = $g['securityEnabled']
                    $gType = 'Other'
                    if ($gTypes -and $gTypes -contains 'Unified') { $gType = 'M365' }
                    elseif ($secEnabled -and -not $mailEnabled)   { $gType = 'Security' }
                    elseif ($secEnabled -and $mailEnabled)        { $gType = 'Mail-Enabled Security' }
                    elseif (-not $secEnabled -and $mailEnabled)   { $gType = 'Distribution' }

                    # Membership type
                    $mType = 'Assigned'
                    if ($gTypes -and $gTypes -contains 'DynamicMembership') { $mType = 'Dynamic' }

                    # Member count
                    $memberCount = '...'
                    try {
                        $cntResp = Invoke-MgGraphRequest -Method GET `
                            -Uri "https://graph.microsoft.com/v1.0/groups/$($g['id'])/members/`$count" `
                            -Headers @{ 'ConsistencyLevel' = 'eventual' }
                        $memberCount = [string]$cntResp
                    } catch {
                        $memberCount = 'N/A'
                    }

                    $resultList.Add([PSCustomObject]@{
                        Owner          = $ownerUpn
                        GroupName      = [string]$g['displayName']
                        GroupId        = [string]$g['id']
                        GroupType      = $gType
                        MembershipType = $mType
                        MemberCount    = $memberCount
                    })
                }

                $resolved++
                $queue.Enqueue(@{ type = 'progress'; msg = "Resolved: $ownerUpn ($($ownedGroups.Count) group(s))"; detail = "$($resolved + $failed) of $($entries.Count)" })

            } catch {
                $failed++
                $queue.Enqueue(@{ type = 'progress'; msg = "Failed: $upn - $($_.Exception.Message)"; detail = "$($resolved + $failed) of $($entries.Count)" })
            }
        }

        }
        $queue.Enqueue(@{ type = 'done'; results = $resultList; resolved = $resolved; failed = $failed })
    }).AddArgument($entries).AddArgument($script:FGOQueue).AddArgument($script:FGOStop).AddArgument($script:FGONoOwnerMode)

    $script:FGOHandle = $script:FGOPs.BeginInvoke()

    # Timer to poll queue
    $script:FGOTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:FGOTimer.Interval = [TimeSpan]::FromMilliseconds(200)
    $script:FGOTimer.Add_Tick({
        try {
            $msg = $null
            while ($script:FGOQueue.TryDequeue([ref]$msg)) {
                switch ($msg.type) {
                    'progress' {
                        $script:Window.FindName('TxtFGOProgressMsg').Text    = $msg.msg
                        $script:Window.FindName('TxtFGOProgressDetail').Text = $msg.detail
                        Write-VerboseLog "FGO: $($msg.msg)" -Level Info
                    }
                    'done' {
                        $script:FGOTimer.Stop()
                        $script:Window.FindName('PnlFGOProgress').Visibility = 'Collapsed'
                        $script:Window.FindName('BtnFGORun').IsEnabled       = $true
                        $results = $msg.results

                        if ($results -and $results.Count -gt 0) {
                            $dg = $script:Window.FindName('DgFGOResults')
                            $dg.ItemsSource = @($results)
                            $script:Window.FindName('TxtFGOCount').Text = "$($results.Count) owned group(s) found  |  Resolved: $($msg.resolved)  Failed: $($msg.failed)"
                            $script:Window.FindName('PnlFGOResults').Visibility = 'Visible'
                            $script:Window.FindName('BtnFGOCopyAll').IsEnabled    = $true
                            $script:Window.FindName('BtnFGOExportXlsx').IsEnabled = $true
                            $script:Window.FindName('BtnFGOFilter').IsEnabled     = $true
                            $script:Window.FindName('BtnFGOFilterClear').IsEnabled = $true
                            $script:Window.FindName('TxtFGOFilter').IsEnabled     = $true
                            Show-Notification "$($results.Count) owned group(s) found." -BgColor '#D4EDDA' -FgColor '#155724'
                        } else {
                            $script:Window.FindName('TxtFGONoResults').Visibility = 'Visible'
                            Show-Notification "No groups found matching the criteria." -BgColor '#FFF3CD' -FgColor '#856404'
                        }
                        Write-VerboseLog "FGO: Complete. $($results.Count) result(s)." -Level Info
                    }
                }
            }
        } catch {
            Write-VerboseLog "FGO timer error: $($_.Exception.Message)" -Level Warning
        }
    })
    $script:FGOTimer.Start()
})

# ── FIND GROUPS BY OWNERS  -  DataGrid SelectedCellsChanged ───────────────────
$script:Window.FindName('DgFGOResults').Add_SelectedCellsChanged({
    $dg       = $script:Window.FindName('DgFGOResults')
    $selCells = $dg.SelectedCells.Count
    $script:Window.FindName('BtnFGOCopyValue').IsEnabled = ($selCells -ge 1)
    $script:Window.FindName('BtnFGOCopyRow').IsEnabled   = ($selCells -gt 0)
})

# ── FIND GROUPS BY OWNERS  -  Filter ──────────────────────────────────────────
$script:Window.FindName('BtnFGOFilter').Add_Click({
    $dg      = $script:Window.FindName('DgFGOResults')
    $keyword = $script:Window.FindName('TxtFGOFilter').Text.Trim()
    if ($null -eq $script:FGOAllData) {
        $script:FGOAllData = @($dg.ItemsSource)
    }
    if ($keyword -eq '') {
        $dg.ItemsSource = $script:FGOAllData
        Show-Notification "Filter cleared - $($script:FGOAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    } else {
        $filtered = @($script:FGOAllData | Where-Object {
            $row = $_
            $match = $false
            foreach ($prop in $row.PSObject.Properties) {
                if ([string]$prop.Value -like "*$keyword*") { $match = $true; break }
            }
            $match
        })
        $dg.ItemsSource = $filtered
        Show-Notification "$($filtered.Count) of $($script:FGOAllData.Count) result(s) shown after filter." -BgColor '#D4EDDA' -FgColor '#155724'
    }
    Write-VerboseLog "FGO: Filter applied - keyword='$keyword'" -Level Info
})

# ── FIND GROUPS BY OWNERS  -  Filter Clear ────────────────────────────────────
$script:Window.FindName('BtnFGOFilterClear').Add_Click({
    $script:Window.FindName('TxtFGOFilter').Text = ''
    $dg = $script:Window.FindName('DgFGOResults')
    if ($null -ne $script:FGOAllData) {
        $dg.ItemsSource = $script:FGOAllData
        Show-Notification "Filter cleared - $($script:FGOAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    }
})

# ── FIND GROUPS BY OWNERS  -  Copy Value ──────────────────────────────────────
$script:Window.FindName('BtnFGOCopyValue').Add_Click({
    $dg = $script:Window.FindName('DgFGOResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $vals = [System.Collections.Generic.List[string]]::new()
    foreach ($cellInfo in @($dg.SelectedCells)) {
        $col = $cellInfo.Column -as [System.Windows.Controls.DataGridBoundColumn]
        if ($col) { $vals.Add(([string]$cellInfo.Item.($col.Binding.Path.Path)).Trim()) }
    }
    [System.Windows.Clipboard]::SetText($vals -join "`r`n")
    Show-Notification "$($vals.Count) value(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
})

# ── FIND GROUPS BY OWNERS  -  Copy Row ────────────────────────────────────────
$script:Window.FindName('BtnFGOCopyRow').Add_Click({
    $dg = $script:Window.FindName('DgFGOResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $seen = [System.Collections.Generic.HashSet[object]]::new()
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($cell in @($dg.SelectedCells)) { if ($seen.Add($cell.Item)) { $rows.Add($cell.Item) } }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc     = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $header = ($colDefs | ForEach-Object { $_.H }) -join "`t"
    $lines  = [System.Collections.Generic.List[string]]::new()
    $lines.Add($header)
    foreach ($r in $rows) {
        $vals = [System.Collections.Generic.List[string]]::new()
        foreach ($cd in $colDefs) {
            $v = ''
            if ($cd.P -ne '') { $v = [string]$r.($cd.P) }
            $vals.Add($v)
        }
        $lines.Add($vals -join "`t")
    }
    [System.Windows.Clipboard]::SetText($lines -join "`r`n")
    Show-Notification "$($rows.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    $dg.UnselectAllCells()
    foreach ($r in $rows) { $null = $dg.SelectedItems.Add($r) }
})

# ── FIND GROUPS BY OWNERS  -  Copy All ────────────────────────────────────────
$script:Window.FindName('BtnFGOCopyAll').Add_Click({
    $dg = $script:Window.FindName('DgFGOResults')
    if ($null -eq $dg.ItemsSource) { return }
    $allRows = @($dg.ItemsSource)
    if ($allRows.Count -eq 0) { return }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc     = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $header = ($colDefs | ForEach-Object { $_.H }) -join "`t"
    $lines  = [System.Collections.Generic.List[string]]::new()
    $lines.Add($header)
    foreach ($r in $allRows) {
        $vals = [System.Collections.Generic.List[string]]::new()
        foreach ($cd in $colDefs) {
            $v = ''
            if ($cd.P -ne '') { $v = [string]$r.($cd.P) }
            $vals.Add($v)
        }
        $lines.Add($vals -join "`t")
    }
    [System.Windows.Clipboard]::SetText($lines -join "`r`n")
    Show-Notification "$($allRows.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
})

# ── FIND GROUPS BY OWNERS  -  Export XLSX ─────────────────────────────────────
$script:Window.FindName('BtnFGOExportXlsx').Add_Click({
    $dg = $script:Window.FindName('DgFGOResults')
    if ($null -eq $dg.ItemsSource) { return }
    $allRows = @($dg.ItemsSource)
    if ($allRows.Count -eq 0) { return }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title      = 'Save Groups by Owners XLSX'
    $dlg.Filter     = 'Excel files (*.xlsx)|*.xlsx'
    $dlg.DefaultExt = '.xlsx'
    $dlg.FileName   = "GroupsByOwners_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    if (-not $dlg.ShowDialog()) { return }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc     = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    try {
        $exportRows = [System.Collections.Generic.List[PSObject]]::new()
        foreach ($r in $allRows) {
            $obj = [ordered]@{}
            foreach ($cd in $colDefs) {
                $v = ''
                if ($cd.P -ne '') { $v = [string]$r.($cd.P) }
                $obj[$cd.H] = $v
            }
            $exportRows.Add([PSCustomObject]$obj)
        }
        $exportRows | Export-Excel -Path $dlg.FileName -AutoSize -FreezeTopRow -BoldTopRow
        Show-Notification "Exported $($allRows.Count) row(s) to XLSX." -BgColor '#D4EDDA' -FgColor '#155724'
        Write-VerboseLog "FGO: Exported to $($dlg.FileName)" -Level Info
    } catch {
        try {
            $csvPath = [System.IO.Path]::ChangeExtension($dlg.FileName, '.csv')
            $exportRows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Show-Notification "ImportExcel not available - exported as CSV." -BgColor '#FFF3CD' -FgColor '#856404'
        } catch {
            Show-Notification "Export failed: $($_.Exception.Message)" -BgColor '#F8D7DA' -FgColor '#721C24'
        }
    }
})

# ── Set domain suffix label from DevConfig (or just "@" if not configured) ──
$domainSuffix = if ($DevConfig.M365MailDomain) { "@$($DevConfig.M365MailDomain)" } else { "@" }
$script:Window.FindName('TxtMailDomainSuffix').Text = $domainSuffix

# ── M365 panel visibility + lazy sensitivity label fetch ──
$script:Window.FindName('CmbCreateType').Add_SelectionChanged({
    $sel   = $script:Window.FindName('CmbCreateType').SelectedItem.Content
    $isM365 = $sel -like '*365*'
    $script:Window.FindName('PnlMailNick').Visibility = if ($isM365) { 'Visible' } else { 'Collapsed' }
})
$script:Window.FindName('PnlMailNick').Visibility = 'Collapsed'  # default hidden (Security selected)


# ── CREATE GROUP  -  Members ↔ Dynamic Rule mutual exclusion ──
# If the user populates the Members list, grey out the Dynamic Rule field (contradictory inputs).
# Conversely, if a Dynamic Rule is entered, grey out the Members field.
$script:Window.FindName('TxtCreateMembers').Add_TextChanged({
    $hasMembersText = -not [string]::IsNullOrWhiteSpace($script:Window.FindName('TxtCreateMembers').Text)
    $dynBox         = $script:Window.FindName('TxtCreateDynamic')
    $dynBox.IsEnabled = -not $hasMembersText
    $dynBox.Opacity   = if ($hasMembersText) { 0.4 } else { 1.0 }
    $dynBox.ToolTip   = if ($hasMembersText) {
        "Clear the Members field before entering a Dynamic Membership Rule  -  these settings are mutually exclusive."
    } else { $null }
})
$script:Window.FindName('TxtCreateDynamic').Add_TextChanged({
    $hasDynText = -not [string]::IsNullOrWhiteSpace($script:Window.FindName('TxtCreateDynamic').Text)
    $membersBox = $script:Window.FindName('TxtCreateMembers')
    $membersBox.IsEnabled = -not $hasDynText
    $membersBox.Opacity   = if ($hasDynText) { 0.4 } else { 1.0 }
    $membersBox.ToolTip   = if ($hasDynText) {
        "Clear the Dynamic Membership Rule field before adding initial members  -  these settings are mutually exclusive."
    } else { $null }
})


# ── CREATE GROUP  -  Validate & Preview ──
$script:Window.FindName('BtnCreateValidate').Add_Click({
    Hide-Notification
    $groupType  = $script:Window.FindName('CmbCreateType').SelectedItem.Content
    $dispName   = $script:Window.FindName('TxtCreateName').Text.Trim()
    $mailNick   = $script:Window.FindName('TxtCreateMailNick').Text.Trim()
    $desc       = $script:Window.FindName('TxtCreateDesc').Text.Trim()
    $ownerUpnsTxt = $script:Window.FindName('TxtCreateOwner').Text
    $membersTxt = $script:Window.FindName('TxtCreateMembers').Text
    $dynRule    = $script:Window.FindName('TxtCreateDynamic').Text.Trim()
    $isM365     = $groupType -like '*365*'

    # Read sensitivity label selection (M365 only)
    $selLabelItem = $script:Window.FindName('CmbSensitivityLabel').SelectedItem
    $script:CreateSensLabelId   = if ($selLabelItem -and $selLabelItem.Tag)     { $selLabelItem.Tag }     else { '' }
    $script:CreateSensLabelName = if ($selLabelItem -and $selLabelItem.Tag)     { $selLabelItem.Content } else { '' }

    # Read Administrative Unit selection
    $selAUItem = $script:Window.FindName('CmbCreateAU').SelectedItem
    $script:CreateSelectedAU = if ($selAUItem -and $selAUItem.Tag) { @{ Id = $selAUItem.Tag; DisplayName = $selAUItem.Content } } else { $null }

    if ([string]::IsNullOrWhiteSpace($dispName)) {
        Show-Notification "Display name is required." -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }
    if ($isM365 -and [string]::IsNullOrWhiteSpace($mailNick)) {
        Show-Notification "Group email address (mail nickname) is required for M365 groups." -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }

    # Capture all input into script-scope so the background worker and Done callback can access them
    $script:CreateParams = @{
        GroupType     = $groupType;  DispName    = $dispName;   MailNick   = $mailNick
        Desc          = $desc;       OwnerUpnsTxt = $ownerUpnsTxt;   MembersTxt = $membersTxt
        DynRule       = $dynRule;    IsM365      = $isM365
        SensLabelId   = $script:CreateSensLabelId
        SensLabelName = $script:CreateSensLabelName
        MailDomain    = $DevConfig.M365MailDomain
        SelectedAU    = $script:CreateSelectedAU
    }

    $btn = $script:Window.FindName('BtnCreateValidate')

    Invoke-OnBackground -DisableButton $btn -BusyMessage "--- Create Group: Validating ---" -Work {

        $p          = $script:CreateParams
        $dispName   = $p['DispName'];   $ownerUpnsTxt = $p['OwnerUpnsTxt']
        $membersTxt = $p['MembersTxt']; $dynRule      = $p['DynRule']
        $groupType  = $p['GroupType'];  $isM365       = $p['IsM365']
        $mailNick   = $p['MailNick'];   $mailDomain   = $p['MailDomain']

        Write-VerboseLog "Type: $groupType | Name: $dispName" -Level Info

        # Resolve owners (supports multiple UPNs, one per line)
        $ownerObjs     = [System.Collections.Generic.List[PSCustomObject]]::new()
        $invalidOwners = [System.Collections.Generic.List[string]]::new()
        $ownerEntries  = $ownerUpnsTxt -split "`n" | Where-Object { $_.Trim() -ne '' }
        foreach ($upn in $ownerEntries) {
            $upn = $upn.Trim()
            Write-VerboseLog "Resolving owner: $upn" -Level Info
            try {
                $u = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$([Uri]::EscapeDataString($upn))?`$select=id,displayName"
                $ownerObjs.Add([PSCustomObject]@{ Id = $u['id']; DisplayName = $u['displayName']; Type = 'User'; Found = $true; Original = $upn })
                Write-VerboseLog "Owner resolved: $($u['displayName']) ($($u['id']))" -Level Success
            } catch {
                $invalidOwners.Add($upn)
                Write-VerboseLog "Owner UPN not found: $upn" -Level Warning
            }
        }

        # Resolve members
        $validMembers   = [System.Collections.Generic.List[PSCustomObject]]::new()
        $invalidMembers = [System.Collections.Generic.List[string]]::new()
        if (-not [string]::IsNullOrWhiteSpace($membersTxt)) {
            $entries = $membersTxt -split "`n" | Where-Object { $_.Trim() -ne '' }
            $valResult = Validate-InputList -Entries $entries
            foreach ($item in $valResult.Valid)   { $validMembers.Add([PSCustomObject]$item) }
            foreach ($item in $valResult.Invalid) { $invalidMembers.Add([string]$item) }
        }

        # M365: remove device members (not supported) and warn
        if ($isM365) {
            $deviceEntries = @($validMembers | Where-Object { $_.Type -eq 'Device' })
            foreach ($dev in $deviceEntries) {
                $validMembers.Remove($dev) | Out-Null
                $invalidMembers.Add("$($dev.DisplayName)  [skipped - devices are not supported as M365 group members]")
                Write-VerboseLog "Device skipped (M365): $($dev.DisplayName)" -Level Warning
            }
        }

        # Check display name uniqueness (warn only - duplicates are allowed)
        Write-VerboseLog "Checking for duplicate display name..." -Level Info
        $existingNameId = $null
        try {
            $safe_dispName = $dispName -replace "'","''"  # PS5-safe: extracted to avoid nested double-quotes in string
            $resp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe_dispName'`&`$select=id`&`$top=1"
            if ($resp['value'] -and $resp['value'].Count -gt 0) { $existingNameId = $resp['value'][0]['id'] }
        } catch {}
        if ($existingNameId) {
            Write-VerboseLog "WARNING: A group named '$dispName' already exists (ID: $existingNameId)" -Level Warning
        }

        # M365: check mail nickname uniqueness (hard block - email addresses must be unique)
        $emailConflictMsg = $null
        if ($isM365) {
            Write-VerboseLog "Checking mail nickname uniqueness..." -Level Info
            try {
                $nickResp = Invoke-MgGraphRequest -Method GET `
                    $safe_mailNick = $mailNick -replace "'","''"  # PS5-safe: extracted to avoid nested double-quotes in string
                    -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=mailNickname eq '$safe_mailNick'`&`$select=id,displayName`&`$top=1"
                if ($nickResp['value'] -and $nickResp['value'].Count -gt 0) {
                    $conflict = $nickResp['value'][0]
                    $emailAddr = if ($mailDomain) { "$mailNick@$mailDomain" } else { $mailNick }
                    $emailConflictMsg = "Email address '$emailAddr' is already in use by group '$($conflict['displayName'])' (ID: $($conflict['id'])). Please choose a different mail nickname."
                    Write-VerboseLog "Email conflict: $emailConflictMsg" -Level Error
                } else {
                    Write-VerboseLog "Mail nickname is available." -Level Success
                }
            } catch {
                Write-VerboseLog "Could not verify mail nickname uniqueness: $($_.Exception.Message)" -Level Warning
            }
        }

        $script:CreateValidated = @{
            OwnerObjs       = $ownerObjs
            InvalidOwners   = $invalidOwners
            ValidMembers    = $validMembers
            InvalidMembers  = $invalidMembers
            ExistingNameId  = $existingNameId
            EmailConflict   = $emailConflictMsg
        }

    } -Done {

        $p    = $script:CreateParams
        $v    = $script:CreateValidated
        $ownerObjs      = @($v['OwnerObjs'])
        $invalidOwners  = @($v['InvalidOwners'])
        $validMembers   = $v['ValidMembers']
        $invalidMembers = $v['InvalidMembers']
        $existingNameId = $v['ExistingNameId']
        $emailConflict  = $v['EmailConflict']
        $dispName       = $p['DispName'];  $groupType    = $p['GroupType']
        $mailNick       = $p['MailNick'];  $mailDomain   = $p['MailDomain']
        $desc           = $p['Desc'];      $dynRule      = $p['DynRule']
        $isM365         = $p['IsM365'];    $sensLabelName = $p['SensLabelName']

        # Hard block on email conflict
        if ($emailConflict) {
            Show-Notification $emailConflict -BgColor '#F8D7DA' -FgColor '#721C24'
            return
        }

        if ($existingNameId) {
            Show-Notification "Warning: A group named '$dispName' already exists (ID: $existingNameId). Proceeding will create a duplicate display name." -BgColor '#FFF3CD' -FgColor '#7A4800'
        }

        $emailDisplay = if ($isM365 -and $mailDomain) { "$mailNick@$mailDomain" } elseif ($isM365) { $mailNick } else { '' }
        $extra = "Group type  : $groupType"
        if ($existingNameId) { $extra = "$([char]0x26A0) Duplicate name: A group named '$dispName' already exists (ID: $existingNameId).`n" + $extra }
        if ($emailDisplay)   { $extra += "`nEmail       : $emailDisplay" }
        if ($sensLabelName)  { $extra += "`nSens. label : $sensLabelName" }
        $auInfo = $script:CreateParams['SelectedAU']
        if ($auInfo) { $extra += "`nAdmin Unit  : $($auInfo.DisplayName) ($($auInfo.Id))" }
        if ($dynRule)        { $extra += "`nDynamic rule: $dynRule" }
        if ($ownerObjs.Count -gt 0) {
            $ownerLines = ($ownerObjs | ForEach-Object { $_.DisplayName + ' (' + $_.Original + ')' }) -join ', '
            $extra += "`nOwners      : $ownerLines"
        }
        if ($invalidOwners.Count -gt 0) {
            $joined_invalidOwners = $invalidOwners -join ', '  # PS5-safe: extracted to avoid nested double-quotes in string
            $extra += "`nInvalid owner UPNs (will be skipped): $joined_invalidOwners"
        }

        $confirmed = Show-ConfirmationDialog `
            -Title "Create Group" `
            -OperationLabel "Create Group ($groupType)" `
            -TargetGroup $dispName -TargetGroupId "(new)" `
            -ValidObjects $validMembers -InvalidEntries $invalidMembers `
            -ExtraInfo $extra

        if (-not $confirmed) { Write-VerboseLog "Action cancelled by user." -Level Warning; return }

        $btn = $script:Window.FindName('BtnCreateValidate')
        Invoke-OnBackground -DisableButton $btn -BusyMessage "--- Creating group: $dispName ---" -Work {

            $p        = $script:CreateParams
            $v        = $script:CreateValidated
            $dispName    = $p['DispName'];   $mailNick    = $p['MailNick']
            $desc        = $p['Desc'];       $dynRule     = $p['DynRule']
            $isM365      = $p['IsM365'];     $sensLabelId = $p['SensLabelId']
            $ownerObjs   = @($v['OwnerObjs']); $validMembers = $v['ValidMembers']

            $body = @{ displayName = $dispName; securityEnabled = $true; mailEnabled = $isM365 }
            if ($isM365) {
                $body.groupTypes   = @('Unified')
                $body.mailNickname = $mailNick
            } else {
                $body.groupTypes   = @()
                $body.mailNickname = ($dispName -replace '[^a-zA-Z0-9]','') + (Get-Random -Minimum 100 -Maximum 999)
            }
            if ($desc) { $body.description = $desc }
            if ($dynRule) {
                $body.membershipRule = $dynRule
                $body.membershipRuleProcessingState = 'On'
                # M365 dynamic requires both Unified AND DynamicMembership in groupTypes
                $body.groupTypes = if ($isM365) { @('Unified','DynamicMembership') } else { @('DynamicMembership') }
            }
            if ($isM365 -and -not [string]::IsNullOrWhiteSpace($sensLabelId)) {
                $body.assignedLabels = @(@{ labelId = $sensLabelId })
                Write-VerboseLog "Sensitivity label will be applied: $($p['SensLabelName']) ($sensLabelId)" -Level Info
            }

            # ── Route through AU endpoint if Administrative Unit selected ──
            $selectedAU = $p['SelectedAU']
            if ($selectedAU) {
                $body['@odata.type'] = '#microsoft.graph.group'
                $auUri = "https://graph.microsoft.com/v1.0/directory/administrativeUnits/$($selectedAU.Id)/members"
                Write-VerboseLog "Creating group in Administrative Unit: $($selectedAU.DisplayName) ($($selectedAU.Id))" -Level Info
                try {
                    $newGroup = Invoke-MgGraphRequest -Method POST -Uri $auUri `
                        -Body ($body | ConvertTo-Json -Depth 5) -ContentType "application/json"
                    Write-VerboseLog "Group created in AU '$($selectedAU.DisplayName)': $($newGroup['displayName']) ($($newGroup['id']))" -Level Success
                } catch {
                    $auErr = $_.Exception.Message
                    if ($auErr -match '403|Forbidden|Authorization_RequestDenied|Insufficient privileges') {
                        throw "Access denied: You need Groups Administrator or User Administrator role scoped to this Administrative Unit. Error: $auErr"
                    }
                    throw
                }
            } else {
                $newGroup = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" `
                    -Body ($body | ConvertTo-Json -Depth 5) -ContentType "application/json"
                Write-VerboseLog "Group created: $($newGroup['displayName']) ($($newGroup['id']))" -Level Success
            }

            if ($ownerObjs.Count -gt 0) {
                Start-Sleep -Seconds 2
                foreach ($ownerObj in $ownerObjs) {
                    try {
                        $ob = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($ownerObj.Id)" }
                        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups/$($newGroup['id'])/owners/`$ref" `
                            -Body ($ob | ConvertTo-Json) -ContentType "application/json"
                        Write-VerboseLog "Owner set: $($ownerObj.DisplayName)" -Level Success
                     } catch {
                         $errMsg = $_.Exception.Message
                         if ($errMsg -match '400' -or $errMsg -match 'BadRequest' -or $errMsg -match 'already') {
                             Write-VerboseLog "Skipped owner '$($ownerObj.DisplayName)' - already an owner (auto-added as creator)" -Level Warning
                         } else {
                             Write-VerboseLog "Failed to set owner '$($ownerObj.DisplayName)': $errMsg" -Level Error
                         }
                     }
                }
            }

            if ($validMembers.Count -gt 0) {
                Start-Sleep -Seconds 2
                $memResult = Add-MembersToGroup -GroupId $newGroup['id'] -Members $validMembers
                Write-VerboseLog "Members added: $($memResult['Ok']) succeeded, $($memResult['Fail']) failed, $($memResult['Skipped']) skipped" -Level Info
            }

            $script:CreateNewGroupId = $newGroup['id']

        } -Done {
            $dispName = $script:CreateParams['DispName']
            $auMsg = if ($script:CreateParams['SelectedAU']) { " in AU '$($script:CreateParams['SelectedAU'].DisplayName)'" } else { '' }
            Show-Notification "Group '$dispName' created successfully${auMsg} (ID: $($script:CreateNewGroupId))" -BgColor '#D4EDDA' -FgColor '#155724'
        }
    }
})


# ── ADD / REMOVE MEMBERS  -  Validate & Preview ──
$script:Window.FindName('BtnMemberValidate').Add_Click({
    Hide-Notification
    if ($null -eq $script:SelectedGroup) {
        Show-Notification "Please select a target group first." -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }
    $membersTxt = $script:Window.FindName('TxtMemberList').Text
    if ([string]::IsNullOrWhiteSpace($membersTxt)) {
        Show-Notification "Please enter at least one member entry." -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }

    $isAdd  = ($script:CurrentMemberOp -eq 'Add')
    $opLabel = if ($isAdd) { "Add Members" } else { "Remove Members" }
    $script:MemberParams = @{ MembersTxt = $membersTxt; OpLabel = $opLabel; IsAdd = $isAdd }
    $btn = $script:Window.FindName('BtnMemberValidate')

    Invoke-OnBackground -DisableButton $btn -BusyMessage "--- ${opLabel}: Validating ---" -Work {

        $entries = $script:MemberParams['MembersTxt'] -split "`n" | Where-Object { $_.Trim() -ne '' }
        Write-VerboseLog "Target group: $($script:SelectedGroup.DisplayName) ($($script:SelectedGroup.Id))" -Level Info
        $valResult = Validate-InputList -Entries $entries
        $script:MemberValidated = @{ Valid = $valResult.Valid; Invalid = $valResult.Invalid }

    } -Done {

        $valid   = $script:MemberValidated['Valid']
        $invalid = $script:MemberValidated['Invalid']
        $opLabel = $script:MemberParams['OpLabel']

        if ($valid.Count -eq 0) {
            Write-VerboseLog "No valid entries found. Cannot proceed." -Level Error
            Show-Notification "No valid Entra objects found. Correct entries and try again." -BgColor '#F8D7DA' -FgColor '#721C24'
            return
        }
        if ($invalid.Count -gt 0) { Write-VerboseLog "$($invalid.Count) invalid entries will be skipped." -Level Warning }

        $confirmed = Show-ConfirmationDialog `
            -Title $opLabel -OperationLabel $opLabel `
            -TargetGroup $script:SelectedGroup.DisplayName `
            -TargetGroupId $script:SelectedGroup.Id `
            -ValidObjects $valid -InvalidEntries $invalid

        if (-not $confirmed) { Write-VerboseLog "Action cancelled by user." -Level Warning; return }

        $btn = $script:Window.FindName('BtnMemberValidate')
        Invoke-OnBackground -DisableButton $btn -BusyMessage "--- Executing: $opLabel ---" -Work {

            $valid   = $script:MemberValidated['Valid']
            $opLabel = $script:MemberParams['OpLabel']
            $isAdd   = $script:MemberParams['IsAdd']
            Write-VerboseLog "Operation: $(if ($isAdd) { 'ADD' } else { 'REMOVE' })  -  Group: $($script:SelectedGroup.DisplayName)" -Level Action
            if ($isAdd) {
                $memResult = Add-MembersToGroup -GroupId $script:SelectedGroup.Id -Members $valid
                $ok = $memResult['Ok']; $fail = $memResult['Fail']; $skipped = $memResult['Skipped']
            } else {
                $memResult = Remove-MembersFromGroup -GroupId $script:SelectedGroup.Id -Members $valid
                $ok = $memResult['Ok']; $fail = $memResult['Fail']; $skipped = 0
            }
            $script:MemberExecResult = @{ Ok = $ok; Fail = $fail; Skipped = $skipped }

        } -Done {
            $ok      = [int]$script:MemberExecResult['Ok']
            $fail    = [int]$script:MemberExecResult['Fail']
            $skipped = [int]$script:MemberExecResult['Skipped']
            $opLabel = $script:MemberParams['OpLabel']
            $msg = "$opLabel complete. Succeeded: $ok | Skipped (already member): $skipped | Failed: $fail"
            $vLevel  = if ($fail -gt 0) { 'Warning' } else { 'Success' }
            $bgColor = if ($fail -gt 0) { '#FFF3CD' } else { '#D4EDDA' }
            $fgColor = if ($fail -gt 0) { '#7A4800' } else { '#155724' }
            Write-VerboseLog $msg -Level $vLevel
            Show-Notification $msg -BgColor $bgColor -FgColor $fgColor
        }
    }
})


# ── EXPORT MEMBERS ──
$script:Window.FindName('BtnRunExport').Add_Click({
    Hide-Notification
    if ($null -eq $script:ExportSelectedGroup) {
        Show-Notification "Please select a target group first." -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }
    $outPath = $script:Window.FindName('TxtExportPath').Text.Trim()
    if ([string]::IsNullOrWhiteSpace($outPath)) {
        Show-Notification "Please specify an output XLSX path." -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }
    $script:ExportOutPath = $outPath
    $btn = $script:Window.FindName('BtnRunExport')

    Invoke-OnBackground -DisableButton $btn -BusyMessage "--- Export: $($script:ExportSelectedGroup.DisplayName) ---" -Work {
        $count = Export-GroupMembersToXlsx -GroupId $script:ExportSelectedGroup.Id `
            -GroupName $script:ExportSelectedGroup.DisplayName -OutputPath $script:ExportOutPath
        $script:ExportCount = $count
    } -Done {
        $msg = "Exported $($script:ExportCount) members to: $($script:ExportOutPath)"
        Write-VerboseLog $msg -Level Success
        Show-Notification $msg -BgColor '#D4EDDA' -FgColor '#155724'
    }
})


# ── SET DYNAMIC QUERY ──
$script:Window.FindName('BtnDynValidate').Add_Click({
    Hide-Notification
    if ($null -eq $script:DynSelectedGroup) {
        Show-Notification "Please select a target group first." -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }
    $rule = $script:Window.FindName('TxtDynRule').Text.Trim()
    if ([string]::IsNullOrWhiteSpace($rule)) {
        Show-Notification "Please enter a membership rule." -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }
    $script:DynRule = $rule
    $btn = $script:Window.FindName('BtnDynValidate')

    Invoke-OnBackground -DisableButton $btn -BusyMessage "Validating group state and rule syntax..." -Work {
        $rule    = $script:DynRule
        $groupId = $script:DynSelectedGroup.Id
        Write-VerboseLog "Group: $($script:DynSelectedGroup.DisplayName) ($groupId)" -Level Info
        Write-VerboseLog "Rule : $rule" -Level Info

        # ── Re-fetch group to get current groupTypes (avoids stale cached data) ──
        $freshGroupTypes = $null
        try {
            $fg = Invoke-MgGraphRequest -Method GET `
                -Uri "https://graph.microsoft.com/v1.0/groups/$groupId`?`$select=id,groupTypes"
            $freshGroupTypes = @($fg['groupTypes'])
        } catch {
            Write-VerboseLog "Could not re-fetch group type: $($_.Exception.Message) - using cached value." -Level Warning
            $freshGroupTypes = @($script:DynSelectedGroup.GroupTypes)
        }

        $isDynamic   = $freshGroupTypes -contains 'DynamicMembership'
        $memberCount = -1

        if ($isDynamic) {
            # Hard block - already dynamic; no further checks needed
            Write-VerboseLog "BLOCKED: Group is already a dynamic group (groupTypes: $($freshGroupTypes -join ', '))." -Level Warning
        } else {
            # Static group - check member count before allowing conversion
            try {
                $cntResp = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/members/`$count" `
                    -Headers @{ ConsistencyLevel = 'eventual' }
                $memberCount = [int]$cntResp
                Write-VerboseLog "Member count: $memberCount" -Level Info
            } catch {
                Write-VerboseLog "Could not get count via `$count endpoint: $($_.Exception.Message)" -Level Warning
            }
            # Fallback: probe for at least 1 member if primary count endpoint failed
            if ($memberCount -eq -1) {
                try {
                    $mResp = Invoke-MgGraphRequest -Method GET `
                        -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/members?`$select=id`&`$top=1"
                    if ($mResp['value'] -and $mResp['value'].Count -gt 0) {
                        $memberCount = 1
                    } else {
                        $memberCount = 0
                    }
                    Write-VerboseLog "Member count (fallback probe): $memberCount" -Level Info
                } catch {
                    Write-VerboseLog "Could not verify member count - will proceed with warning." -Level Warning
                }
            }

            if ($memberCount -gt 0) {
                Write-VerboseLog "BLOCKED: Static group has $memberCount active member(s) - conversion not allowed." -Level Warning
            } else {
                # Group is eligible - validate rule syntax
                try {
                    $valBody = @{ ruleExpression = $rule }
                    Invoke-MgGraphRequest -Method POST `
                        -Uri "https://graph.microsoft.com/v1.0/groups/validateProperties" `
                        -Body ($valBody | ConvertTo-Json) -ContentType "application/json"
                    Write-VerboseLog "Rule syntax: valid" -Level Success
                } catch {
                    Write-VerboseLog "Rule syntax warning: $($_.Exception.Message) (proceeding)" -Level Warning
                }
            }
        }

        $shared['DynGroupState'] = @{ IsDynamic = $isDynamic; MemberCount = $memberCount }
    } -Done {
        $state = $script:DynGroupState
        if ($null -eq $state) {
            Show-Notification "Validation did not complete - please try again." -BgColor '#F8D7DA' -FgColor '#721C24'
            return
        }

        # ── Hard block: group is already dynamic ──
        if ($state.IsDynamic) {
            Show-Notification "Cannot set dynamic rule: '$($script:DynSelectedGroup.DisplayName)' is already a dynamic group. To modify its existing membership rule, edit the group directly in the Entra admin portal." `
                -BgColor '#F8D7DA' -FgColor '#721C24'
            return
        }

        # ── Hard block: static group has active members ──
        if ($state.MemberCount -gt 0) {
            if ($state.MemberCount -eq 1) {
                $countLabel = "1 active member"
            } else {
                $countLabel = "$($state.MemberCount) active members"
            }
            Show-Notification "Cannot convert '$($script:DynSelectedGroup.DisplayName)' to dynamic membership: the group currently has $countLabel. Converting to dynamic would remove all existing members. This operation is only permitted on static groups with no members." `
                -BgColor '#F8D7DA' -FgColor '#721C24'
            return
        }

        # ── Soft warning: member count could not be verified ──
        if ($state.MemberCount -eq -1) {
            Show-Notification "Warning: could not verify the member count for '$($script:DynSelectedGroup.DisplayName)'. Ensure the group has no members before applying the rule." `
                -BgColor '#FFF3CD' -FgColor '#7A4800'
        }

        # ── Group passes all checks (static, 0 members) - show confirmation ──
        $rule = $script:DynRule
        $emptyValid   = [System.Collections.Generic.List[PSCustomObject]]::new()
        $emptyInvalid = [System.Collections.Generic.List[string]]::new()

        $confirmed = Show-ConfirmationDialog `
            -Title "Set Dynamic Membership Rule" -OperationLabel "Set Dynamic Membership Rule" `
            -TargetGroup $script:DynSelectedGroup.DisplayName `
            -TargetGroupId $script:DynSelectedGroup.Id `
            -ValidObjects $emptyValid -InvalidEntries $emptyInvalid `
            -ExtraInfo "New rule: $rule`n`nWARNING: This will convert the group to dynamic membership.`nAll static members will be removed."

        if (-not $confirmed) { Write-VerboseLog "Action cancelled by user." -Level Warning; return }

        $btn = $script:Window.FindName('BtnDynValidate')
        Invoke-OnBackground -DisableButton $btn -BusyMessage "--- Applying dynamic rule ---" -Work {
            $body = @{
                membershipRule                = $script:DynRule
                membershipRuleProcessingState = "On"
                groupTypes                    = @("DynamicMembership")
            }
            Invoke-MgGraphRequest -Method PATCH `
                -Uri "https://graph.microsoft.com/v1.0/groups/$($script:DynSelectedGroup.Id)" `
                -Body ($body | ConvertTo-Json -Depth 3) -ContentType "application/json"
            Write-VerboseLog "Dynamic rule applied successfully." -Level Success
        } -Done {
            Show-Notification "Dynamic membership rule applied to '$($script:DynSelectedGroup.DisplayName)'" -BgColor '#D4EDDA' -FgColor '#155724'
        }
    }
})


# ── SEARCH ENTRA OBJECTS ──
$script:Window.FindName('BtnEntraSearch').Add_Click({
    Hide-Notification
    $inputTxt = $script:Window.FindName('TxtSearchKeyword').Text
    $entries  = @($inputTxt -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
    if ($entries.Count -eq 0) {
        Show-Notification "Please enter at least one search term." -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }
    $inclUsers   = $script:Window.FindName('ChkSrchUsers').IsChecked
    $inclDevices = $script:Window.FindName('ChkSrchDevices').IsChecked
    $inclSG      = $script:Window.FindName('ChkSrchSG').IsChecked
    $inclM365    = $script:Window.FindName('ChkSrchM365').IsChecked
    $script:SearchExactMatch = [bool]$script:Window.FindName('ChkSrchExactMatch').IsChecked
    $script:SearchGetManager = [bool]$script:Window.FindName('ChkSrchGetManager').IsChecked
    $anyChecked  = $inclUsers -or $inclDevices -or $inclSG -or $inclM365
    $script:SearchEntries = $entries
    $script:SearchTypes   = @{
        Users   = if ($anyChecked) { [bool]$inclUsers   } else { $true }
        Devices = if ($anyChecked) { [bool]$inclDevices } else { $true }
        SG      = if ($anyChecked) { [bool]$inclSG      } else { $true }
        M365    = if ($anyChecked) { [bool]$inclM365    } else { $true }
    }
    $script:Window.FindName('PnlSearchResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtSearchNoResults').Visibility = 'Collapsed'
    $btn = $script:Window.FindName('BtnEntraSearch')
    Invoke-OnBackground -DisableButton $btn -BusyMessage "Searching Entra objects ($($entries.Count) term(s))..." -Work {
        $allEntries = $script:SearchEntries
        $types      = $script:SearchTypes
        $exactMatch = $script:SearchExactMatch
        $hdrs       = @{ ConsistencyLevel = 'eventual' }
        $results    = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($kw in $allEntries) {
            Write-VerboseLog "Search term: '$kw'" -Level Info

            # Users
            if ($types['Users']) {
                try {
                    $uri = "https://graph.microsoft.com/v1.0/users?`$search=`"displayName:$kw`" OR `"userPrincipalName:$kw`"`&`$select=id,displayName,userPrincipalName,mail`&`$count=true`&`$top=999"
                    $userCount = 0
                    do {
                        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $hdrs
                        foreach ($u in @($resp['value'])) {
                            $det = if ($u['userPrincipalName']) { [string]$u['userPrincipalName'] } elseif ($u['mail']) { [string]$u['mail'] } else { '' }
                            $results.Add([PSCustomObject]@{ DisplayName = [string]$u['displayName']; Detail = $det; ObjectType = 'User'; ObjectId = [string]$u['id'] })
                        }
                        $userCount += @($resp['value']).Count
                        $uri = if ($resp['@odata.nextLink']) { [string]$resp['@odata.nextLink'] } else { $null }
                    } while ($uri)
                    Write-VerboseLog "Users found for '$kw': $userCount" -Level Info
                } catch { Write-VerboseLog "User search error for '$kw': $($_.Exception.Message)" -Level Warning }
            }

            # Devices
            if ($types['Devices']) {
                try {
                    $uri = "https://graph.microsoft.com/v1.0/devices?`$search=`"displayName:$kw`"`&`$select=id,displayName,operatingSystem,operatingSystemVersion,deviceId`&`$count=true`&`$top=999"
                    $devCount = 0
                    do {
                        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $hdrs
                        foreach ($d in @($resp['value'])) {
                            $det = [string]$d['deviceId']
                            $results.Add([PSCustomObject]@{ DisplayName = [string]$d['displayName']; Detail = $det; ObjectType = 'Device'; ObjectId = [string]$d['id'] })
                        }
                        $devCount += @($resp['value']).Count
                        $uri = if ($resp['@odata.nextLink']) { [string]$resp['@odata.nextLink'] } else { $null }
                    } while ($uri)
                    Write-VerboseLog "Devices found for '$kw': $devCount" -Level Info
                } catch { Write-VerboseLog "Device search error for '$kw': $($_.Exception.Message)" -Level Warning }
            }

            # Groups
            if ($types['SG'] -or $types['M365']) {
                try {
                    $uri = "https://graph.microsoft.com/v1.0/groups?`$search=`"displayName:$kw`"`&`$select=id,displayName,groupTypes,mail`&`$count=true`&`$top=999"
                    $grpCount = 0
                    do {
                        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $hdrs
                        foreach ($g in @($resp['value'])) {
                            $gTypes = @($g['groupTypes'])
                            $isM365 = $gTypes -contains 'Unified'
                            $isDyn  = $gTypes -contains 'DynamicMembership'
                            if ($isM365 -and $types['M365']) {
                                $results.Add([PSCustomObject]@{ DisplayName = [string]$g['displayName']; Detail = [string]$g['id']; ObjectType = 'M365 Group'; ObjectId = [string]$g['id'] })
                            } elseif (-not $isM365 -and $types['SG']) {
                                $tl = if ($isDyn) { 'Security Group (Dynamic)' } else { 'Security Group' }
                                $results.Add([PSCustomObject]@{ DisplayName = [string]$g['displayName']; Detail = [string]$g['id']; ObjectType = $tl; ObjectId = [string]$g['id'] })
                            }
                        }
                        $grpCount += @($resp['value']).Count
                        $uri = if ($resp['@odata.nextLink']) { [string]$resp['@odata.nextLink'] } else { $null }
                    } while ($uri)
                    Write-VerboseLog "Groups found for '$kw': $grpCount" -Level Info
                } catch { Write-VerboseLog "Group search error for '$kw': $($_.Exception.Message)" -Level Warning }
            }
        }

        # De-duplicate by ObjectId (same object may appear from multiple search terms)
        $seen = [System.Collections.Generic.HashSet[string]]::new()
        $deduped = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($r in $results) {
            if ($seen.Add($r.ObjectId)) { $deduped.Add($r) }
        }

        # Client-side exact-match filter when checkbox is checked
        if ($exactMatch) {
            $allKw = $script:SearchEntries
            $deduped = [System.Collections.Generic.List[PSCustomObject]]($deduped | Where-Object {
                $dn = $_.DisplayName
                $matched = $false
                foreach ($k in $allKw) { if ($dn -ieq $k) { $matched = $true; break } }
                $matched
            })
        }

        Write-VerboseLog "Search total (deduped): $($deduped.Count)" -Level Info

        # Get Manager UPN for User objects (if checkbox checked)
        $getManager = $script:SearchGetManager
        if ($getManager) {
            $finalList = [System.Collections.Generic.List[PSCustomObject]]::new()
            $userCount = ($deduped | Where-Object { $_.ObjectType -eq 'User' }).Count
            $uIdx = 0
            foreach ($item in $deduped) {
                $mgrUpn = ''
                if ($item.ObjectType -eq 'User') {
                    $uIdx++
                    try {
                        $mUri = "https://graph.microsoft.com/v1.0/users/$($item.ObjectId)/manager?`$select=userPrincipalName,displayName"
                        $mResp = Invoke-MgGraphRequest -Method GET -Uri $mUri
                        $mgrUpn = if ($mResp['userPrincipalName']) { [string]$mResp['userPrincipalName'] } else { '(no manager)' }
                    } catch {
                        if ($_.Exception.Message -match '404' -or $_.Exception.Message -match 'Resource .* not found') {
                            $mgrUpn = '(no manager)'
                        } else {
                            $mgrUpn = '(error)'
                            Write-VerboseLog "Manager lookup error for $($item.DisplayName): $($_.Exception.Message)" -Level Warning
                        }
                    }
                }
                $finalList.Add([PSCustomObject]@{
                    DisplayName = $item.DisplayName
                    Detail      = $item.Detail
                    Manager     = $mgrUpn
                    ObjectType  = $item.ObjectType
                    ObjectId    = $item.ObjectId
                })
            }
            $deduped = $finalList
            Write-VerboseLog "Manager lookup completed for $userCount user(s) out of $($finalList.Count) total object(s)" -Level Info
        }

        $shared['SearchResults'] = $deduped
    } -Done {
        $results = $script:SearchResults
        if ($null -eq $results) {
            Show-Notification "Search did not complete - please try again." -BgColor '#F8D7DA' -FgColor '#721C24'
            return
        }
        if ($results.Count -eq 0) {
            $script:Window.FindName('TxtSearchNoResults').Visibility = 'Visible'
            $script:Window.FindName('PnlSearchResults').Visibility   = 'Collapsed'
            return
        }
        $dg = $script:Window.FindName('DgSearchResults')
        $dg.ItemsSource = $results
        # Dynamic Manager column: remove old one if present, add if GetManager checked
        $existingMgrCol = $dg.Columns | Where-Object { $_.Header -eq "Manager" }
        if ($existingMgrCol) { $dg.Columns.Remove($existingMgrCol) }
        if ($script:SearchGetManager) {
            $mgrCol = [System.Windows.Controls.DataGridTextColumn]::new()
            $mgrCol.Header  = "Manager"
            $mgrCol.Binding = [System.Windows.Data.Binding]::new("Manager")
            $mgrCol.Width   = [System.Windows.Controls.DataGridLength]::new(220)
            # Insert after Details (index 1) so order is: Object Name | Details | Manager | Object Type
            $dg.Columns.Insert(2, $mgrCol)
        }
        $dg.UnselectAll()
        $n = $results.Count
        $script:Window.FindName('TxtSearchCount').Text = "$n result$(if ($n -ne 1) { 's' } else { '' })"
        $script:Window.FindName('BtnSearchCopyValue').IsEnabled  = $false
        $script:Window.FindName('BtnSearchCopyRow').IsEnabled    = $false
        $script:Window.FindName('BtnSearchCopyAll').IsEnabled    = $true
        $script:Window.FindName('TxtSearchFilter').IsEnabled  = $true
        $script:Window.FindName('BtnSearchFilter').IsEnabled  = $true
        $script:Window.FindName('BtnSearchFilterClear').IsEnabled = $true
        $script:Window.FindName('BtnSearchExportXlsx').IsEnabled = $true
        $script:Window.FindName('PnlSearchResults').Visibility   = 'Visible'
        $script:Window.FindName('TxtSearchNoResults').Visibility = 'Collapsed'
        Write-VerboseLog "Search returned $n result(s)." -Level Success
    }
})

# DgSearchResults: update Copy Value / Copy Row buttons on CELL selection change
# Note: SelectionUnit="CellOrRowHeader" allows cell selection by default; Copy Row switches to full-row selection
$script:Window.FindName('DgSearchResults').Add_SelectedCellsChanged({
    $dg       = $script:Window.FindName('DgSearchResults')
    $selCells = $dg.SelectedCells.Count
    $script:Window.FindName('BtnSearchCopyValue').IsEnabled = ($selCells -ge 1)
    $script:Window.FindName('BtnSearchCopyRow').IsEnabled   = ($selCells -gt 0)
})

# BtnSearchCopyValue: copy exact value of the single selected cell
# ── BtnSearchFilter ──
$script:Window.FindName('BtnSearchFilter').Add_Click({
    $dg      = $script:Window.FindName('DgSearchResults')
    $keyword = $script:Window.FindName('TxtSearchFilter').Text.Trim()
    if ($null -eq $script:SearchAllData) {
        $script:SearchAllData = @($dg.ItemsSource)
    }
    if ($keyword -eq '') {
        $dg.ItemsSource = $script:SearchAllData
        Show-Notification "Filter cleared - $($script:SearchAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    } else {
        $filtered = @($script:SearchAllData | Where-Object {
            $row = $_
            $match = $false
            foreach ($prop in $row.PSObject.Properties) {
                if ([string]$prop.Value -like "*$keyword*") { $match = $true; break }
            }
            $match
        })

# ── BtnSearchFilterClear ──
$script:Window.FindName('BtnSearchFilterClear').Add_Click({
    $script:Window.FindName('TxtSearchFilter').Text = ''
    $dg = $script:Window.FindName('DgSearchResults')
    if ($null -ne $script:SearchAllData) {
        $dg.ItemsSource = $script:SearchAllData
        Show-Notification "Filter cleared - $($script:SearchAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    }
})
        $dg.ItemsSource = $filtered
        Show-Notification "$($filtered.Count) of $($script:SearchAllData.Count) result(s) shown after filter." -BgColor '#D4EDDA' -FgColor '#155724'
    }
    Write-VerboseLog "Search: Filter applied - keyword='$keyword'" -Level Info
})

$script:Window.FindName('BtnSearchCopyValue').Add_Click({
    $dg = $script:Window.FindName('DgSearchResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $vals = [System.Collections.Generic.List[string]]::new()
    foreach ($cellInfo in @($dg.SelectedCells)) {
        $col = $cellInfo.Column -as [System.Windows.Controls.DataGridBoundColumn]
        if ($col) { $vals.Add(([string]$cellInfo.Item.($col.Binding.Path.Path)).Trim()) }
    }
    [System.Windows.Clipboard]::SetText($vals -join "`r`n")
    Show-Notification "$($vals.Count) value(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "Copied cell value to clipboard." -Level Success
})

# BtnSearchCopyRow: copy full row(s) with headers for every row containing a selected cell
$script:Window.FindName('BtnSearchCopyRow').Add_Click({
    $dg = $script:Window.FindName('DgSearchResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $seen = [System.Collections.Generic.HashSet[object]]::new()
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($cell in @($dg.SelectedCells)) { if ($seen.Add($cell.Item)) { $rows.Add($cell.Item) } }
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("Object Name`tDetails`tObject Type")
    foreach ($row in $rows) { $lines.Add("$($row.DisplayName)`t$($row.Detail)`t$($row.ObjectType)") }
    [System.Windows.Clipboard]::SetText($lines -join "`r`n")
    Show-Notification "$($rows.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "Copied $($rows.Count) row(s) to clipboard." -Level Success
    # Highlight full rows
    $dg.UnselectAllCells()
    foreach ($r in $rows) { $null = $dg.SelectedItems.Add($r) }
})

# BtnSearchCopyAll: copy all result rows with headers
$script:Window.FindName('BtnSearchCopyAll').Add_Click({
    $dg = $script:Window.FindName('DgSearchResults')
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) { return }
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("Object Name`tDetails`tObject Type")
    foreach ($row in $allItems) { $lines.Add("$($row.DisplayName)`t$($row.Detail)`t$($row.ObjectType)") }
    [System.Windows.Clipboard]::SetText($lines -join "`r`n")
    Show-Notification "$($allItems.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "Copied all $($allItems.Count) result row(s) to clipboard." -Level Success
})

# BtnSearchExportXlsx: export all search results via SaveFileDialog
$script:Window.FindName('BtnSearchExportXlsx').Add_Click({
    $dg = $script:Window.FindName('DgSearchResults')
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) {
        Show-Notification "No results to export." -BgColor '#FFF3CD' -FgColor '#856404'
        return
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title    = "Export Search Results"
    $dlg.Filter   = "Excel Workbook (*.xlsx)|*.xlsx|CSV File (*.csv)|*.csv|All Files (*.*)|*.*"
    $dlg.FileName = "EntraSearch_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    $initDir = try { if ($script:ExportOutPath) { Split-Path $script:ExportOutPath -Parent } else { $null } } catch { $null }
    $dlg.InitialDirectory = if ($initDir -and (Test-Path $initDir)) { $initDir } else { [Environment]::GetFolderPath('Desktop') }
    if ($dlg.ShowDialog() -ne $true) { return }
    $outPath = $dlg.FileName
    try {
        $folder = Split-Path $outPath -Parent
        if ($folder -and -not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        $exportRows = $allItems | Select-Object DisplayName, Detail, ObjectType
        if (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue) {
            Import-Module ImportExcel -ErrorAction SilentlyContinue
            $exportRows | Export-Excel -Path $outPath -WorksheetName "SearchResults" -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium2
            Show-Notification "Exported $($allItems.Count) rows: $outPath" -BgColor '#D4EDDA' -FgColor '#155724'
        } else {
            $exportRows | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8
            Show-Notification "ImportExcel not found - saved as CSV: $(Split-Path $outPath -Leaf)" -BgColor '#FFF3CD' -FgColor '#856404'
        }
        Write-VerboseLog "Exported $($allItems.Count) search results to: $outPath" -Level Success
    } catch {
        Show-Notification "Export failed: $($_.Exception.Message)" -BgColor '#F8D7DA' -FgColor '#721C24'
        Write-VerboseLog "Search export error: $($_.Exception.Message)" -Level Error
    }
})



# ════════════════════════════════════════════════════════════════════════════════
# LIST GROUP MEMBERS — helper functions and handlers
# ════════════════════════════════════════════════════════════════════════════════

function Set-LMColumns {
    param([string]$Schema, [System.Windows.Controls.DataGrid]$DataGrid)
    $DataGrid.Columns.Clear()
    $colDefs = switch ($Schema) {
        'Users'   { @(
            @{H='Object Name';         B='ObjectName'},
            @{H='Object ID';           B='ObjectID'},
            @{H='Object Type';         B='ObjectType'},
            @{H='UPN';                 B='UPN'},
            @{H='Member';           B='MemberOf'}
        )}
        'Devices' { @(
            @{H='Device Name';         B='ObjectName'},
            @{H='Device ObjectID';     B='DeviceObjectID'},
            @{H='Entra Device ID';     B='EntraDeviceID'},
            @{H='OS Platform';         B='OSPlatform'},
            @{H='OS Version';          B='OSVersion'},
            @{H='Ownership Type';      B='OwnershipType'},
            @{H='Registered User UPN'; B='RegisteredUserUPN'},
            @{H='Member';           B='MemberOf'}
        )}
        'Groups'  { @(
            @{H='Group Name';      B='ObjectName'},
            @{H='Group Type';      B='GroupType'},
            @{H='Group ID';        B='GroupID'},
            @{H='Membership Type'; B='MembershipType'},
            @{H='Member';       B='MemberOf'}
        )}
        default   { @(
            @{H='Object Name'; B='ObjectName'},
            @{H='Object Type'; B='ObjectType'},
            @{H='Object ID';   B='ObjectID'},
            @{H='Member';   B='MemberOf'}
        )}
    }
    foreach ($def in $colDefs) {
        $col            = [System.Windows.Controls.DataGridTextColumn]::new()
        $col.Header     = $def['H']
        $col.Binding    = [System.Windows.Data.Binding]::new($def['B'])
        $col.IsReadOnly = $true
        $col.Width      = [System.Windows.Controls.DataGridLength]::new(
            1, [System.Windows.Controls.DataGridLengthUnitType]::Star)
        $DataGrid.Columns.Add($col)
    }
}

function Apply-LMFilter {
    if (-not $script:LMAllResults -or $script:LMAllResults.Count -eq 0) { return }
    $showUsers   = $script:Window.FindName('ChkLMUsers').IsChecked
    $showDevices = $script:Window.FindName('ChkLMDevices').IsChecked
    $showGroups  = $script:Window.FindName('ChkLMGroups').IsChecked
    $allUnchecked = (-not $showUsers) -and (-not $showDevices) -and (-not $showGroups)
    $filtered = @($script:LMAllResults | Where-Object {
        $allUnchecked -or
        ($showUsers   -and $_.ObjectType -eq 'User')   -or
        ($showDevices -and $_.ObjectType -eq 'Device') -or
        ($showGroups  -and $_.ObjectType -eq 'Group')
    })
    $dg = $script:Window.FindName('DgLMResults')
    if ($filtered.Count -eq 0) {
        $dg.ItemsSource = $null
        $script:Window.FindName('PnlLMResults').Visibility   = 'Collapsed'
        $script:Window.FindName('TxtLMNoResults').Visibility = 'Visible'
        return
    }
    $types  = @($filtered | Select-Object -ExpandProperty ObjectType -Unique)
    $schema = if ($types.Count -eq 1) {
        switch ($types[0]) { 'User'{'Users'} 'Device'{'Devices'} 'Group'{'Groups'} default{'Mixed'} }
    } else { 'Mixed' }
    Set-LMColumns -Schema $schema -DataGrid $dg
    $dg.ItemsSource = $filtered
    $script:Window.FindName('TxtLMCount').Text           = "$($filtered.Count) member(s) $([char]0x2014) schema: $($schema)"
    $script:Window.FindName('TxtLMNoResults').Visibility = 'Collapsed'
    $script:Window.FindName('PnlLMResults').Visibility   = 'Visible'
    $script:Window.FindName('BtnLMCopyAll').IsEnabled    = $true
    $script:Window.FindName('TxtLMFilter').IsEnabled  = $true
    $script:Window.FindName('BtnLMFilter').IsEnabled  = $true
    $script:Window.FindName('BtnLMFilterClear').IsEnabled = $true
    $script:Window.FindName('BtnLMExportXlsx').IsEnabled = $true
    $script:Window.FindName('BtnLMCopyValue').IsEnabled  = $false
    $script:Window.FindName('BtnLMCopyRow').IsEnabled    = $false
}

# ── BtnLMQuery: List Group Members ──────────────────────────────────────────
$script:Window.FindName('BtnLMQuery').Add_Click({
    Hide-Notification
    $groupTxt = $script:Window.FindName('TxtLMGroupInput').Text
    if ([string]::IsNullOrWhiteSpace($groupTxt)) {
        Show-Notification "Please enter at least one group name or Object ID." -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }
    $script:LMGroupInput = $groupTxt
    $script:LMAllResults = $null
    $script:Window.FindName('PnlLMResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtLMNoResults').Visibility = 'Collapsed'
    $script:Window.FindName('PnlLMProgress').Visibility  = 'Visible'
    $script:Window.FindName('PbLMProgress').IsIndeterminate = $true
    $script:Window.FindName('TxtLMProgressMsg').Text    = 'Querying groups...'
    $script:Window.FindName('TxtLMProgressDetail').Text = ''

    # Start progress update timer (reads SharedBg written by background runspace)
    if ($script:LMProgressTimer) { $script:LMProgressTimer.Stop() }
    $script:LMProgressTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:LMProgressTimer.Interval = [TimeSpan]::FromMilliseconds(400)
    $script:LMProgressTimer.Add_Tick({
        $prog = if ($script:SharedBg) { $script:SharedBg['LMProgress'] } else { $null }
        if ($prog) {
            $script:Window.FindName('TxtLMProgressMsg').Text    = [string]$prog['Message']
            $script:Window.FindName('TxtLMProgressDetail').Text = [string]$prog['Detail']
        }
    })
    $script:LMProgressTimer.Start()

    $btn = $script:Window.FindName('BtnLMQuery')
    Invoke-OnBackground -DisableButton $btn -BusyMessage "--- List Group Members: Querying ---" -Work {

        $script:LMGroupInput = $Shared['LMGroupInput']
        $groupEntries = $script:LMGroupInput -split "`n" | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() }
        $totalGroups  = $groupEntries.Count
        $allResults   = [System.Collections.Generic.List[PSCustomObject]]::new()

        $Shared['LMProgress'] = @{ Message = "Starting..."; Detail = "0 groups queued" }

        $gi = 0
        foreach ($entry in $groupEntries) {
            $gi++
            $Shared['LMProgress'] = @{ Message = "Resolving group $gi of $($totalGroups): $entry"; Detail = "$($allResults.Count) members so far" }

            # Resolve: GUID → direct GET, otherwise search by displayName
            $groupId   = $null
            $groupName = $entry
            if ($entry -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                $groupId = $entry
                try {
                    $gInfo     = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$groupId?`$select=displayName"
                    $groupName = [string]$gInfo['displayName']
                } catch { Write-VerboseLog "Could not fetch display name for ID $entry" -Level Warning }
            } else {
                try {
                    $enc   = [Uri]::EscapeDataString($entry)
                    $safe_entry = $entry -replace "'","''"  # PS5-safe: extracted to avoid nested double-quotes in string
                    $gResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe_entry'`&`$select=id,displayName`&`$top=1"
                    if ($gResp['value'] -and $gResp['value'].Count -gt 0) {
                        $groupId   = [string]$gResp['value'][0]['id']
                        $groupName = [string]$gResp['value'][0]['displayName']
                    } else {
                        Write-VerboseLog "Group not found: $entry" -Level Warning
                        continue
                    }
                } catch { Write-VerboseLog "Error resolving group '$entry': $($_.Exception.Message)" -Level Error; continue }
            }
            Write-VerboseLog ('Querying members of: ' + $groupName + ' (' + $groupId + ')') -Level Info
            $Shared['LMProgress'] = @{ Message = "Querying group $gi of $($totalGroups): $groupName"; Detail = "$($allResults.Count) members so far" }

            # ── Users ──────────────────────────────────────────────────────────
            try {
                $uri = "https://graph.microsoft.com/v1.0/groups/$groupId/members/microsoft.graph.user?`$select=id,displayName,userPrincipalName`&`$top=999"
                do {
                    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
                    foreach ($u in @($resp['value'])) {
                        $allResults.Add([PSCustomObject]@{
                            ObjectName        = [string]$u['userPrincipalName']
                            ObjectType        = 'User'
                            ObjectID          = [string]$u['id']
                            MemberOf          = $groupName
                            UPN               = [string]$u['userPrincipalName']
                            DeviceObjectID    = ''
                            EntraDeviceID     = ''
                            OSPlatform        = ''
                            OSVersion         = ''
                            OwnershipType     = ''
                            RegisteredUserUPN = ''
                            GroupType         = ''
                            GroupID           = ''
                            MembershipType    = ''
                        })
                    }
                    $Shared['LMProgress'] = @{ Message = "Group $gi/$($totalGroups): $groupName  (users)"; Detail = "$($allResults.Count) members so far" }
                    $uri = if ($resp['@odata.nextLink']) { [string]$resp['@odata.nextLink'] } else { $null }
                } while ($uri)
            } catch { Write-VerboseLog "Error querying users from '$groupName': $($_.Exception.Message)" -Level Warning }

            # ── Devices ────────────────────────────────────────────────────────
            try {
                $uri = "https://graph.microsoft.com/v1.0/groups/$groupId/members/microsoft.graph.device?`$select=id,displayName,deviceId,operatingSystem,operatingSystemVersion,deviceOwnership`&`$top=999"
                do {
                    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
                    foreach ($d in @($resp['value'])) {
                        $allResults.Add([PSCustomObject]@{
                            ObjectName        = [string]$d['displayName']
                            ObjectType        = 'Device'
                            ObjectID          = [string]$d['id']
                            MemberOf          = $groupName
                            UPN               = ''
                            DeviceObjectID    = [string]$d['id']
                            EntraDeviceID     = [string]$d['deviceId']
                            OSPlatform        = [string]$d['operatingSystem']
                            OSVersion         = [string]$d['operatingSystemVersion']
                            OwnershipType     = [string]$d['deviceOwnership']
                            RegisteredUserUPN = ''
                            GroupType         = ''
                            GroupID           = ''
                            MembershipType    = ''
                        })
                    }
                    $Shared['LMProgress'] = @{ Message = "Group $gi/$($totalGroups): $groupName  (devices)"; Detail = "$($allResults.Count) members so far" }
                    $uri = if ($resp['@odata.nextLink']) { [string]$resp['@odata.nextLink'] } else { $null }
                } while ($uri)
            } catch { Write-VerboseLog "Error querying devices from '$groupName': $($_.Exception.Message)" -Level Warning }

            # ── Nested groups ──────────────────────────────────────────────────
            try {
                $uri = "https://graph.microsoft.com/v1.0/groups/$groupId/members/microsoft.graph.group?`$select=id,displayName,groupTypes,membershipRuleProcessingState`&`$top=999"
                do {
                    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
                    foreach ($g in @($resp['value'])) {
                        $gt = if ($g['groupTypes'] -and ($g['groupTypes'] -contains 'Unified')) { 'Microsoft 365' } else { 'Security' }
                        $mt = if ($g['membershipRuleProcessingState'] -eq 'On') { 'Dynamic' } else { 'Static' }
                        $allResults.Add([PSCustomObject]@{
                            ObjectName        = [string]$g['displayName']
                            ObjectType        = 'Group'
                            ObjectID          = [string]$g['id']
                            MemberOf          = $groupName
                            UPN               = ''
                            DeviceObjectID    = ''
                            EntraDeviceID     = ''
                            OSPlatform        = ''
                            OSVersion         = ''
                            OwnershipType     = ''
                            RegisteredUserUPN = ''
                            GroupType         = $gt
                            GroupID           = [string]$g['id']
                            MembershipType    = $mt
                        })
                    }
                    $Shared['LMProgress'] = @{ Message = "Group $gi/$($totalGroups): $groupName  (nested groups)"; Detail = "$($allResults.Count) members so far" }
                    $uri = if ($resp['@odata.nextLink']) { [string]$resp['@odata.nextLink'] } else { $null }
                } while ($uri)
            } catch { Write-VerboseLog "Error querying nested groups from '$groupName': $($_.Exception.Message)" -Level Warning }
        }

        # If result set is devices-only → fetch RegisteredUser UPN per device
        $deviceRows    = @($allResults | Where-Object { $_.ObjectType -eq 'Device' })
        $nonDeviceRows = @($allResults | Where-Object { $_.ObjectType -ne 'Device' })
        if ($nonDeviceRows.Count -eq 0 -and $deviceRows.Count -gt 0) {
            Write-VerboseLog "Devices-only result set — fetching registered user UPNs..." -Level Info
            $di = 0
            foreach ($dev in $deviceRows) {
                $di++
                $Shared['LMProgress'] = @{ Message = ('Fetching registered users (' + $di + ' / ' + $deviceRows.Count + ')...'); Detail = ([string]$deviceRows.Count + ' devices total') }
                try {
                    $rResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/devices/$($dev.DeviceObjectID)/registeredUsers?`$select=userPrincipalName`&`$top=1"
                    if ($rResp['value'] -and $rResp['value'].Count -gt 0) {
                        $dev.RegisteredUserUPN = [string]$rResp['value'][0]['userPrincipalName']
                    }
                } catch { Write-VerboseLog "Could not get registered user for $($dev.ObjectName): $($_.Exception.Message)" -Level Warning }
            }
        }

        $Shared['LMProgress'] = @{ Message = "Processing complete."; Detail = "$($allResults.Count) member(s) collected" }
        Write-VerboseLog "LM query complete: $($allResults.Count) member(s) from $gi group(s)" -Level Success
        $script:LMAllResults = $allResults

    } -Done {
        if ($script:LMProgressTimer) { $script:LMProgressTimer.Stop() }
        $script:Window.FindName('PnlLMProgress').Visibility = 'Collapsed'

        if (-not $script:LMAllResults -or $script:LMAllResults.Count -eq 0) {
            $script:Window.FindName('PnlLMResults').Visibility   = 'Collapsed'
            $script:Window.FindName('TxtLMNoResults').Visibility = 'Visible'
            Show-Notification "No members found for the specified group(s)." -BgColor '#FFF3CD' -FgColor '#856404'
            return
        }
        Apply-LMFilter
    }
})

# ── DgLMResults: cell selection enables Copy Value / Copy Row ────────────────
$script:Window.FindName('DgLMResults').Add_SelectedCellsChanged({
    $dg       = $script:Window.FindName('DgLMResults')
    $selCells = $dg.SelectedCells.Count
    $script:Window.FindName('BtnLMCopyValue').IsEnabled = ($selCells -ge 1)
    $script:Window.FindName('BtnLMCopyRow').IsEnabled   = ($selCells -gt 0)
})

# ── BtnLMCopyValue ───────────────────────────────────────────────────────────
# ── BtnLMFilter ──
$script:Window.FindName('BtnLMFilter').Add_Click({
    $dg      = $script:Window.FindName('DgLMResults')
    $keyword = $script:Window.FindName('TxtLMFilter').Text.Trim()
    if ($null -eq $script:LMAllData) {
        $script:LMAllData = @($dg.ItemsSource)
    }
    if ($keyword -eq '') {
        $dg.ItemsSource = $script:LMAllData
        Show-Notification "Filter cleared - $($script:LMAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    } else {
        $filtered = @($script:LMAllData | Where-Object {
            $row = $_
            $match = $false
            foreach ($prop in $row.PSObject.Properties) {
                if ([string]$prop.Value -like "*$keyword*") { $match = $true; break }
            }
            $match
        })

# ── BtnLMFilterClear ──
$script:Window.FindName('BtnLMFilterClear').Add_Click({
    $script:Window.FindName('TxtLMFilter').Text = ''
    $dg = $script:Window.FindName('DgLMResults')
    if ($null -ne $script:LMAllData) {
        $dg.ItemsSource = $script:LMAllData
        Show-Notification "Filter cleared - $($script:LMAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    }
})
        $dg.ItemsSource = $filtered
        Show-Notification "$($filtered.Count) of $($script:LMAllData.Count) result(s) shown after filter." -BgColor '#D4EDDA' -FgColor '#155724'
    }
    Write-VerboseLog "LM: Filter applied - keyword='$keyword'" -Level Info
})

$script:Window.FindName('BtnLMCopyValue').Add_Click({
    $dg = $script:Window.FindName('DgLMResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $vals = [System.Collections.Generic.List[string]]::new()
    foreach ($cellInfo in @($dg.SelectedCells)) {
        $col = $cellInfo.Column -as [System.Windows.Controls.DataGridBoundColumn]
        if ($col) { $vals.Add(([string]$cellInfo.Item.($col.Binding.Path.Path)).Trim()) }
    }
    [System.Windows.Clipboard]::SetText($vals -join "`r`n")
    Show-Notification "$($vals.Count) value(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "LM: Copied cell value to clipboard." -Level Success
})

# ── BtnLMCopyRow ─────────────────────────────────────────────────────────────
$script:Window.FindName('BtnLMCopyRow').Add_Click({
    $dg = $script:Window.FindName('DgLMResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $seen = [System.Collections.Generic.HashSet[object]]::new()
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($cell in @($dg.SelectedCells)) { if ($seen.Add($cell.Item)) { $rows.Add($cell.Item) } }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc    = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add(($colDefs | ForEach-Object { $_.H }) -join "`t")
    foreach ($row in $rows) {
        $vals = $colDefs | ForEach-Object { [string]$row.($_.P) }
        $lines.Add($vals -join "`t")
    }
    [System.Windows.Clipboard]::SetText($lines -join "`r`n")
    Show-Notification "$($rows.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "LM: Copied $($rows.Count) row(s) to clipboard." -Level Success
    # Highlight full rows
    $dg.UnselectAllCells()
    foreach ($r in $rows) { $null = $dg.SelectedItems.Add($r) }
})

# ── BtnLMCopyAll ─────────────────────────────────────────────────────────────
$script:Window.FindName('BtnLMCopyAll').Add_Click({
    $dg       = $script:Window.FindName('DgLMResults')
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) { return }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc    = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add(($colDefs | ForEach-Object { $_.H }) -join "`t")
    foreach ($row in $allItems) {
        $vals = $colDefs | ForEach-Object { [string]$row.($_.P) }
        $lines.Add($vals -join "`t")
    }
    [System.Windows.Clipboard]::SetText($lines -join "`r`n")
    Show-Notification "$($allItems.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "LM: Copied all $($allItems.Count) row(s) to clipboard." -Level Success
})

# ── BtnLMExportXlsx ──────────────────────────────────────────────────────────
$script:Window.FindName('BtnLMExportXlsx').Add_Click({
    $dg       = $script:Window.FindName('DgLMResults')
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) {
        Show-Notification "No results to export." -BgColor '#FFF3CD' -FgColor '#856404'
        return
    }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc    = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title    = "Export Group Members"
    $dlg.Filter   = "Excel Workbook (*.xlsx)|*.xlsx|CSV File (*.csv)|*.csv|All Files (*.*)|*.*"
    $dlg.FileName = "GroupMembers_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    $initDir = try { if ($script:ExportOutPath) { Split-Path $script:ExportOutPath -Parent } else { $null } } catch { $null }
    $dlg.InitialDirectory = if ($initDir -and (Test-Path $initDir)) { $initDir } else { [Environment]::GetFolderPath('Desktop') }
    if ($dlg.ShowDialog() -ne $true) { return }
    $outPath = $dlg.FileName
    try {
        $folder = Split-Path $outPath -Parent
        if ($folder -and -not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        $exportRows = $allItems | ForEach-Object {
            $row = $_; $ht = [ordered]@{}
            foreach ($cd in $colDefs) { $ht[$cd['H']] = [string]$row.($cd['P']) }
            [PSCustomObject]$ht
        }
        if (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue) {
            Import-Module ImportExcel -ErrorAction SilentlyContinue
            $exportRows | Export-Excel -Path $outPath -WorksheetName "GroupMembers" -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium2
            Show-Notification "Exported $($allItems.Count) rows: $outPath" -BgColor '#D4EDDA' -FgColor '#155724'
        } else {
            $exportRows | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8
            Show-Notification "ImportExcel not found — saved as CSV: $(Split-Path $outPath -Leaf)" -BgColor '#FFF3CD' -FgColor '#856404'
        }
        Write-VerboseLog "LM: Exported $($allItems.Count) members to: $outPath" -Level Success
    } catch {
        Show-Notification "Export failed: $($_.Exception.Message)" -BgColor '#F8D7DA' -FgColor '#721C24'
        Write-VerboseLog "LM: Export failed: $($_.Exception.Message)" -Level Error
    }
})

# ── Filter checkboxes re-apply filter without re-querying ────────────────────
foreach ($cbName in @('ChkLMUsers','ChkLMDevices','ChkLMGroups')) {
    $script:Window.FindName($cbName).Add_Click({
        if ($script:LMAllResults -and $script:LMAllResults.Count -gt 0) { Apply-LMFilter }
    })
}



# ════════════════════════════════════════════════════════════════════════════════
# COMPARE GROUPS — handlers
# ════════════════════════════════════════════════════════════════════════════════

# ── BtnCGCommon ── (Common Members) ─────────────────────────────────────────
$script:Window.FindName('BtnCGCommon').Add_Click({
    Hide-Notification
    $groupTxt = $script:Window.FindName('TxtCGGroupInput').Text
    if ([string]::IsNullOrWhiteSpace($groupTxt)) {
        Show-Notification "Please enter at least one group name or Object ID." -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }
    $script:CGGroupInput = $groupTxt
    $script:CGAllResults = $null
    $script:Window.FindName('PnlCGResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtCGNoResults').Visibility = 'Collapsed'
    $script:Window.FindName('PnlCGProgress').Visibility  = 'Visible'
    $script:Window.FindName('PbCGProgress').IsIndeterminate = $true
    $script:Window.FindName('TxtCGProgressMsg').Text    = 'Finding common members...'
    $script:Window.FindName('TxtCGProgressDetail').Text = ''

    if ($script:CGProgressTimer) { $script:CGProgressTimer.Stop() }
    $script:CGProgressTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:CGProgressTimer.Interval = [TimeSpan]::FromMilliseconds(400)
    $script:CGProgressTimer.Add_Tick({
        $prog = if ($script:SharedBg) { $script:SharedBg['CGProgress'] } else { $null }
        if ($prog) {
            $script:Window.FindName('TxtCGProgressMsg').Text    = [string]$prog['Message']
            $script:Window.FindName('TxtCGProgressDetail').Text = [string]$prog['Detail']
        }
    })
    $script:CGProgressTimer.Start()

    $script:Window.FindName('BtnCGDistinct').IsEnabled = $false
    $btn = $script:Window.FindName('BtnCGCommon')
    Invoke-OnBackground -DisableButton $btn -BusyMessage "--- Compare Groups: Common Members ---" -Work {

        $script:CGGroupInput = $Shared['CGGroupInput']
        $groupEntries = $script:CGGroupInput -split "`n" | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() }
        $totalGroups  = $groupEntries.Count
        $allResults   = [System.Collections.Generic.List[PSCustomObject]]::new()

        $Shared['CGProgress'] = @{ Message = "Starting..."; Detail = "0 groups queued" }

        $gi = 0
        foreach ($entry in $groupEntries) {
            $gi++
            $Shared['CGProgress'] = @{ Message = "Resolving group $gi of $($totalGroups): $entry"; Detail = "$($allResults.Count) objects so far" }

            $groupId   = $null
            $groupName = $entry
            if ($entry -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                $groupId = $entry
                try {
                    $gInfo     = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$groupId?`$select=displayName"
                    $groupName = [string]$gInfo['displayName']
                } catch { Write-VerboseLog "CG: Could not fetch display name for ID $entry" -Level Warning }
            } else {
                try {
                    $safe_entry = $entry -replace "'","''"  # PS5-safe: extracted to avoid nested double-quotes in string
                    $gResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe_entry'`&`$select=id,displayName`&`$top=1"
                    if ($gResp['value'] -and $gResp['value'].Count -gt 0) {
                        $groupId   = [string]$gResp['value'][0]['id']
                        $groupName = [string]$gResp['value'][0]['displayName']
                    } else {
                        Write-VerboseLog "CG: Group not found: $entry" -Level Warning
                        continue
                    }
                } catch { Write-VerboseLog "CG: Error resolving '$entry': $($_.Exception.Message)" -Level Error; continue }
            }

            $Shared['CGProgress'] = @{ Message = "Querying group $gi of $($totalGroups): $groupName"; Detail = "$($allResults.Count) objects so far" }
            Write-VerboseLog ('CG: Querying members of ' + $groupName + ' (' + $groupId + ')') -Level Info

            # ── Users ──────────────────────────────────────────────────────────
            try {
                $uri = "https://graph.microsoft.com/v1.0/groups/$groupId/members/microsoft.graph.user?`$select=id,userPrincipalName`&`$top=999"
                do {
                    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
                    foreach ($u in @($resp['value'])) {
                        $allResults.Add([PSCustomObject]@{
                            DisplayName = [string]$u['userPrincipalName']
                            ObjectID    = [string]$u['id']
                            ObjectType  = 'User'
                            SourceGroup = $groupName
                        })
                    }
                    $Shared['CGProgress'] = @{ Message = "Group $gi/$($totalGroups): $groupName (users)"; Detail = "$($allResults.Count) objects so far" }
                    $uri = if ($resp['@odata.nextLink']) { [string]$resp['@odata.nextLink'] } else { $null }
                } while ($uri)
            } catch { Write-VerboseLog "CG: Error querying users from '$groupName': $($_.Exception.Message)" -Level Warning }

            # ── Devices ────────────────────────────────────────────────────────
            try {
                $uri = "https://graph.microsoft.com/v1.0/groups/$groupId/members/microsoft.graph.device?`$select=id,displayName`&`$top=999"
                do {
                    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
                    foreach ($d in @($resp['value'])) {
                        $allResults.Add([PSCustomObject]@{
                            DisplayName = [string]$d['displayName']
                            ObjectID    = [string]$d['id']
                            ObjectType  = 'Device'
                            SourceGroup = $groupName
                        })
                    }
                    $Shared['CGProgress'] = @{ Message = "Group $gi/$($totalGroups): $groupName (devices)"; Detail = "$($allResults.Count) objects so far" }
                    $uri = if ($resp['@odata.nextLink']) { [string]$resp['@odata.nextLink'] } else { $null }
                } while ($uri)
            } catch { Write-VerboseLog "CG: Error querying devices from '$groupName': $($_.Exception.Message)" -Level Warning }

            # ── Nested Groups ──────────────────────────────────────────────────
            try {
                $uri = "https://graph.microsoft.com/v1.0/groups/$groupId/members/microsoft.graph.group?`$select=id,displayName,groupTypes,membershipRuleProcessingState`&`$top=999"
                do {
                    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
                    foreach ($g in @($resp['value'])) {
                        $isM365    = $g['groupTypes'] -and ($g['groupTypes'] -contains 'Unified')
                        $isDynamic = $g['membershipRuleProcessingState'] -eq 'On'
                        $gt = if ($isM365) { 'Microsoft 365' } else { 'Security' }
                        $mt = if ($isDynamic) { 'Dynamic' } else { 'Assigned' }
                        $allResults.Add([PSCustomObject]@{
                            DisplayName = [string]$g['displayName']
                            ObjectID    = [string]$g['id']
                            ObjectType  = ('Group (' + $gt + ', ' + $mt + ')')
                            SourceGroup = $groupName
                        })
                    }
                    $Shared['CGProgress'] = @{ Message = "Group $gi/$($totalGroups): $groupName (nested groups)"; Detail = "$($allResults.Count) objects so far" }
                    $uri = if ($resp['@odata.nextLink']) { [string]$resp['@odata.nextLink'] } else { $null }
                } while ($uri)
            } catch { Write-VerboseLog "CG: Error querying nested groups from '$groupName': $($_.Exception.Message)" -Level Warning }
        }

        $Shared['CGProgress'] = @{ Message = "Complete."; Detail = "$($allResults.Count) object(s) from $($totalGroups) group(s)" }
        Write-VerboseLog "CG: Complete — $($allResults.Count) object(s) from $gi group(s)" -Level Success
        $script:CGAllResults = $allResults
    } -Done {
        if ($script:CGProgressTimer) { $script:CGProgressTimer.Stop() }
        $script:Window.FindName('BtnCGDistinct').IsEnabled = $true
        $script:Window.FindName('PnlCGProgress').Visibility = 'Collapsed'

        if (-not $script:CGAllResults -or $script:CGAllResults.Count -eq 0) {
            $script:Window.FindName('TxtCGNoResults').Visibility = 'Visible'
            Show-Notification "No members found for the specified group(s)." -BgColor '#FFF3CD' -FgColor '#856404'
            return
        }

        # Intersect: ObjectIDs that appear in every successfully resolved source group
        $resolvedGroups = @($script:CGAllResults | Select-Object -ExpandProperty SourceGroup | Sort-Object -Unique)
        $resolvedCount  = $resolvedGroups.Count
        $grouped        = $script:CGAllResults | Group-Object -Property ObjectID
        $commonItems    = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($grp in $grouped) {
            $distinctSrc = @($grp.Group | Select-Object -ExpandProperty SourceGroup | Sort-Object -Unique)
            if ($distinctSrc.Count -ge $resolvedCount) {
                $first = $grp.Group[0]
                $commonItems.Add([PSCustomObject]@{
                    DisplayName = $first.DisplayName
                    ObjectID    = $first.ObjectID
                    ObjectType  = $first.ObjectType
                    SourceGroup = 'Common'
                })
            }
        }

        $dg = $script:Window.FindName('DgCGResults')
        if ($commonItems.Count -eq 0) {
            $dg.ItemsSource = $null
            $script:Window.FindName('TxtCGNoResults').Visibility = 'Visible'
            Show-Notification "No members are common across all $resolvedCount queried group(s)." -BgColor '#FFF3CD' -FgColor '#856404'
            return
        }

        $dg.ItemsSource = $commonItems
        $script:Window.FindName('TxtCGCount').Text           = "$($commonItems.Count) common member(s) across $resolvedCount group(s)"
        $script:Window.FindName('TxtCGNoResults').Visibility = 'Collapsed'
        $script:Window.FindName('PnlCGResults').Visibility   = 'Visible'
        $script:Window.FindName('BtnCGCopyAll').IsEnabled    = $true
        $script:Window.FindName('TxtCGFilter').IsEnabled  = $true
        $script:Window.FindName('BtnCGFilter').IsEnabled  = $true
        $script:Window.FindName('BtnCGFilterClear').IsEnabled = $true
        $script:Window.FindName('BtnCGExportXlsx').IsEnabled = $true
        $script:Window.FindName('BtnCGCopyValue').IsEnabled  = $false
        $script:Window.FindName('BtnCGCopyRow').IsEnabled    = $false
    }
})

# ── BtnCGDistinct ── (Distinct Members) ──────────────────────────────────────
$script:Window.FindName('BtnCGDistinct').Add_Click({
    Hide-Notification
    $groupTxt = $script:Window.FindName('TxtCGGroupInput').Text
    if ([string]::IsNullOrWhiteSpace($groupTxt)) {
        Show-Notification "Please enter at least one group name or Object ID." -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }
    $script:CGGroupInput = $groupTxt
    $script:CGAllResults = $null
    $script:Window.FindName('PnlCGResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtCGNoResults').Visibility = 'Collapsed'
    $script:Window.FindName('PnlCGProgress').Visibility  = 'Visible'
    $script:Window.FindName('PbCGProgress').IsIndeterminate = $true
    $script:Window.FindName('TxtCGProgressMsg').Text    = 'Querying distinct members...'
    $script:Window.FindName('TxtCGProgressDetail').Text = ''

    if ($script:CGProgressTimer) { $script:CGProgressTimer.Stop() }
    $script:CGProgressTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:CGProgressTimer.Interval = [TimeSpan]::FromMilliseconds(400)
    $script:CGProgressTimer.Add_Tick({
        $prog = if ($script:SharedBg) { $script:SharedBg['CGProgress'] } else { $null }
        if ($prog) {
            $script:Window.FindName('TxtCGProgressMsg').Text    = [string]$prog['Message']
            $script:Window.FindName('TxtCGProgressDetail').Text = [string]$prog['Detail']
        }
    })
    $script:CGProgressTimer.Start()

    $script:Window.FindName('BtnCGCommon').IsEnabled = $false
    $btn = $script:Window.FindName('BtnCGDistinct')
    Invoke-OnBackground -DisableButton $btn -BusyMessage "--- Compare Groups: Distinct Members ---" -Work {

        $script:CGGroupInput = $Shared['CGGroupInput']
        $groupEntries = $script:CGGroupInput -split "`n" | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() }
        $totalGroups  = $groupEntries.Count
        $allResults   = [System.Collections.Generic.List[PSCustomObject]]::new()

        $Shared['CGProgress'] = @{ Message = "Starting..."; Detail = "0 groups queued" }

        $gi = 0
        foreach ($entry in $groupEntries) {
            $gi++
            $Shared['CGProgress'] = @{ Message = "Resolving group $gi of $($totalGroups): $entry"; Detail = "$($allResults.Count) objects so far" }

            $groupId   = $null
            $groupName = $entry
            if ($entry -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                $groupId = $entry
                try {
                    $gInfo     = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$groupId?`$select=displayName"
                    $groupName = [string]$gInfo['displayName']
                } catch { Write-VerboseLog "CG: Could not fetch display name for ID $entry" -Level Warning }
            } else {
                try {
                    $safe_entry = $entry -replace "'","''"  # PS5-safe: extracted to avoid nested double-quotes in string
                    $gResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe_entry'`&`$select=id,displayName`&`$top=1"
                    if ($gResp['value'] -and $gResp['value'].Count -gt 0) {
                        $groupId   = [string]$gResp['value'][0]['id']
                        $groupName = [string]$gResp['value'][0]['displayName']
                    } else {
                        Write-VerboseLog "CG: Group not found: $entry" -Level Warning
                        continue
                    }
                } catch { Write-VerboseLog "CG: Error resolving '$entry': $($_.Exception.Message)" -Level Error; continue }
            }

            $Shared['CGProgress'] = @{ Message = "Querying group $gi of $($totalGroups): $groupName"; Detail = "$($allResults.Count) objects so far" }
            Write-VerboseLog ('CG: Querying members of ' + $groupName + ' (' + $groupId + ')') -Level Info

            # ── Users ──────────────────────────────────────────────────────────
            try {
                $uri = "https://graph.microsoft.com/v1.0/groups/$groupId/members/microsoft.graph.user?`$select=id,userPrincipalName`&`$top=999"
                do {
                    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
                    foreach ($u in @($resp['value'])) {
                        $allResults.Add([PSCustomObject]@{
                            DisplayName = [string]$u['userPrincipalName']
                            ObjectID    = [string]$u['id']
                            ObjectType  = 'User'
                            SourceGroup = $groupName
                        })
                    }
                    $Shared['CGProgress'] = @{ Message = "Group $gi/$($totalGroups): $groupName (users)"; Detail = "$($allResults.Count) objects so far" }
                    $uri = if ($resp['@odata.nextLink']) { [string]$resp['@odata.nextLink'] } else { $null }
                } while ($uri)
            } catch { Write-VerboseLog "CG: Error querying users from '$groupName': $($_.Exception.Message)" -Level Warning }

            # ── Devices ────────────────────────────────────────────────────────
            try {
                $uri = "https://graph.microsoft.com/v1.0/groups/$groupId/members/microsoft.graph.device?`$select=id,displayName`&`$top=999"
                do {
                    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
                    foreach ($d in @($resp['value'])) {
                        $allResults.Add([PSCustomObject]@{
                            DisplayName = [string]$d['displayName']
                            ObjectID    = [string]$d['id']
                            ObjectType  = 'Device'
                            SourceGroup = $groupName
                        })
                    }
                    $Shared['CGProgress'] = @{ Message = "Group $gi/$($totalGroups): $groupName (devices)"; Detail = "$($allResults.Count) objects so far" }
                    $uri = if ($resp['@odata.nextLink']) { [string]$resp['@odata.nextLink'] } else { $null }
                } while ($uri)
            } catch { Write-VerboseLog "CG: Error querying devices from '$groupName': $($_.Exception.Message)" -Level Warning }

            # ── Nested Groups ──────────────────────────────────────────────────
            try {
                $uri = "https://graph.microsoft.com/v1.0/groups/$groupId/members/microsoft.graph.group?`$select=id,displayName,groupTypes,membershipRuleProcessingState`&`$top=999"
                do {
                    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
                    foreach ($g in @($resp['value'])) {
                        $isM365    = $g['groupTypes'] -and ($g['groupTypes'] -contains 'Unified')
                        $isDynamic = $g['membershipRuleProcessingState'] -eq 'On'
                        $gt = if ($isM365) { 'Microsoft 365' } else { 'Security' }
                        $mt = if ($isDynamic) { 'Dynamic' } else { 'Assigned' }
                        $allResults.Add([PSCustomObject]@{
                            DisplayName = [string]$g['displayName']
                            ObjectID    = [string]$g['id']
                            ObjectType  = ('Group (' + $gt + ', ' + $mt + ')')
                            SourceGroup = $groupName
                        })
                    }
                    $Shared['CGProgress'] = @{ Message = "Group $gi/$($totalGroups): $groupName (nested groups)"; Detail = "$($allResults.Count) objects so far" }
                    $uri = if ($resp['@odata.nextLink']) { [string]$resp['@odata.nextLink'] } else { $null }
                } while ($uri)
            } catch { Write-VerboseLog "CG: Error querying nested groups from '$groupName': $($_.Exception.Message)" -Level Warning }
        }

        $Shared['CGProgress'] = @{ Message = "Complete."; Detail = "$($allResults.Count) object(s) from $($totalGroups) group(s)" }
        Write-VerboseLog "CG: Complete — $($allResults.Count) object(s) from $gi group(s)" -Level Success
        $script:CGAllResults = $allResults
    } -Done {
        if ($script:CGProgressTimer) { $script:CGProgressTimer.Stop() }
        $script:Window.FindName('BtnCGCommon').IsEnabled  = $true
        $script:Window.FindName('PnlCGProgress').Visibility = 'Collapsed'

        if (-not $script:CGAllResults -or $script:CGAllResults.Count -eq 0) {
            $script:Window.FindName('TxtCGNoResults').Visibility = 'Visible'
            Show-Notification "No members found for the specified group(s)." -BgColor '#FFF3CD' -FgColor '#856404'
            return
        }

        $dg = $script:Window.FindName('DgCGResults')
        $dg.ItemsSource = $script:CGAllResults
        $groupCount = ($script:CGGroupInput -split "`n" | Where-Object { $_.Trim() -ne '' }).Count
        $script:Window.FindName('TxtCGCount').Text           = "$($script:CGAllResults.Count) object(s) across $groupCount group(s)"
        $script:Window.FindName('TxtCGNoResults').Visibility = 'Collapsed'
        $script:Window.FindName('PnlCGResults').Visibility   = 'Visible'
        $script:Window.FindName('BtnCGCopyAll').IsEnabled    = $true
        $script:Window.FindName('TxtCGFilter').IsEnabled  = $true
        $script:Window.FindName('BtnCGFilter').IsEnabled  = $true
        $script:Window.FindName('BtnCGFilterClear').IsEnabled = $true
        $script:Window.FindName('BtnCGExportXlsx').IsEnabled = $true
        $script:Window.FindName('BtnCGCopyValue').IsEnabled  = $false
        $script:Window.FindName('BtnCGCopyRow').IsEnabled    = $false
    }
})

# ── DgCGResults: cell selection ───────────────────────────────────────────────
$script:Window.FindName('DgCGResults').Add_SelectedCellsChanged({
    $dg       = $script:Window.FindName('DgCGResults')
    $selCells = $dg.SelectedCells.Count
    $script:Window.FindName('BtnCGCopyValue').IsEnabled = ($selCells -ge 1)
    $script:Window.FindName('BtnCGCopyRow').IsEnabled   = ($selCells -gt 0)
})

# ── BtnCGCopyValue ────────────────────────────────────────────────────────────
# ── BtnCGFilter ──
$script:Window.FindName('BtnCGFilter').Add_Click({
    $dg      = $script:Window.FindName('DgCGResults')
    $keyword = $script:Window.FindName('TxtCGFilter').Text.Trim()
    if ($null -eq $script:CGAllData) {
        $script:CGAllData = @($dg.ItemsSource)
    }
    if ($keyword -eq '') {
        $dg.ItemsSource = $script:CGAllData
        Show-Notification "Filter cleared - $($script:CGAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    } else {
        $filtered = @($script:CGAllData | Where-Object {
            $row = $_
            $match = $false
            foreach ($prop in $row.PSObject.Properties) {
                if ([string]$prop.Value -like "*$keyword*") { $match = $true; break }
            }
            $match
        })

# ── BtnCGFilterClear ──
$script:Window.FindName('BtnCGFilterClear').Add_Click({
    $script:Window.FindName('TxtCGFilter').Text = ''
    $dg = $script:Window.FindName('DgCGResults')
    if ($null -ne $script:CGAllData) {
        $dg.ItemsSource = $script:CGAllData
        Show-Notification "Filter cleared - $($script:CGAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    }
})
        $dg.ItemsSource = $filtered
        Show-Notification "$($filtered.Count) of $($script:CGAllData.Count) result(s) shown after filter." -BgColor '#D4EDDA' -FgColor '#155724'
    }
    Write-VerboseLog "CG: Filter applied - keyword='$keyword'" -Level Info
})

$script:Window.FindName('BtnCGCopyValue').Add_Click({
    $dg = $script:Window.FindName('DgCGResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $vals = [System.Collections.Generic.List[string]]::new()
    foreach ($cellInfo in @($dg.SelectedCells)) {
        $col = $cellInfo.Column -as [System.Windows.Controls.DataGridBoundColumn]
        if ($col) { $vals.Add(([string]$cellInfo.Item.($col.Binding.Path.Path)).Trim()) }
    }
    [System.Windows.Clipboard]::SetText($vals -join "`r`n")
    Show-Notification "$($vals.Count) value(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "CG: Copied cell value to clipboard." -Level Success
})

# ── BtnCGCopyRow ──────────────────────────────────────────────────────────────
$script:Window.FindName('BtnCGCopyRow').Add_Click({
    $dg = $script:Window.FindName('DgCGResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $seen = [System.Collections.Generic.HashSet[object]]::new()
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($cell in @($dg.SelectedCells)) { if ($seen.Add($cell.Item)) { $rows.Add($cell.Item) } }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc    = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add(($colDefs | ForEach-Object { $_.H }) -join "`t")
    foreach ($row in $rows) {
        $vals = $colDefs | ForEach-Object { [string]$row.($_.P) }
        $lines.Add($vals -join "`t")
    }
    [System.Windows.Clipboard]::SetText($lines -join "`r`n")
    Show-Notification "$($rows.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "CG: Copied $($rows.Count) row(s) to clipboard." -Level Success
    # Highlight full rows
    $dg.UnselectAllCells()
    foreach ($r in $rows) { $null = $dg.SelectedItems.Add($r) }
})

# ── BtnCGCopyAll ──────────────────────────────────────────────────────────────
$script:Window.FindName('BtnCGCopyAll').Add_Click({
    $dg       = $script:Window.FindName('DgCGResults')
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) { return }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc    = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add(($colDefs | ForEach-Object { $_.H }) -join "`t")
    foreach ($row in $allItems) {
        $vals = $colDefs | ForEach-Object { [string]$row.($_.P) }
        $lines.Add($vals -join "`t")
    }
    [System.Windows.Clipboard]::SetText($lines -join "`r`n")
    Show-Notification "$($allItems.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "CG: Copied all $($allItems.Count) row(s) to clipboard." -Level Success
})

# ── BtnCGExportXlsx ───────────────────────────────────────────────────────────
$script:Window.FindName('BtnCGExportXlsx').Add_Click({
    $dg       = $script:Window.FindName('DgCGResults')
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) {
        Show-Notification "No results to export." -BgColor '#FFF3CD' -FgColor '#856404'
        return
    }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc    = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title    = "Export Compare Groups Results"
    $dlg.Filter   = "Excel Workbook (*.xlsx)|*.xlsx|CSV File (*.csv)|*.csv|All Files (*.*)|*.*"
    $dlg.FileName = "CompareGroups_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    $initDir = try { if ($script:ExportOutPath) { Split-Path $script:ExportOutPath -Parent } else { $null } } catch { $null }
    $dlg.InitialDirectory = if ($initDir -and (Test-Path $initDir)) { $initDir } else { [Environment]::GetFolderPath('Desktop') }
    if ($dlg.ShowDialog() -ne $true) { return }
    $outPath = $dlg.FileName
    try {
        $folder = Split-Path $outPath -Parent
        if ($folder -and -not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        $exportRows = $allItems | ForEach-Object {
            $row = $_; $ht = [ordered]@{}
            foreach ($cd in $colDefs) { $ht[$cd['H']] = [string]$row.($cd['P']) }
            [PSCustomObject]$ht
        }
        if (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue) {
            Import-Module ImportExcel -ErrorAction SilentlyContinue
            $exportRows | Export-Excel -Path $outPath -WorksheetName "CompareGroups" -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium6
            Show-Notification "Exported $($allItems.Count) rows: $outPath" -BgColor '#D4EDDA' -FgColor '#155724'
        } else {
            $exportRows | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8
            Show-Notification "ImportExcel not found — saved as CSV: $(Split-Path $outPath -Leaf)" -BgColor '#FFF3CD' -FgColor '#856404'
        }
        Write-VerboseLog "CG: Exported $($allItems.Count) rows to: $outPath" -Level Success
    } catch {
        Show-Notification "Export failed: $($_.Exception.Message)" -BgColor '#F8D7DA' -FgColor '#721C24'
        Write-VerboseLog "CG: Export failed: $($_.Exception.Message)" -Level Error
    }
})

# ── RENAME GROUP ──
$script:Window.FindName('BtnRenValidate').Add_Click({
    Hide-Notification
    if ($null -eq $script:RenSelectedGroup) {
        Show-Notification "Please select a target group first." -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }
    $newName = $script:Window.FindName('TxtNewGroupName').Text.Trim()
    if ([string]::IsNullOrWhiteSpace($newName)) {
        Show-Notification "Please enter a new display name." -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }
    $script:RenNewName = $newName
    $btn = $script:Window.FindName('BtnRenValidate')

    Invoke-OnBackground -DisableButton $btn -BusyMessage "Checking name uniqueness..." -Work {
        $newName = $script:RenNewName
        Write-VerboseLog "Current: $($script:RenSelectedGroup.DisplayName) -> New: $newName" -Level Info
        $existing = $null
        try {
            $safe_newName = $newName -replace "'","''"  # PS5-safe: extracted to avoid nested double-quotes in string
            $resp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe_newName'`&`$select=id`&`$top=1"
            if ($resp['value'] -and $resp['value'].Count -gt 0) { $existing = $resp['value'][0]['id'] }
        } catch {}
        if ($existing) { Write-VerboseLog "WARNING: A group named '$newName' already exists." -Level Warning }
        $script:RenExistingId = $existing
    } -Done {
        $newName  = $script:RenNewName
        $existing = $script:RenExistingId
        if ($existing) {
            Show-Notification "Warning: A group named '$newName' already exists. Continuing will create a duplicate." -BgColor '#FFF3CD' -FgColor '#7A4800'
        }
        $emptyValid   = [System.Collections.Generic.List[PSCustomObject]]::new()
        $emptyInvalid = [System.Collections.Generic.List[string]]::new()

        $confirmed = Show-ConfirmationDialog `
            -Title "Rename Group" -OperationLabel "Rename Group" `
            -TargetGroup $script:RenSelectedGroup.DisplayName `
            -TargetGroupId $script:RenSelectedGroup.Id `
            -ValidObjects $emptyValid -InvalidEntries $emptyInvalid `
            -ExtraInfo "Current name : $($script:RenSelectedGroup.DisplayName)`nNew name     : $newName"

        if (-not $confirmed) { Write-VerboseLog "Action cancelled by user." -Level Warning; return }

        $btn = $script:Window.FindName('BtnRenValidate')
        Invoke-OnBackground -DisableButton $btn -BusyMessage "--- Renaming group ---" -Work {
            Invoke-MgGraphRequest -Method PATCH `
                -Uri "https://graph.microsoft.com/v1.0/groups/$($script:RenSelectedGroup.Id)" `
                -Body (@{ displayName = $script:RenNewName } | ConvertTo-Json) -ContentType "application/json"
            Write-VerboseLog "Group renamed to '$($script:RenNewName)'" -Level Success
        } -Done {
            Show-Notification "Group renamed to '$($script:RenNewName)'" -BgColor '#D4EDDA' -FgColor '#155724'
            $script:Window.FindName('TxtRenGroupCurrentName').Text = $script:RenNewName
            $script:RenSelectedGroup = [PSCustomObject]@{
                Id = $script:RenSelectedGroup.Id; DisplayName = $script:RenNewName
                GroupTypes = $script:RenSelectedGroup.GroupTypes; MembershipRule = $script:RenSelectedGroup.MembershipRule
            }
        }
    }
})

# ── SET GROUP OWNER (multi-group) ──
$script:Window.FindName('BtnOwnerValidate').Add_Click({
    Hide-Notification
    $groupListTxt = $script:Window.FindName('TxtOwnerGroupList').Text
    if ([string]::IsNullOrWhiteSpace($groupListTxt)) {
        Show-Notification "Please enter at least one target group (name or ID)." -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }
    $ownerTxt = $script:Window.FindName('TxtOwnerList').Text
    if ([string]::IsNullOrWhiteSpace($ownerTxt)) {
        Show-Notification "Please enter at least one owner UPN." -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }
    $script:OwnerGroupsTxt = $groupListTxt
    $script:OwnerUpnsTxt   = $ownerTxt
    $btn = $script:Window.FindName('BtnOwnerValidate')

    Invoke-OnBackground -DisableButton $btn -BusyMessage "--- Set Owner: Validating groups and UPNs ---" -Work {

        # -- Resolve target groups --
        $groupLines    = $script:OwnerGroupsTxt -split "`n" | Where-Object { $_.Trim() -ne '' }
        $validGroups   = [System.Collections.Generic.List[PSCustomObject]]::new()
        $invalidGroups = [System.Collections.Generic.List[string]]::new()
        foreach ($line in $groupLines) {
            $entry = $line.Trim()
            if ([string]::IsNullOrWhiteSpace($entry)) { continue }
            Write-VerboseLog "Resolving group: $entry" -Level Info
            $g = Resolve-GroupByNameOrId -GroupEntry $entry
            if ($g) {
                $validGroups.Add($g)
                Write-VerboseLog "  $([char]0x2713) $($g.DisplayName) ($($g.Id))" -Level Success
            } else {
                $invalidGroups.Add($entry)
                Write-VerboseLog "  $([char]0x2717) Not found: $entry" -Level Warning
            }
        }

        # -- Resolve owner UPNs --
        $upns          = $script:OwnerUpnsTxt -split "`n" | Where-Object { $_.Trim() -ne '' }
        $validOwners   = [System.Collections.Generic.List[PSCustomObject]]::new()
        $invalidOwners = [System.Collections.Generic.List[string]]::new()
        foreach ($upn in $upns) {
            $u = $upn.Trim()
            if ([string]::IsNullOrWhiteSpace($u)) { continue }
            Write-VerboseLog "Resolving owner UPN: $u" -Level Info
            try {
                $userObj = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/users/$([Uri]::EscapeDataString($u))?`$select=id,displayName"
                $validOwners.Add([PSCustomObject]@{ Id = $userObj['id']; DisplayName = $userObj['displayName']; Type = 'User'; Found = $true; Original = $u })
                Write-VerboseLog "  $([char]0x2713) $($userObj['displayName']) ($($userObj['id']))" -Level Success
            } catch {
                $invalidOwners.Add($u)
                Write-VerboseLog "  $([char]0x2717) Not found: $u" -Level Warning
            }
        }

        $script:OwnerGroupsValidated = @{ ValidGroups = $validGroups; InvalidGroups = $invalidGroups }
        $script:OwnerValidated       = @{ Valid = $validOwners; Invalid = $invalidOwners }

    } -Done {
        $validGroups   = $script:OwnerGroupsValidated['ValidGroups']
        $invalidGroups = $script:OwnerGroupsValidated['InvalidGroups']
        $validOwners   = $script:OwnerValidated['Valid']
        $invalidOwners = $script:OwnerValidated['Invalid']

        if ($validGroups.Count -eq 0) {
            Show-Notification "No valid target groups found. Cannot proceed." -BgColor '#F8D7DA' -FgColor '#721C24'
            return
        }
        if ($validOwners.Count -eq 0) {
            Show-Notification "No valid owner UPNs found. Cannot proceed." -BgColor '#F8D7DA' -FgColor '#721C24'
            return
        }

        # Build ExtraInfo block listing all target groups for the confirmation dialog
        $sbGroups = [System.Text.StringBuilder]::new()
        $null = $sbGroups.AppendLine(('TARGET GROUPS (' + $validGroups.Count + ' valid, ' + $invalidGroups.Count + ' not found):'))
        foreach ($g   in $validGroups)   { $null = $sbGroups.AppendLine("  $([char]0x2713) $($g.DisplayName) ($($g.Id))") }
        foreach ($inv in $invalidGroups) { $null = $sbGroups.AppendLine("  $([char]0x2717) $inv  (not found - will be skipped)") }
        $extraInfo = $sbGroups.ToString().TrimEnd()

        $confirmed = Show-ConfirmationDialog `
            -Title "Set Group Owner (Multi-Group)" -OperationLabel "Set Group Owner" `
            -TargetGroup "Multiple groups ($($validGroups.Count))" `
            -TargetGroupId "(see group list below)" `
            -ValidObjects $validOwners -InvalidEntries $invalidOwners `
            -ExtraInfo $extraInfo

        if (-not $confirmed) { Write-VerboseLog "Action cancelled by user." -Level Warning; return }

        $btn = $script:Window.FindName('BtnOwnerValidate')
        Invoke-OnBackground -DisableButton $btn -BusyMessage "--- Adding owners to groups ---" -Work {
            $ok = 0; $fail = 0; $skipped = 0
            $validGroups = $script:OwnerGroupsValidated['ValidGroups']
            foreach ($grp in $validGroups) {
                Write-VerboseLog "Processing group: $($grp.DisplayName)" -Level Info
                foreach ($owner in $script:OwnerValidated['Valid']) {
                    try {
                        $ob = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($owner.Id)" }
                        Invoke-MgGraphRequest -Method POST `
                            -Uri "https://graph.microsoft.com/v1.0/groups/$($grp.Id)/owners/`$ref" `
                            -Body ($ob | ConvertTo-Json) -ContentType "application/json"
                        Write-VerboseLog "Owner added: $($owner.DisplayName) -> $($grp.DisplayName)" -Level Success
                        $ok++
                    } catch {
                        $errMsg    = $_.Exception.Message
                        $errDetail = if ($_.ErrorDetails) { $_.ErrorDetails.Message } else { '' }
                        if ($errMsg -match 'already exist' -or $errDetail -match 'already exist') {
                            Write-VerboseLog "Skipping $($owner.DisplayName) - already an owner of $($grp.DisplayName)" -Level Warning
                            $skipped++
                        } else {
                            Write-VerboseLog "Failed: $($owner.DisplayName) -> $($grp.DisplayName) - $errMsg" -Level Error
                            $fail++
                        }
                    }
                }
            }
            $script:OwnerExecResult = @{ Ok = $ok; Fail = $fail; Skipped = $skipped }
        } -Done {
            $ok      = $script:OwnerExecResult['Ok']
            $fail    = $script:OwnerExecResult['Fail']
            $skipped = $script:OwnerExecResult['Skipped']
            $msg     = "Set Owner complete. Succeeded: $ok | Skipped (already owner): $skipped | Failed: $fail"
            $vLevel  = if ($fail -gt 0) { 'Warning' } else { 'Success' }
            $bgColor = if ($fail -gt 0) { '#FFF3CD' } else { '#D4EDDA' }
            $fgColor = if ($fail -gt 0) { '#7A4800' } else { '#155724' }
            Write-VerboseLog $msg -Level $vLevel
            Show-Notification $msg -BgColor $bgColor -FgColor $fgColor
        }
    }
})

# ── REMOVE GROUPS ─────────────────────────────────────────────────────────────

$script:Window.FindName('BtnOpRemoveGroups').Add_Click({
    $script:Window.FindName('TxtRGGroupList').Text            = ''
    $script:Window.FindName('PnlRGProgress').Visibility       = 'Collapsed'
    $script:Window.FindName('PnlRGWarning').Visibility        = 'Collapsed'
    $script:Window.FindName('PnlRGPreview').Visibility        = 'Collapsed'
    $script:Window.FindName('PnlRGExecProgress').Visibility   = 'Collapsed'
    $script:Window.FindName('PnlRGResults').Visibility        = 'Collapsed'
    $script:Window.FindName('TxtRGNoResults').Visibility      = 'Collapsed'
    $dg = $script:Window.FindName('DgRGPreview')
    if ($dg) { $dg.ItemsSource = $null }
    $dg2 = $script:Window.FindName('DgRGResults')
    if ($dg2) { $dg2.ItemsSource = $null }
    $script:Window.FindName('BtnRGExecute').IsEnabled         = $false
    $script:Window.FindName('BtnRGCopyAll').IsEnabled         = $false
    $script:Window.FindName('BtnRGExportXlsx').IsEnabled      = $false
    $script:RGGroupList  = $null
    $script:RGValidated  = $null
    $script:RGExecResult = $null
    Show-Panel 'PanelRemoveGroups'
    Hide-Notification
    Write-VerboseLog 'Panel: Remove Groups' -Level Info
})

$script:Window.FindName('BtnRGValidate').Add_Click({
    Hide-Notification
    $inputTxt = $script:Window.FindName('TxtRGGroupList').Text
    if ([string]::IsNullOrWhiteSpace($inputTxt)) {
        Show-Notification 'Please enter at least one group name or Object ID.' -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }
    $script:RGGroupList = $inputTxt
    $dg = $script:Window.FindName('DgRGPreview')
    if ($dg) { $dg.ItemsSource = $null }
    $dg2 = $script:Window.FindName('DgRGResults')
    if ($dg2) { $dg2.ItemsSource = $null }
    $script:Window.FindName('PnlRGPreview').Visibility        = 'Collapsed'
    $script:Window.FindName('PnlRGResults').Visibility        = 'Collapsed'
    $script:Window.FindName('PnlRGWarning').Visibility        = 'Collapsed'
    $script:Window.FindName('TxtRGNoResults').Visibility      = 'Collapsed'
    $script:Window.FindName('PnlRGExecProgress').Visibility   = 'Collapsed'
    $script:Window.FindName('BtnRGExecute').IsEnabled         = $false
    $script:Window.FindName('BtnRGCopyAll').IsEnabled         = $false
    $script:Window.FindName('BtnRGExportXlsx').IsEnabled      = $false
    $script:Window.FindName('PnlRGProgress').Visibility       = 'Visible'
    $script:Window.FindName('TxtRGProgressMsg').Text          = 'Validating groups...'
    $script:Window.FindName('TxtRGProgressDetail').Text       = ''

    $btn = $script:Window.FindName('BtnRGValidate')
    Invoke-OnBackground -DisableButton $btn -BusyMessage '--- Remove Groups: Validating ---' -Work {
        $inputLines = ($script:RGGroupList -split "`n") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
        $validated  = [System.Collections.Generic.List[PSCustomObject]]::new()
        $total      = $inputLines.Count
        $idx        = 0

        foreach ($entry in $inputLines) {
            $idx++
            $Shared['LogQueue'].Enqueue(@{ Line = "RG: Resolving ($idx/$total): $entry"; Level = 'Info' })

            $grp = Resolve-GroupByNameOrId -GroupEntry $entry
            if ($null -eq $grp) {
                $validated.Add([PSCustomObject]@{
                    DisplayName = $entry
                    GroupID     = '(not found)'
                    GroupType   = [char]0x2014
                    MemberCount = [char]0x2014
                    HasMembers  = $false
                    IsResolved  = $false
                    OrigEntry   = $entry
                })
                $Shared['LogQueue'].Enqueue(@{ Line = "RG:   Not found: $entry"; Level = 'Warning' })
                continue
            }

            # Determine group type from Graph metadata
            $groupTypeName = 'Unknown'
            try {
                $meta         = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($grp.Id)?`$select=mailEnabled,securityEnabled,groupTypes"
                $mailEnabled  = [bool]$meta['mailEnabled']
                $secEnabled   = [bool]$meta['securityEnabled']
                $gtArr        = @($meta['groupTypes'])
                $groupTypeName = if ($gtArr -contains 'Unified')             { 'Microsoft 365' }
                                 elseif ($mailEnabled -and -not $secEnabled) { 'Distribution' }
                                 elseif ($secEnabled  -and -not $mailEnabled) { 'Security' }
                                 elseif ($secEnabled  -and $mailEnabled)      { 'Mail-Enabled Security' }
                                 else                                          { 'Unknown' }
            } catch {}

            $memberCount = Get-GroupMemberCount -GroupId $grp.Id
            $mc          = if ($null -eq $memberCount) { '?' } else { [string]$memberCount }
            $hasMembers  = ($null -ne $memberCount -and $memberCount -gt 0)

            $validated.Add([PSCustomObject]@{
                DisplayName = $grp.DisplayName
                GroupID     = $grp.Id
                GroupType   = $groupTypeName
                MemberCount = $mc
                HasMembers  = $hasMembers
                IsResolved  = $true
                OrigEntry   = $entry
            })
            $Shared['LogQueue'].Enqueue(@{ Line = "RG:   Resolved: $($grp.DisplayName) [$($grp.Id)] | $groupTypeName | Members: $mc"; Level = 'Success' })
        }
        $script:RGValidated = $validated

    } -Done {
        $script:Window.FindName('PnlRGProgress').Visibility = 'Collapsed'
        $rows = $script:RGValidated

        if ($null -eq $rows -or $rows.Count -eq 0) {
            $script:Window.FindName('TxtRGNoResults').Visibility = 'Visible'
            Show-Notification 'No groups resolved from the provided input.' -BgColor '#FFF3CD' -FgColor '#7A4800'
            return
        }

        $resolved    = @($rows | Where-Object { $_.IsResolved })
        $withMembers = @($rows | Where-Object { $_.HasMembers })
        $notFound    = @($rows | Where-Object { -not $_.IsResolved })

        if ($resolved.Count -eq 0) {
            Show-Notification 'No groups could be resolved. Check the names/IDs and try again.' -BgColor '#FFF3CD' -FgColor '#7A4800'
            return
        }

        # Build ValidObjects and InvalidEntries for Show-ConfirmationDialog
        $validObjs      = @($resolved | ForEach-Object { [PSCustomObject]@{ DisplayName = $_.DisplayName; Type = $_.GroupType; Id = $_.GroupID } })
        $invalidEntries = @($notFound  | ForEach-Object { $_.OrigEntry })

        # ExtraInfo — warn about groups that still have members
        $extra = ''
        if ($withMembers.Count -gt 0) {
            $memberLines = ($withMembers | ForEach-Object { "  $([char]0x26A0) $($_.DisplayName) ($($_.MemberCount) member(s))" }) -join "`n"
            $extra = "$([char]0x26A0) $($withMembers.Count) group$(if($withMembers.Count -eq 1){''}else{'s'}) still ha$(if($withMembers.Count -eq 1){'s'}else{'ve'}) members.`nMembers lose group access but are NOT deleted:`n$memberLines"
        }

        $msg = "Validation complete. $($resolved.Count) group(s) ready, $($notFound.Count) not found."
        Write-VerboseLog "RG: $msg" -Level Info

        $confirmed = Show-ConfirmationDialog `
            -Title 'Remove Groups' -OperationLabel 'Remove Groups' `
            -TargetGroup "$($resolved.Count) group(s) selected" -TargetGroupId '(see list below)' `
            -ValidObjects $validObjs -InvalidEntries $invalidEntries `
            -ExtraInfo $extra

        if (-not $confirmed) { Write-VerboseLog 'RG: Action cancelled by user.' -Level Warning; return }

        # Set up UI for execution progress
        $dg2 = $script:Window.FindName('DgRGResults')
        if ($dg2) { $dg2.ItemsSource = $null }
        $script:Window.FindName('PnlRGResults').Visibility      = 'Collapsed'
        $script:Window.FindName('BtnRGCopyAll').IsEnabled       = $false
        $script:Window.FindName('BtnRGExportXlsx').IsEnabled    = $false
        $script:Window.FindName('PnlRGExecProgress').Visibility = 'Visible'
        $script:Window.FindName('TxtRGExecProgressMsg').Text    = 'Removing groups...'
        $script:Window.FindName('TxtRGExecProgressDetail').Text = ''

        $btn = $script:Window.FindName('BtnRGValidate')
        Invoke-OnBackground -DisableButton $btn -BusyMessage '--- Remove Groups: Executing ---' -Work {

            $toDelete = @($script:RGValidated | Where-Object { $_.IsResolved })
            $results  = [System.Collections.Generic.List[PSCustomObject]]::new()
            $total    = $toDelete.Count
            $idx      = 0

            foreach ($g in $toDelete) {
                $idx++
                $Shared['LogQueue'].Enqueue(@{ Line = "RG: Deleting ($idx/$total): $($g.DisplayName)"; Level = 'Action' })
                try {
                    $null = Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$($g.GroupID)"
                    $results.Add([PSCustomObject]@{
                        DisplayName = $g.DisplayName
                        GroupID     = $g.GroupID
                        Status      = 'Removed'
                        Message     = 'Successfully deleted.'
                        IsSuccess   = $true
                    })
                    $Shared['LogQueue'].Enqueue(@{ Line = "RG:   Removed: $($g.DisplayName) [$($g.GroupID)]"; Level = 'Success' })
                } catch {
                    $errMsg = $_.Exception.Message
                    $results.Add([PSCustomObject]@{
                        DisplayName = $g.DisplayName
                        GroupID     = $g.GroupID
                        Status      = 'Failed'
                        Message     = $errMsg
                        IsSuccess   = $false
                    })
                    $Shared['LogQueue'].Enqueue(@{ Line = "RG:   Failed: $($g.DisplayName) - $errMsg"; Level = 'Error' })
                }
            }
            $script:RGExecResult = $results

        } -Done {
            $script:Window.FindName('PnlRGExecProgress').Visibility = 'Collapsed'
            $exRows = $script:RGExecResult

            if ($null -eq $exRows -or $exRows.Count -eq 0) {
                Show-Notification 'No results returned from deletion.' -BgColor '#FFF3CD' -FgColor '#7A4800'
                return
            }

            $dg2 = $script:Window.FindName('DgRGResults')
            $dg2.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSCustomObject]]::new($exRows)
            $ok   = @($exRows | Where-Object { $_.IsSuccess  }).Count
            $fail = @($exRows | Where-Object { -not $_.IsSuccess }).Count
            $script:Window.FindName('TxtRGResultCount').Text       = "Removed: $ok  |  Failed: $fail"
            $script:Window.FindName('PnlRGResults').Visibility     = 'Visible'
            $script:Window.FindName('BtnRGCopyAll').IsEnabled      = $true
            $script:Window.FindName('BtnRGExportXlsx').IsEnabled   = $true
            $rgMsg   = "Remove Groups complete. Removed: $ok | Failed: $fail"
            $bgColor = if ($fail -gt 0) { '#FFF3CD' } else { '#D4EDDA' }
            $fgColor = if ($fail -gt 0) { '#7A4800' } else { '#155724' }
            Show-Notification $rgMsg -BgColor $bgColor -FgColor $fgColor
            Write-VerboseLog "RG: $rgMsg" -Level $(if ($fail -gt 0) { 'Warning' } else { 'Success' })
        }
    }
})

# BtnRGExecute is no longer used as a separate button — Remove Groups validation now flows
# directly into Show-ConfirmationDialog and execution is triggered from BtnRGValidate Done.

# ── BtnRGCopyAll ───────────────────────────────────────────────────────────────
$script:Window.FindName('BtnRGCopyAll').Add_Click({
    $dg       = $script:Window.FindName('DgRGResults')
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) { return }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc     = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $txtLines = [System.Collections.Generic.List[string]]::new()
    $txtLines.Add(($colDefs | ForEach-Object { $_.H }) -join "`t")
    foreach ($row in $allItems) {
        $vals = $colDefs | ForEach-Object { [string]$row.($_.P) }
        $txtLines.Add($vals -join "`t")
    }
    [System.Windows.Clipboard]::SetText($txtLines -join "`r`n")
    Show-Notification "$($allItems.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "RG: Copied all $($allItems.Count) row(s) to clipboard." -Level Success
})

# ── BtnRGExportXlsx ───────────────────────────────────────────────────────────
$script:Window.FindName('BtnRGExportXlsx').Add_Click({
    $dg       = $script:Window.FindName('DgRGResults')
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) {
        Show-Notification 'No results to export.' -BgColor '#FFF3CD' -FgColor '#856404'
        return
    }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc     = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title            = 'Export Remove Groups Results'
    $dlg.Filter           = 'Excel Workbook (*.xlsx)|*.xlsx|CSV File (*.csv)|*.csv|All Files (*.*)|*.*'
    $dlg.FileName         = "RemoveGroups_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($dlg.ShowDialog() -ne $true) { return }
    $outPath = $dlg.FileName
    try {
        $folder = Split-Path $outPath -Parent
        if ($folder -and -not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        $exportRows = $allItems | ForEach-Object {
            $rowObj = $_; $ht = [ordered]@{}
            foreach ($cd in $colDefs) { $ht[$cd['H']] = [string]$rowObj.($cd['P']) }
            [PSCustomObject]$ht
        }
        if (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue) {
            Import-Module ImportExcel -ErrorAction SilentlyContinue
            $exportRows | Export-Excel -Path $outPath -WorksheetName 'RemoveGroups' -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium6
            Show-Notification "Exported $($allItems.Count) rows: $outPath" -BgColor '#D4EDDA' -FgColor '#155724'
        } else {
            $exportRows | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8
            Show-Notification "ImportExcel not found — saved as CSV: $(Split-Path $outPath -Leaf)" -BgColor '#FFF3CD' -FgColor '#856404'
        }
        Write-VerboseLog "RG: Exported $($allItems.Count) rows to: $outPath" -Level Success
    } catch {
        Show-Notification "Export failed: $($_.Exception.Message)" -BgColor '#F8D7DA' -FgColor '#721C24'
        Write-VerboseLog "RG: Export failed: $($_.Exception.Message)" -Level Error
    }
})

# ── SESSION NOTES ──────────────────────────────────────────────────────────────
$script:Window.FindName('BtnSessionNotes').Add_Click({
    $script:Window.FindName('OverlaySessionNotes').Visibility = 'Visible'
    $script:Window.FindName('TxtSessionNotes').Focus() | Out-Null
})
$script:Window.FindName('BtnSessionNotesClose').Add_Click({
    $script:Window.FindName('OverlaySessionNotes').Visibility = 'Collapsed'
})
$script:Window.FindName('SnDimmer').Add_MouseLeftButtonDown({
    $script:Window.FindName('OverlaySessionNotes').Visibility = 'Collapsed'
})

# ── ADD USER DEVICES TO GROUP ──────────────────────────────────────────────

Wire-GroupPicker `
    -SearchBoxName 'TxtUDGroupSearch' -SearchBtnName 'BtnUDGroupSearch' -ByIdBtnName 'BtnUDGroupById' `
    -ResultsListName 'LstUDGroupResults' -ResultsBorderName 'UDGroupSearchResults' `
    -BadgeName 'UDSelectedGroupBadge' -BadgeNameTxt 'TxtUDGroupName' -BadgeIdTxt 'TxtUDGroupId' `
    -ScriptVarName 'UDSelectedGroup'

$script:Window.FindName('BtnOpUserDevices').Add_Click({
    $script:Window.FindName('TxtUDGroupSearch').Text            = ''
    $script:Window.FindName('TxtUDUpnList').Text                = ''
    $script:Window.FindName('UDSelectedGroupBadge').Visibility  = 'Collapsed'
    $script:Window.FindName('UDGroupSearchResults').Visibility  = 'Collapsed'
    $script:Window.FindName('ChkUDWindows').IsChecked           = $true
    $script:Window.FindName('ChkUDAndroid').IsChecked           = $true
    $script:Window.FindName('ChkUDiOS').IsChecked               = $true
    $script:Window.FindName('ChkUDMacOS').IsChecked             = $true
    $script:Window.FindName('CmbUDOwnership').SelectedIndex     = 0
    $script:Window.FindName('ChkUDIntuneOnly').IsChecked        = $false
    $script:UDSelectedGroup = $null
    Show-Panel 'PanelUserDevices'
    Hide-Notification
    Write-VerboseLog 'Panel: Add User Devices to Group' -Level Info
})

$script:Window.FindName('BtnUDValidate').Add_Click({
    Hide-Notification

    if ($null -eq $script:UDSelectedGroup) {
        Show-Notification 'Please select a target group first.' -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }
    $upnTxt = $script:Window.FindName('TxtUDUpnList').Text
    if ([string]::IsNullOrWhiteSpace($upnTxt)) {
        Show-Notification 'Please enter at least one user UPN.' -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }

    # Collect filter selections on the UI thread before handing off to background
    $selPlatforms = [System.Collections.Generic.List[string]]::new()
    if ($script:Window.FindName('ChkUDWindows').IsChecked) { $selPlatforms.Add('Windows') }
    if ($script:Window.FindName('ChkUDAndroid').IsChecked) { $selPlatforms.Add('Android') }
    if ($script:Window.FindName('ChkUDiOS').IsChecked)     { $selPlatforms.Add('iOS') }
    if ($script:Window.FindName('ChkUDMacOS').IsChecked)   { $selPlatforms.Add('macOS') }
    if ($selPlatforms.Count -eq 0) {
        Show-Notification 'Please select at least one device platform.' -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }

    $ownershipSel    = $script:Window.FindName('CmbUDOwnership').SelectedItem.Content
    $ownershipFilter = switch ($ownershipSel) {
        'Company only'  { 'Company' }
        'Personal only' { 'Personal' }
        default         { 'All' }
    }

    $intuneOnly = [bool]$script:Window.FindName('ChkUDIntuneOnly').IsChecked

    $script:UDParams = @{
        UpnTxt          = $upnTxt
        Platforms       = $selPlatforms
        OwnershipFilter = $ownershipFilter
        IntuneOnly      = $intuneOnly
    }
    $btn = $script:Window.FindName('BtnUDValidate')

    Invoke-OnBackground -DisableButton $btn -BusyMessage '--- Add User Devices: Validating ---' -Work {
        $params          = $script:UDParams
        $upns            = $params['UpnTxt'] -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
        $platforms       = $params['Platforms']
        $ownershipFilter = $params['OwnershipFilter']
        $intuneOnly      = $params['IntuneOnly']

        $validDevices   = [System.Collections.Generic.List[PSCustomObject]]::new()
        $invalidUpns    = [System.Collections.Generic.List[string]]::new()
        $skippedDevices = [System.Collections.Generic.List[string]]::new()

        Write-VerboseLog "Target group: $($script:UDSelectedGroup.DisplayName) ($($script:UDSelectedGroup.Id))" -Level Info
        Write-VerboseLog "Platform filter : $($platforms -join ', ')" -Level Info
        Write-VerboseLog "Ownership filter: $ownershipFilter" -Level Info
        Write-VerboseLog "Processing $($upns.Count) UPN(s)..." -Level Info

        foreach ($upn in $upns) {
            Write-VerboseLog "Resolving user: $upn" -Level Info

            $userId          = $null
            $userDisplayName = $upn
            try {
                $userObj         = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/users/$([Uri]::EscapeDataString($upn))?`$select=id,displayName"
                $userId          = $userObj['id']
                $userDisplayName = $userObj['displayName']
                Write-VerboseLog "  $([char]0x2713) User found: $userDisplayName ($userId)" -Level Success
            } catch {
                Write-VerboseLog "  $([char]0x2717) User not found: $upn" -Level Warning
                $invalidUpns.Add($upn)
                continue
            }

            # ── Choose data source based on IntuneOnly filter ──────────────────
            # IntuneOnly = $true  → Intune managedDevices endpoint:
#                 - Only returns devices enrolled in Intune (filter is implicit)
#                 - operatingSystem values are clean: iOS, Windows, Android, macOS
            # IntuneOnly = $false → Entra ownedDevices endpoint:
#                 - Returns all Entra-registered devices (enrolled or not)
#                 - operatingSystem strings vary: iPhone, iPadOS, MacMDM, AndroidForWork, etc.
#                    -  normalised below before comparison.
            if ($intuneOnly) {
                # Intune endpoint  -  requires DeviceManagementManagedDevices.Read.All
                $deviceUri = "https://graph.microsoft.com/v1.0/users/$userId/managedDevices" +
                             "?`$select=id,deviceName,operatingSystem,managedDeviceOwnerType,azureADDeviceId`&`$top=999"
            } else {
                # Entra endpoint  -  requires Device.Read.All (already consented)
                $deviceUri = "https://graph.microsoft.com/v1.0/users/$userId/ownedDevices" +
                             "?`$select=id,displayName,operatingSystem,deviceOwnership`&`$top=999"
            }

            $devicesRaw = @()
            try {
                $devResp    = Invoke-MgGraphRequest -Method GET -Uri $deviceUri
                $devicesRaw = @($devResp['value'])
            } catch {
                Write-VerboseLog "  [!] Could not fetch devices for $upn : $($_.Exception.Message)" -Level Warning
                continue
            }

            if ($devicesRaw.Count -eq 0) {
                Write-VerboseLog "  No owned devices found for $upn" -Level Info
                continue
            }
            $srcLabel = if ($intuneOnly) { 'Intune' } else { 'Entra' }
            Write-VerboseLog "  Found $($devicesRaw.Count) device(s) from $srcLabel  -  applying filters..." -Level Info

            foreach ($dev in $devicesRaw) {
                # Field names differ between the two endpoints
                if ($intuneOnly) {
                    $devId        = $dev['id']
                    $devName      = $dev['deviceName']
                    $devOS        = $dev['operatingSystem']          # already clean: iOS/Windows/Android/macOS
                    $devOwnership = $dev['managedDeviceOwnerType']   # 'company' or 'personal'
                    $mgmtLabel    = 'Intune'
                } else {
                    $devId        = $dev['id']
                    $devName      = $dev['displayName']
                    $devOS        = $dev['operatingSystem']
                    $devOwnership = $dev['deviceOwnership']          # 'Company' or 'Personal'
                    $mgmtLabel    = 'Entra'
                }

                # Platform filter
                # Entra stores OS with varying strings  -  normalise to the four UI values.
                # Intune values are already clean so the switch still matches correctly.
                $osNorm = switch -Wildcard ($devOS) {
                    'Windows*'      { 'Windows' }
                    'Android*'      { 'Android' }
                    'iOS'           { 'iOS' }
                    'iPadOS'        { 'iOS' }    # Entra: iPads enrolled as iPadOS
                    'iPhone'        { 'iOS' }    # Entra: iPhones stored as 'iPhone'
                    'MacMDM'        { 'macOS' }  # Entra: Macs enrolled via MDM profile
                    'Mac OS X'      { 'macOS' }  # Entra: legacy macOS string
                    'macOS'         { 'macOS' }
                    default         { $devOS }
                }
                if ($osNorm -notin $platforms) {
                    $skippedDevices.Add(($devName + ' (' + $devOS + '  -  not in platform filter)'))
                    Write-VerboseLog "    Skipped (platform): $devName  [$devOS]" -Level Info
                    continue
                }

                # Ownership filter
                # Intune uses lowercase ('company'/'personal'); Entra uses title-case ('Company'/'Personal').
                # Normalise both to title-case for the comparison.
                $ownerNorm = if ($devOwnership) {
                    [System.Globalization.CultureInfo]::InvariantCulture.TextInfo.ToTitleCase($devOwnership.ToLower())
                } else { '' }
                if ($ownershipFilter -ne 'All' -and $ownerNorm -ne $ownershipFilter) {
                    $skippedDevices.Add("$devName (ownership: $devOwnership  -  filter: $ownershipFilter)")
                    Write-VerboseLog "    Skipped (ownership): $devName  [$devOwnership]" -Level Info
                    continue
                }
                $_devSummary = $osNorm + ' | ' + $ownerNorm + ' | ' + $mgmtLabel + ' | owner: ' + $userDisplayName
                Write-VerboseLog "    $([char]0x2713) Matched: $devName  [$_devSummary]" -Level Success
                # ── For Intune devices: resolve the Entra directory object ID ────────────────────────
                # Intune managedDevices returns its own managed-device ID, which is NOT a valid Entra
                # directory object ID. We use azureADDeviceId to look up the Entra device object ID.
                if ($intuneOnly) {
                    $aadDevId = $dev['azureADDeviceId']
                    if ([string]::IsNullOrEmpty($aadDevId)) {
                        Write-VerboseLog "    [!] No azureADDeviceId on Intune record for $devName — skipping." -Level Warning
                        $skippedDevices.Add("$devName (Intune record missing azureADDeviceId)")
                        continue
                    }
                    try {
                        $entraLookup = Invoke-MgGraphRequest -Method GET `
                            -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$aadDevId'&`$select=id"
                        $entraObjId = $entraLookup['value'][0]['id']
                        if ([string]::IsNullOrEmpty($entraObjId)) {
                            Write-VerboseLog "    [!] Entra device not found for azureADDeviceId=$aadDevId ($devName) — skipping." -Level Warning
                            $skippedDevices.Add("$devName (device not found in Entra directory)")
                            continue
                        }
                        $devId = $entraObjId   # override with Entra object ID for group membership add
                        Write-VerboseLog "    Resolved Entra object ID for $devName : $devId" -Level Info
                    } catch {
                        Write-VerboseLog "    [!] Entra ID lookup failed for $devName : $($_.Exception.Message) — skipping." -Level Warning
                        $skippedDevices.Add("$devName (Entra ID resolution failed: $($_.Exception.Message))")
                        continue
                    }
                }

                $validDevices.Add([PSCustomObject]@{
                    Id          = $devId
                    DisplayName = "$devName  [$_devSummary]"
                    Type        = 'Device'
                    Found       = $true
                    Original    = $devName
                })
            }
        }

        $script:UDValidated = @{
            ValidDevices   = $validDevices
            InvalidUpns    = $invalidUpns
            SkippedDevices = $skippedDevices
        }

    } -Done {
        $validDevices   = $script:UDValidated['ValidDevices']
        $invalidUpns    = $script:UDValidated['InvalidUpns']
        $skippedDevices = $script:UDValidated['SkippedDevices']
        $params         = $script:UDParams

        if ($validDevices.Count -eq 0) {
            $msg = 'No devices matched the selected filters. Check UPNs, platform and ownership selections.'
            Write-VerboseLog $msg -Level Warning
            Show-Notification $msg -BgColor '#FFF3CD' -FgColor '#7A4800'
            return
        }

        $extraLines = [System.Collections.Generic.List[string]]::new()
        $extraLines.Add("Platform filter  : $($params['Platforms'] -join ', ')")
        $extraLines.Add("Ownership filter : $($params['OwnershipFilter'])")
        $extraLines.Add("Intune only      : $(if ($params['IntuneOnly']) { 'Yes' } else { 'No' })")
        if ($invalidUpns.Count -gt 0)    { $extraLines.Add("Users not found  : $($invalidUpns.Count) UPN(s) could not be resolved") }
        if ($skippedDevices.Count -gt 0) { $extraLines.Add("Devices skipped (filter mismatch): $($skippedDevices.Count)") }

        $invalidForDlg = [System.Collections.Generic.List[string]]::new()
        foreach ($u in $invalidUpns) { $invalidForDlg.Add($u) }

        $confirmed = Show-ConfirmationDialog `
            -Title 'Add User Devices to Group' -OperationLabel 'Add User Devices to Group' `
            -TargetGroup $script:UDSelectedGroup.DisplayName `
            -TargetGroupId $script:UDSelectedGroup.Id `
            -ValidObjects $validDevices -InvalidEntries $invalidForDlg `
            -ExtraInfo ($extraLines -join "`n")

        if (-not $confirmed) { Write-VerboseLog 'Action cancelled by user.' -Level Warning; return }

        $btn = $script:Window.FindName('BtnUDValidate')
        Invoke-OnBackground -DisableButton $btn -BusyMessage '--- Adding devices to group ---' -Work {
            $devices = $script:UDValidated['ValidDevices']
            $groupId = $script:UDSelectedGroup.Id
            $ok = 0; $fail = 0

            foreach ($dev in $devices) {
                if ($null -ne $Shared -and $Shared['StopRequested']) {
                    Write-VerboseLog 'Stop requested  -  halting.' -Level Warning
                    break
                }
                Write-VerboseLog "Adding device: $($dev.Original)..." -Level Action
                try {
                    $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($dev.Id)" }
                    Invoke-MgGraphRequest -Method POST `
                        -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/members/`$ref" `
                        -Body ($body | ConvertTo-Json) -ContentType 'application/json'
                    Write-VerboseLog "  $([char]0x2713) Added: $($dev.Original)" -Level Success
                    $ok++
                } catch {
                    $errMsg = $_.Exception.Message
                    if ($errMsg -match 'already exist') {
                        Write-VerboseLog "  Already a member: $($dev.Original)" -Level Warning
                    } else {
                        Write-VerboseLog "  $([char]0x2717) Failed: $($dev.Original)  -  $errMsg" -Level Error
                        $fail++
                    }
                }
            }
            $script:UDExecResult = @{ Ok = $ok; Fail = $fail }

        } -Done {
            $ok   = $script:UDExecResult['Ok']
            $fail = $script:UDExecResult['Fail']
            $msg  = "Add User Devices complete. Added: $ok | Failed: $fail"
            $vLevel  = if ($fail -gt 0) { 'Warning' } else { 'Success' }
            $bgColor = if ($fail -gt 0) { '#FFF3CD' } else { '#D4EDDA' }
            $fgColor = if ($fail -gt 0) { '#7A4800' } else { '#155724' }
            Write-VerboseLog $msg -Level $vLevel
            Show-Notification $msg -BgColor $bgColor -FgColor $fgColor
        }
    }
})





# ── GET POLICY ASSIGNMENTS ───────────────────────────────────────────────────

# ── Sub-type normalisation ──
function Get-GPASubType {
    param([string]$PolicyCollection, [string]$ODataType, [string]$TemplateDisplayName)
    switch ($PolicyCollection) {
        'configurationPolicies' {
            if (-not [string]::IsNullOrEmpty($TemplateDisplayName)) {
                switch -Wildcard ($TemplateDisplayName) {
                    '*Antivirus exclusion*'              { return 'Microsoft Defender Antivirus Exclusions' }
                    '*Defender Antivirus*'               { return 'Microsoft Defender Antivirus' }
                    '*Defender Update*'                  { return 'Defender Update Controls' }
                    '*Attack Surface*'                   { return 'Attack Surface Reduction Rules' }
                    '*BitLocker*'                        { return 'BitLocker' }
                    '*Windows Firewall Rules*'           { return 'Windows Firewall Rules' }
                    '*Windows Firewall*'                 { return 'Windows Firewall' }
                    '*MDE*onboard*'                      { return 'MDE Onboarding Policy' }
                    '*Exploit Protection*'               { return 'Exploit Protection' }
                    '*App Control*'                      { return 'App Control for Business' }
                    '*Windows Security Experience*'      { return 'Windows Security Experience' }
                    '*Identity Protection*'              { return 'Identity Protection' }
                    '*Local user group*'                 { return 'Local User Group Membership' }
                    '*LAPS*'                             { return 'Local Admin Password Solution (Windows LAPS)' }
                    '*Health Monitoring*'                { return 'Health Monitoring' }
                    '*Windows operating system recovery*'{ return 'Windows OS Recovery' }
                    '*Update ring*'                      { return 'Update Rings' }
                    '*Software update*'                  { return 'Software Update' }
                    '*Endpoint detection*'               { return 'Endpoint Protection' }
                    '*Extension*'                        { return 'Extensions' }
                    default                               { return $TemplateDisplayName }
                }
            } else { return 'Settings Catalog' }
        }
        'deviceConfigurations' {
            $c = $ODataType -replace '^#microsoft\.graph\.', ''
            switch -Wildcard ($c) {
                '*trustedRoot*'          { return 'Trusted Certificate' }
                '*scepCertificate*'      { return 'SCEP Certificate' }
                '*pkcsCertificate*'      { return 'Certificate' }
                '*pkcs12Import*'         { return 'Certificate' }
                '*vpn*'                  { return 'VPN' }
                '*wifiConfiguration*'    { return 'Wi-Fi' }
                '*wiFiConfiguration*'    { return 'Wi-Fi' }
                '*wifi*'                 { return 'Wi-Fi' }
                '*sharedPC*'             { return 'Shared PC' }
                '*editionUpgrade*'       { return 'Edition Upgrade' }
                '*kiosk*'               { return 'Kiosk' }
                '*updateForBusiness*'    { return 'Update Rings' }
                '*softwareUpdate*'       { return 'Software Update' }
                '*endpointProtection*'   { return 'Endpoint Protection' }
                '*bitLocker*'            { return 'BitLocker' }
                '*deviceFirmware*'       { return 'Device Features' }
                '*deviceFeatures*'       { return 'Device Features' }
                '*windowsHealthMonitoring*' { return 'Health Monitoring' }
                '*localAdminPassword*'   { return 'Local Admin Password Solution (Windows LAPS)' }
                '*custom*'              { return 'Custom' }
                '*general*'             { return 'General Device' }
                default                  { return $c }
            }
        }
        'groupPolicyConfigurations'     { return 'Administrative Templates' }
        'deviceManagementScripts'       { return 'PowerShell' }
        'deviceHealthScripts'           { return 'Remediation' }
        'windowsAutopilotDeploymentProfiles' {
            $c = $ODataType -replace '^#microsoft\.graph\.', ''
            if ($c -match 'activeDirectory|azureAD') { return 'Autopilot' }
            else { return 'Autopilot' }
        }
        'configurationPolicies_devicePrep' { return 'Autopilot (Device Prep)' }
        'deviceEnrollmentConfigurations' { return 'Autopilot ESP' }
        'deviceCompliancePolicies'       { return '' }
        default                          { return '' }
    }
}

# ── OS Platform normalisation ──
function Get-GPAPlatform {
    param([string]$PolicyCollection, [string]$ODataType, [string]$Platforms)
    switch ($PolicyCollection) {
        'configurationPolicies' {
            switch -Wildcard ($Platforms) {
                'windows10'    { return 'Windows' }
                'windows*'     { return 'Windows' }
                'iOS'          { return 'iOS/iPadOS' }
                'macOS'        { return 'macOS' }
                'android'      { return 'Android' }
                'androidAOSP'  { return 'AOSP' }
                'linux'        { return 'Linux' }
                default         { return $Platforms }
            }
        }
        'deviceConfigurations' {
            $c = $ODataType -replace '^#microsoft\.graph\.', ''
            switch -Wildcard ($c) {
                'windows*'         { return 'Windows' }
                'sharedPC*'        { return 'Windows' }
                'editionUpgrade*'  { return 'Windows' }
                'windows81*'       { return 'Windows' }
                'windowsPhone*'    { return 'Windows' }
                'ios*'             { return 'iOS/iPadOS' }
                'macOS*'           { return 'macOS' }
                'android*'         { return 'Android' }
                'aosp*'            { return 'AOSP' }
                'linux*'           { return 'Linux' }
                default             { return '' }
            }
        }
        'groupPolicyConfigurations'     { return 'Windows' }
        'deviceManagementScripts'       { return 'Windows' }
        'deviceHealthScripts'           { return 'Windows' }
        'windowsAutopilotDeploymentProfiles' { return 'Windows' }
        'deviceEnrollmentConfigurations'        { return 'Windows' }
        'configurationPolicies_devicePrep' { return 'Windows' }
        'deviceCompliancePolicies' {
            $c = $ODataType -replace '^#microsoft\.graph\.', ''
            switch -Wildcard ($c) {
                'windows*'         { return 'Windows' }
                'ios*'             { return 'iOS/iPadOS' }
                'macOS*'           { return 'macOS' }
                'android*'         { return 'Android' }
                'aosp*'            { return 'AOSP' }
                default            { return '' }
            }
        }
        default { return '' }
    }
}

# ── Panel open ──
$script:Window.FindName('BtnOpGetPolicyAssignments').Add_Click({
    $script:Window.FindName('TxtGPANameFilter').Text            = ''
    $script:Window.FindName('CmbGPAType').SelectedIndex         = 0
    $script:Window.FindName('CmbGPASubType').SelectedIndex      = 0
    $script:Window.FindName('CmbGPAPlatform').SelectedIndex     = 0
    $script:Window.FindName('PnlGPAProgress').Visibility        = 'Collapsed'
    $script:Window.FindName('PnlGPAResults').Visibility         = 'Collapsed'
    $script:Window.FindName('TxtGPANoResults').Visibility       = 'Collapsed'
    foreach ($b in @('BtnGPACopyValue','BtnGPACopyRow','BtnGPACopyAll','BtnGPAExportXlsx')) {
        $script:Window.FindName($b).IsEnabled = $false
    }
    $dg = $script:Window.FindName('DgGPAResults'); if ($dg) { $dg.ItemsSource = $null }
    $script:GPAResult = $null
    Show-Panel 'PanelGetPolicyAssignments'
    Hide-Notification
    Write-VerboseLog 'Panel: Get Policy Info' -Level Info
})

# ── Shared launch function (called by both Get Selected and Get All buttons) ──
function Start-GPAQuery {
    param([bool]$ApplyFilters)

    Hide-Notification
    $script:Window.FindName('PnlGPAResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtGPANoResults').Visibility = 'Collapsed'
    foreach ($b in @('BtnGPACopyValue','BtnGPACopyRow','BtnGPACopyAll','BtnGPAExportXlsx')) {
        $script:Window.FindName($b).IsEnabled = $false
    }
    $dg = $script:Window.FindName('DgGPAResults'); if ($dg) { $dg.ItemsSource = $null }

    $nameKw   = if ($ApplyFilters) { $script:Window.FindName('TxtGPANameFilter').Text.Trim() } else { '' }
    $typeF    = if ($ApplyFilters) { $script:Window.FindName('CmbGPAType').SelectedItem.Content }  else { '(All types)' }
    $subTypeF = if ($ApplyFilters) { $script:Window.FindName('CmbGPASubType').SelectedItem.Content } else { '(All sub-types)' }
    $platF    = if ($ApplyFilters) { $script:Window.FindName('CmbGPAPlatform').SelectedItem.Content } else { '(All platforms)' }

    $script:GPAParams = @{
        NameKeyword = $nameKw
        TypeFilter  = $typeF
        SubTypeFilter = $subTypeF
        PlatFilter  = $platF
    }

    $script:Window.FindName('PnlGPAProgress').Visibility  = 'Visible'
    $script:Window.FindName('TxtGPAProgressMsg').Text     = 'Loading policies...'
    $script:Window.FindName('TxtGPAProgressDetail').Text  = ''

    foreach ($b in @('BtnGPAGetSelected','BtnGPAGetAll')) {
        $script:Window.FindName($b).IsEnabled = $false
    }
    Write-VerboseLog ('--- Get Policy Info: starting (ApplyFilters=' + $ApplyFilters + ') ---') -Level Action

    Invoke-OnBackground -BusyMessage '--- Get Policy Info ---' -Work {

        $params     = $script:GPAParams
        $nameKw     = $params['NameKeyword']
        $typeFilter = $params['TypeFilter']
        $subFilter  = $params['SubTypeFilter']
        $platFilter = $params['PlatFilter']

        # ── Helper: page through Graph collection ──────────────────────────────
        function GPA-Page { param([string]$Uri)
            $all = [System.Collections.Generic.List[object]]::new()
            $next = $Uri
            while ($next) {
                $r = Invoke-MgGraphRequest -Method GET -Uri $next
                if ($r['value']) { foreach ($v in @($r['value'])) { $all.Add($v) } }
                $next = $r['@odata.nextLink']
            }
            return $all
        }

        # ── Load assignment filters lookup ─────────────────────────────────────
        Write-VerboseLog 'GPA: Loading assignment filters...' -Level Info
        $filterLookup = @{}
        try {
            $filters = GPA-Page -Uri 'https://graph.microsoft.com/beta/deviceManagement/assignmentFilters?$select=id,displayName'
            foreach ($f in $filters) { if ($f['id']) { $filterLookup[$f['id']] = $f['displayName'] } }
            Write-VerboseLog "GPA: $($filterLookup.Count) assignment filter(s) loaded" -Level Info
        } catch { Write-VerboseLog "GPA: Could not load assignment filters: $($_.Exception.Message)" -Level Warning }

        # ── Group display name cache ───────────────────────────────────────────
        $groupCache = @{}
        function GPA-GroupName { param([string]$GroupId)
            if (-not $GroupId) { return '' }
            if ($groupCache.ContainsKey($GroupId)) { return $groupCache[$GroupId] }
            try {
                $enc = [Uri]::EscapeDataString($GroupId)
                $g = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/${enc}?`$select=id,displayName"
                $groupCache[$GroupId] = $g['displayName']; return $g['displayName']
            } catch { $groupCache[$GroupId] = $GroupId; return $GroupId }
        }

        # ── Decide which collections to query ──────────────────────────────────
        $collections = @(
            @{ Key='deviceConfigurations';              Label='Device Configuration';                  Uri='https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?$select=id,displayName' }
            @{ Key='configurationPolicies';             Label='Configuration Policy';                  Uri='https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$select=id,name,platforms,technologies,templateReference' }
            @{ Key='groupPolicyConfigurations';         Label='Administrative Templates';              Uri='https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations?$select=id,displayName' }
            @{ Key='deviceManagementScripts';           Label='Platform Script (Windows PowerShell)'; Uri='https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts?$select=id,displayName' }
            @{ Key='deviceHealthScripts';               Label='Detection/Remediation Script';         Uri='https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts?$select=id,displayName' }
            @{ Key='windowsAutopilotDeploymentProfiles';Label='Autopilot Profile';               Uri='https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles?$select=id,displayName,description,outOfBoxExperienceSetting' }
            @{ Key='configurationPolicies_devicePrep'; AssignKey='configurationPolicies'; Label='Autopilot Profile'; Uri='https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$filter=(technologies has ''enrollment'')&$select=id,name,description,platforms,technologies,templateReference' }
            @{ Key='deviceEnrollmentConfigurations';    Label='Autopilot Profile';               Uri='https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations?$select=id,displayName,priority,description' }
            @{ Key='deviceCompliancePolicies';          Label='Compliance Policy';                    Uri='https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies?$select=id,displayName' }
        )

        # Skip collections not matching the type filter
        if ($typeFilter -ne '(All types)') {
            $collections = $collections | Where-Object { $_['Label'] -eq $typeFilter }
        }

        # ── Platform-based collection pre-filter (skip Windows-only endpoints) ──
        $winOnlyKeys = @('groupPolicyConfigurations','deviceManagementScripts',
                         'deviceHealthScripts','windowsAutopilotDeploymentProfiles','configurationPolicies_devicePrep','deviceEnrollmentConfigurations')
        if ($platFilter -ne '(All platforms)' -and $platFilter -ne 'Windows') {
            $before = @($collections).Count
            $skippedNames = @($collections | Where-Object { $winOnlyKeys -contains $_['Key'] }) | ForEach-Object { $_['Label'] }
            $collections  = @($collections | Where-Object { $winOnlyKeys -notcontains $_['Key'] })
            if ($skippedNames.Count -gt 0) {
                Write-VerboseLog ("GPA: Skipped " + $skippedNames.Count + " Windows-only collection(s): " + ($skippedNames -join ', ')) -Level Info
            }
        }

        # ── Server-side platform filter for configurationPolicies ──
        $platApiMap = @{ 'Windows'='windows10'; 'macOS'='macOS'; 'iOS/iPadOS'='iOS';
                         'Android'='android'; 'AOSP'='androidAOSP'; 'Linux'='linux' }
        if ($platFilter -ne '(All platforms)' -and $platApiMap.ContainsKey($platFilter)) {
            $apiPlat = $platApiMap[$platFilter]
            foreach ($c in $collections) {
                if ($c['Key'] -eq 'configurationPolicies') {
                    $c['OriginalUri'] = $c['Uri']
                    $c['Uri'] += "&`$filter=platforms has '$apiPlat'"
                    Write-VerboseLog "GPA: Server-side platform filter added for Configuration Policy ($apiPlat)" -Level Info
                    break
                }
            }
        }

        $rows = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($col in $collections) {
            Write-VerboseLog "GPA: Loading $($col['Label'])..." -Level Info
            $policies = @()
            try { $policies = @(GPA-Page -Uri $col['Uri']) }
            catch {
                if ($col['OriginalUri']) {
                    Write-VerboseLog "GPA: Server-side filter failed, falling back to full list..." -Level Warning
                    try { $policies = @(GPA-Page -Uri $col['OriginalUri']) }
                    catch { Write-VerboseLog "GPA: Failed to load $($col['Key']): $($_.Exception.Message)" -Level Warning; continue }
                } else { Write-VerboseLog "GPA: Failed to load $($col['Key']): $($_.Exception.Message)" -Level Warning; continue }
            }
            Write-VerboseLog "GPA: $($policies.Count) $($col['Label']) policy/policies found" -Level Info

            $pIdx = 0
            foreach ($policy in $policies) {
                $pIdx++

                # Get display name (field differs per endpoint)

                # For deviceEnrollmentConfigurations, only process ESP (windows10EnrollmentCompletionPageConfiguration)
                if ($col['Key'] -eq 'deviceEnrollmentConfigurations') {
                    $odType = [string]$policy['@odata.type']
                    if ($odType -notmatch 'windows10EnrollmentCompletionPageConfiguration') { continue }
                }
                $pName = if ($col['Key'] -eq 'configurationPolicies') { [string]$policy['name'] }
                         else { [string]$policy['displayName'] }

                # Keyword filter
                if ($nameKw -and ($pName -notlike "*$nameKw*")) { continue }

                # Derive sub-type and platform
                $odataType  = [string]$policy['@odata.type']
                $tmplName   = if ($policy['templateReference']) { [string]$policy['templateReference']['templateDisplayName'] } else { '' }
                $platforms  = [string]$policy['platforms']

                # Autopilot-specific: extract description and deployment mode
                $apDescription = ''
                $apDeployMode  = ''
                $espPriority   = ''
                if ($col['Key'] -eq 'windowsAutopilotDeploymentProfiles') {
                    $apDescription = if ($policy['description']) { [string]$policy['description'] } else { '' }
                    $oobe = $policy['outOfBoxExperienceSetting']
                    if ($oobe) {
                        $ut = [string]$oobe['userType']
                        $apDeployMode = switch ($ut) {
                            'standard'       { 'User-Driven' }
                            'administrator'  { 'User-Driven (Admin)' }
                            'standardUser'   { 'User-Driven' }
                            default          { $ut }
                        }
                    }
                    $odtShort = $odataType -replace '^#microsoft\.graph\.', ''
                    if ($odtShort -match 'azureAD' -and [string]::IsNullOrEmpty($apDeployMode)) {
                        $apDeployMode = 'Self-Deploying'
                    }
                    Write-VerboseLog ("GPA: AP profile '" + $pName + "' | type=" + $odataType + " | mode=" + $apDeployMode) -Level Info
                }
                # Device Prep: extract description from configurationPolicies
                if ($col['Key'] -eq 'configurationPolicies_devicePrep') {
                    $apDescription = if ($policy['description']) { [string]$policy['description'] } else { '' }
                    $apDeployMode  = 'Device Prep'
                    Write-VerboseLog ('GPA: Device Prep profile ''' + $pName + '''') -Level Info
                }
                # ESP: extract priority and description
                $espPriority = ''
                if ($col['Key'] -eq 'deviceEnrollmentConfigurations') {
                    $espPriority = if ($null -ne $policy['priority']) { [string]$policy['priority'] } else { '' }
                    $apDescription = if ($policy['description']) { [string]$policy['description'] } else { '' }
                    Write-VerboseLog ('GPA: ESP profile ''' + $pName + ''' | priority=' + $espPriority) -Level Info
                }
                $subType    = Get-GPASubType -PolicyCollection $col['Key'] -ODataType $odataType -TemplateDisplayName $tmplName
                $platform   = Get-GPAPlatform -PolicyCollection $col['Key'] -ODataType $odataType -Platforms $platforms

                # Sub-type filter
                if ($subFilter -ne '(All sub-types)' -and $subType -ne $subFilter) { continue }
                # Platform filter
                if ($platFilter -ne '(All platforms)' -and $platform -ne $platFilter) { continue }

                # Get assignments for this policy
                $assignKey  = if ($col['AssignKey']) { $col['AssignKey'] } else { $col['Key'] }
                $assignUri = "https://graph.microsoft.com/beta/deviceManagement/$($assignKey)/$($policy['id'])/assignments"
                $assignments = @()
                try { $assignments = @(GPA-Page -Uri $assignUri) }
                catch { Write-VerboseLog "GPA: Could not get assignments for '$pName': $($_.Exception.Message)" -Level Warning; continue }

                if ($assignments.Count -eq 0) {
                    $rows.Add([PSCustomObject]@{
                        PolicyName   = $pName
                        PolicyType   = $col['Label']
                        PolicySubType = $subType
                        OSPlatform   = $platform
                        IncludedGroup = ''
                        ExcludedGroup = ''
                        IncludeFilter = ''
                        ExcludeFilter = ''
                        Description   = $apDescription
                        DeploymentMode = $apDeployMode
                        Priority       = $espPriority
                    }); continue
                }

                foreach ($asgn in $assignments) {
                    $target    = $asgn['target']
                    if (-not $target) { continue }

                    $otype     = [string]$target['@odata.type']
                    $isExclude = $otype -match 'exclusionGroup'
                    $groupId   = if ($target['groupId']) { [string]$target['groupId'] } else { [string]$target['entraObjectId'] }

                    $targetName = ''
                    if ($groupId) {
                        $targetName = GPA-GroupName -GroupId $groupId
                    } elseif ($otype -match 'allDevices') {
                        $targetName = '[All devices]'
                    } elseif ($otype -match 'allLicensedUsers|allUsers') {
                        $targetName = '[All users]'
                    } else { continue }

                    $filterId   = [string]$target['deviceAndAppManagementAssignmentFilterId']
                    $filterType = [string]$target['deviceAndAppManagementAssignmentFilterType']
                    $filterName = if ($filterId -and $filterLookup.ContainsKey($filterId)) { $filterLookup[$filterId] } else { $filterId }

                    $rows.Add([PSCustomObject]@{
                        PolicyName    = $pName
                        PolicyType    = $col['Label']
                        PolicySubType = $subType
                        OSPlatform    = $platform
                        IncludedGroup = if (-not $isExclude) { $targetName } else { '' }
                        ExcludedGroup = if ($isExclude) { $targetName } else { '' }
                        IncludeFilter = if ($filterType -eq 'include') { $filterName } else { '' }
                        ExcludeFilter = if ($filterType -eq 'exclude') { $filterName } else { '' }
                        Description    = $apDescription
                        DeploymentMode = $apDeployMode
                        Priority       = $espPriority
                    })
                }
            }
        }

        $script:GPAResult = $rows
        Write-VerboseLog "GPA: Complete. $($rows.Count) assignment row(s) returned." -Level Info

    } -Done {

        foreach ($b in @('BtnGPAGetSelected','BtnGPAGetAll')) {
            $script:Window.FindName($b).IsEnabled = $true
        }
        $script:Window.FindName('PnlGPAProgress').Visibility = 'Collapsed'

        $rows = $script:GPAResult
        if (-not $rows -or $rows.Count -eq 0) {
            $script:Window.FindName('TxtGPANoResults').Visibility = 'Visible'
            Write-VerboseLog 'GPA: No results to display.' -Level Warning
            return
        }

        $sortedArr = @($rows | Sort-Object PolicyName, PolicyType, IncludedGroup, ExcludedGroup)
        $sortedList = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($r in $sortedArr) { $sortedList.Add($r) }
        $dg = $script:Window.FindName('DgGPAResults')
        $dg.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSCustomObject]]::new($sortedList)

        $script:Window.FindName('TxtGPACount').Text = "$($rows.Count) assignment row(s)"
        $script:Window.FindName('BtnGPACopyAll').IsEnabled    = $true
        $script:Window.FindName('TxtGPAFilter').IsEnabled  = $true
        $script:Window.FindName('BtnGPAFilter').IsEnabled  = $true
        $script:Window.FindName('BtnGPAFilterClear').IsEnabled = $true
        $script:Window.FindName('BtnGPAExportXlsx').IsEnabled = $true
        $script:Window.FindName('PnlGPAResults').Visibility   = 'Visible'

        # Enable Copy Value / Copy Row on selection
        $dg.Add_SelectedCellsChanged({
            $selCells = $script:Window.FindName('DgGPAResults').SelectedCells.Count
            $script:Window.FindName('BtnGPACopyValue').IsEnabled = ($selCells -ge 1)
            $script:Window.FindName('BtnGPACopyRow').IsEnabled   = ($selCells -gt 0)
        })
    }
}

$script:Window.FindName('BtnGPAGetSelected').Add_Click({ Start-GPAQuery -ApplyFilters $true })
$script:Window.FindName('BtnGPAGetAll').Add_Click(     { Start-GPAQuery -ApplyFilters $false })

# ── Sub-Type filter: grey-out Policy Type when a specific sub-type is chosen ──
$script:Window.FindName('CmbGPASubType').Add_SelectionChanged({
    $sub  = $script:Window.FindName('CmbGPASubType')
    $type = $script:Window.FindName('CmbGPAType')
    if ($sub.SelectedIndex -le 0) {
        # Back to "All sub-types" — restore Policy Type filter
        $type.IsEnabled = $true
    } else {
        # Specific sub-type chosen — clear and disable Policy Type to avoid contradictions
        $type.SelectedIndex = 0
        $type.IsEnabled     = $false
    }
})

# ── Copy / Export buttons ──
# ── BtnGPAFilter ──
$script:Window.FindName('BtnGPAFilter').Add_Click({
    $dg      = $script:Window.FindName('DgGPAResults')
    $keyword = $script:Window.FindName('TxtGPAFilter').Text.Trim()
    if ($null -eq $script:GPAAllData) {
        $script:GPAAllData = @($dg.ItemsSource)
    }
    if ($keyword -eq '') {
        $dg.ItemsSource = $script:GPAAllData
        Show-Notification "Filter cleared - $($script:GPAAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    } else {
        $filtered = @($script:GPAAllData | Where-Object {
            $row = $_
            $match = $false
            foreach ($prop in $row.PSObject.Properties) {
                if ([string]$prop.Value -like "*$keyword*") { $match = $true; break }
            }
            $match
        })

# ── BtnGPAFilterClear ──
$script:Window.FindName('BtnGPAFilterClear').Add_Click({
    $script:Window.FindName('TxtGPAFilter').Text = ''
    $dg = $script:Window.FindName('DgGPAResults')
    if ($null -ne $script:GPAAllData) {
        $dg.ItemsSource = $script:GPAAllData
        Show-Notification "Filter cleared - $($script:GPAAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    }
})
        $dg.ItemsSource = $filtered
        Show-Notification "$($filtered.Count) of $($script:GPAAllData.Count) result(s) shown after filter." -BgColor '#D4EDDA' -FgColor '#155724'
    }
    Write-VerboseLog "GPA: Filter applied - keyword='$keyword'" -Level Info
})

$script:Window.FindName('BtnGPACopyValue').Add_Click({
    $dg = $script:Window.FindName('DgGPAResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $vals = [System.Collections.Generic.List[string]]::new()
    foreach ($cellInfo in @($dg.SelectedCells)) {
        $col = $cellInfo.Column -as [System.Windows.Controls.DataGridBoundColumn]
        if ($col) { $vals.Add(([string]$cellInfo.Item.($col.Binding.Path.Path)).Trim()) }
    }
    [System.Windows.Clipboard]::SetText($vals -join "`r`n")
    Show-Notification "$($vals.Count) value(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
})

$script:Window.FindName('BtnGPACopyRow').Add_Click({
    $dg   = $script:Window.FindName('DgGPAResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $seen = [System.Collections.Generic.HashSet[object]]::new()
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($cell in @($dg.SelectedCells)) { if ($seen.Add($cell.Item)) { $rows.Add($cell.Item) } }
    $lines = foreach ($r in $rows) {
        "$($r.PolicyName)`t$($r.PolicyType)`t$($r.PolicySubType)`t$($r.OSPlatform)`t$($r.IncludedGroup)`t$($r.ExcludedGroup)`t$($r.IncludeFilter)`t$($r.ExcludeFilter)"
    }
    [System.Windows.Clipboard]::SetText(($lines -join [Environment]::NewLine))
    Show-Notification "Copied $($rows.Count) row(s)." -BgColor '#D4EDDA' -FgColor '#155724'
    # Highlight full rows
    $dg.UnselectAllCells()
    foreach ($r in $rows) { $null = $dg.SelectedItems.Add($r) }
})

$script:Window.FindName('BtnGPACopyAll').Add_Click({
    $dg   = $script:Window.FindName('DgGPAResults')
    $src  = @($dg.ItemsSource)
    $hdr  = "Policy Name`tType`tSub-Type`tOS Platform`tIncluded Group`tExcluded Group`tInclude Filter`tExclude Filter"
    $body = foreach ($r in $src) {
        "$($r.PolicyName)`t$($r.PolicyType)`t$($r.PolicySubType)`t$($r.OSPlatform)`t$($r.IncludedGroup)`t$($r.ExcludedGroup)`t$($r.IncludeFilter)`t$($r.ExcludeFilter)"
    }
    [System.Windows.Clipboard]::SetText(($hdr, ($body -join [Environment]::NewLine) -join [Environment]::NewLine))
    Show-Notification "Copied $($src.Count) row(s) with headers." -BgColor '#D4EDDA' -FgColor '#155724'
})

$script:Window.FindName('BtnGPAExportXlsx').Add_Click({
    $dg  = $script:Window.FindName('DgGPAResults')
    $src = @($dg.ItemsSource)
    if ($src.Count -eq 0) { Show-Notification 'No data to export.' -BgColor '#FFF3CD' -FgColor '#7A4800'; return }

    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title    = "Export Policy Assignments"
    $dlg.Filter   = "Excel Workbook (*.xlsx)|*.xlsx|CSV File (*.csv)|*.csv|All Files (*.*)|*.*"
    $dlg.FileName = "GPAExport_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    $initDir = try { if ($script:ExportOutPath) { Split-Path $script:ExportOutPath -Parent } else { $null } } catch { $null }
    $dlg.InitialDirectory = if ($initDir -and (Test-Path $initDir)) { $initDir } else { [Environment]::GetFolderPath('Desktop') }
    if ($dlg.ShowDialog() -ne $true) { return }
    $outPath = $dlg.FileName

    try {
        $folder = Split-Path $outPath -Parent
        if ($folder -and -not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        $exportRows = $src | Select-Object PolicyName, PolicyType, PolicySubType, OSPlatform, IncludedGroup, ExcludedGroup, IncludeFilter, ExcludeFilter
        if (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue) {
            Import-Module ImportExcel -ErrorAction SilentlyContinue
            $exportRows | Export-Excel -Path $outPath -WorksheetName "PolicyInfo" -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium2
            Show-Notification "Exported $($src.Count) rows: $(Split-Path $outPath -Leaf)" -BgColor '#D4EDDA' -FgColor '#155724'
        } else {
            $exportRows | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8
            Show-Notification "ImportExcel not found - saved as CSV: $(Split-Path $outPath -Leaf)" -BgColor '#FFF3CD' -FgColor '#856404'
        }
        Write-VerboseLog "GPA: Exported $($src.Count) rows to: $outPath" -Level Success
    } catch {
        Show-Notification "Export failed: $($_.Exception.Message)" -BgColor '#F8D7DA' -FgColor '#721C24'
        Write-VerboseLog "GPA: XLSX export error: $($_.Exception.Message)" -Level Error
    }
})


# ── FIND COMMON GROUPS ────────────────────────────────────────────────────

# Update the input label when the input type RadioButton changes
$script:Window.FindName('RbFCUsers').Add_Checked({
    $script:Window.FindName('TxtFCInputLabel').Text = 'USER UPNs  -  one per line'
    $script:Window.FindName('TxtFCInputList').Tag   = 'user1@domain.com'
})
$script:Window.FindName('RbFCDevices').Add_Checked({
    $script:Window.FindName('TxtFCInputLabel').Text = 'DEVICE NAMES or Object IDs  -  one per line'
    $script:Window.FindName('TxtFCInputList').Tag   = 'LAPTOP-001'
})
$script:Window.FindName('RbFCGroups').Add_Checked({
    $script:Window.FindName('TxtFCInputLabel').Text = 'GROUP NAMES or Object IDs  -  one per line'
    $script:Window.FindName('TxtFCInputList').Tag   = 'SG-Finance-All'
})

$script:Window.FindName('BtnOpFindCommon').Add_Click({
    $script:Window.FindName('TxtFCInputList').Text  = ''
    $script:Window.FindName('RbFCUsers').IsChecked  = $true
    $script:Window.FindName('TxtFCInputLabel').Text = 'USER UPNs  -  one per line'
    $script:Window.FindName('TxtFCInputList').Tag   = 'user1@domain.com'
    Show-Panel 'PanelFindCommon'
    Hide-Notification
    Write-VerboseLog 'Panel: Find Common Groups' -Level Info
})

$script:Window.FindName('BtnFCFind').Add_Click({
    Hide-Notification

    $inputTxt = $script:Window.FindName('TxtFCInputList').Text
    if ([string]::IsNullOrWhiteSpace($inputTxt)) {
        Show-Notification 'Please enter at least one identifier.' -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }

    $inputType = if ($script:Window.FindName('RbFCUsers').IsChecked)      { 'Users' }
                 elseif ($script:Window.FindName('RbFCDevices').IsChecked) { 'Devices' }
                 else { 'Groups' }

    $script:FCParams = @{ InputTxt = $inputTxt; InputType = $inputType }
    $btn = $script:Window.FindName('BtnFCFind')
    $script:Window.FindName('PnlFCResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtFCNoResults').Visibility = 'Collapsed'
    $script:Window.FindName('PnlFCProgress').Visibility  = 'Visible'
    $script:Window.FindName('TxtFCProgressMsg').Text     = "Resolving $inputType..."
    $script:Window.FindName('TxtFCProgressDetail').Text  = ''
    $script:Window.FindName('BtnFCCopyValue').IsEnabled  = $false
    $script:Window.FindName('BtnFCCopyRow').IsEnabled    = $false
    $script:Window.FindName('BtnFCCopyAll').IsEnabled    = $false
    $script:Window.FindName('TxtFCFilter').IsEnabled  = $false
    $script:Window.FindName('BtnFCFilter').IsEnabled  = $false
    $script:Window.FindName('BtnFCFilterClear').IsEnabled = $false
    $script:Window.FindName('TxtFCFilter').Text       = ''
    $script:FCAllData = $null
    $script:Window.FindName('BtnFCExportXlsx').IsEnabled = $false

    Invoke-OnBackground -DisableButton $btn -BusyMessage "--- Find Common Groups: resolving $inputType ---" -Work {
        $params    = $script:FCParams
        $inputType = $params['InputType']
        $entries   = $params['InputTxt'] -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }

        Write-VerboseLog "Input type : $inputType" -Level Info
        Write-VerboseLog "Entries    : $($entries.Count)" -Level Info

        if ($entries.Count -lt 2) {
            throw 'Please enter at least two identifiers to find common groups.'
        }

        # ── Step 1: resolve each entry to an Entra Object ID ─────────────────
        $resolvedIds    = [System.Collections.Generic.List[string]]::new()
        $resolvedLabels = [System.Collections.Generic.List[string]]::new()
        $failedEntries  = [System.Collections.Generic.List[string]]::new()
        $isGuidRx = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'

        foreach ($entry in $entries) {
            $objId = $null; $label = $entry
            try {
                switch ($inputType) {
                    'Users' {
                        $u = Invoke-MgGraphRequest -Method GET `
                            -Uri "https://graph.microsoft.com/v1.0/users/$([Uri]::EscapeDataString($entry))?`$select=id,displayName"
                        $objId = $u['id']; $label = $u['displayName']
                    }
                    'Devices' {
                        if ($entry -match $isGuidRx) {
                            $d = Invoke-MgGraphRequest -Method GET `
                                -Uri "https://graph.microsoft.com/v1.0/devices/$entry`?`$select=id,displayName"
                            $objId = $d['id']; $label = $d['displayName']
                        } else {
                            $safe = $entry -replace "'","''"
                            $dr = Invoke-MgGraphRequest -Method GET `
                                -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=displayName eq '$safe'`&`$select=id,displayName`&`$top=1"
                            if ($dr['value'] -and $dr['value'].Count -gt 0) {
                                $objId = $dr['value'][0]['id']; $label = $dr['value'][0]['displayName']
                            } else { throw "Device not found: $entry" }
                        }
                    }
                    'Groups' {
                        if ($entry -match $isGuidRx) {
                            $g = Invoke-MgGraphRequest -Method GET `
                                -Uri "https://graph.microsoft.com/v1.0/groups/$entry`?`$select=id,displayName"
                            $objId = $g['id']; $label = $g['displayName']
                        } else {
                            $safe = $entry -replace "'","''"
                            $gr = Invoke-MgGraphRequest -Method GET `
                                -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe'`&`$select=id,displayName`&`$top=1"
                            if ($gr['value'] -and $gr['value'].Count -gt 0) {
                                $objId = $gr['value'][0]['id']; $label = $gr['value'][0]['displayName']
                            } else { throw "Group not found: $entry" }
                        }
                    }
                }
                Write-VerboseLog "  $([char]0x2713) $label  ($objId)" -Level Success
                $resolvedIds.Add($objId)
                $resolvedLabels.Add($label)
            } catch {
                Write-VerboseLog "  $([char]0x2717) Could not resolve: $entry  -  $($_.Exception.Message)" -Level Warning
                $failedEntries.Add($entry)
            }
        }

        if ($resolvedIds.Count -eq 0) {
            throw 'None of the entries could be resolved in Entra. Check identifiers and try again.'
        }
        if ($resolvedIds.Count -lt 2) {
            throw "Only one entry was resolved ($($resolvedLabels[0])). At least two are needed to find common groups."
        }

        # ── Step 2: fetch group memberships for each object ──────────────────
        $membershipSets = [System.Collections.Generic.List[System.Collections.Generic.HashSet[string]]]::new()
        $groupMeta      = [System.Collections.Generic.Dictionary[string,hashtable]]::new()

        for ($i = 0; $i -lt $resolvedIds.Count; $i++) {
            $objId  = $resolvedIds[$i]
            $objLbl = $resolvedLabels[$i]
            Write-VerboseLog "Fetching memberships for: $objLbl" -Level Info

            $groupIds     = [System.Collections.Generic.HashSet[string]]::new()
            $baseEndpoint = if ($inputType -eq 'Users') {
                "https://graph.microsoft.com/v1.0/users/$objId/transitiveMemberOf"
            } else {
                "https://graph.microsoft.com/v1.0/directoryObjects/$objId/transitiveMemberOf"
            }
            $nextUri = $baseEndpoint + "?`$select=id,displayName,groupTypes,mailEnabled,securityEnabled,membershipRuleProcessingState`&`$top=999"
            do {
                try {
                    $resp = Invoke-MgGraphRequest -Method GET -Uri $nextUri
                    foreach ($obj in @($resp['value'])) {
                        $odataType = [string]$obj['@odata.type']
                        if ($odataType -and $odataType -ne '#microsoft.graph.group') { continue }
                        $gid = $obj['id']
                        $null = $groupIds.Add($gid)
                        if (-not $groupMeta.ContainsKey($gid)) {
                            $groupMeta[$gid] = @{
                                DisplayName                   = [string]$obj['displayName']
                                GroupTypes                    = @($obj['groupTypes'])
                                MailEnabled                   = [bool]$obj['mailEnabled']
                                SecurityEnabled               = [bool]$obj['securityEnabled']
                                MembershipRuleProcessingState = [string]$obj['membershipRuleProcessingState']
                            }
                        }
                    }
                    $nextUri = if ($resp.ContainsKey('@odata.nextLink')) { $resp['@odata.nextLink'] } else { $null }
                } catch {
                    Write-VerboseLog "  [!] Membership fetch error for $objLbl : $($_.Exception.Message)" -Level Warning
                    $nextUri = $null
                }
            } while ($nextUri)

            Write-VerboseLog "  $($groupIds.Count) group(s)" -Level Info
            $membershipSets.Add($groupIds)
        }

        # ── Step 3: intersect ────────────────────────────────────────────────
        $commonIds = [System.Collections.Generic.HashSet[string]]::new($membershipSets[0])
        for ($i = 1; $i -lt $membershipSets.Count; $i++) {
            $commonIds.IntersectWith($membershipSets[$i])
        }

        # ── Step 4: build flat result list ────────────────────────────────────
        $fcList = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($gid in $commonIds) {
            $meta = $groupMeta[$gid]
            if ($meta['GroupTypes'] -contains 'Unified') {
                $gType = 'M365'
            } elseif ($meta['MailEnabled'] -and $meta['SecurityEnabled']) {
                $gType = 'Mail-Enabled Security'
            } elseif ($meta['SecurityEnabled']) {
                $gType = 'Security'
            } else {
                $gType = 'Other'
            }
            $mType = if ($meta['MembershipRuleProcessingState'] -eq 'On') { 'Dynamic' } else { 'Assigned' }
            $fcList.Add([PSCustomObject]@{
                GroupName      = $meta['DisplayName']
                GroupID        = $gid
                GroupType      = $gType
                MembershipType = $mType
            })
        }
        $script:FCResult = [System.Collections.Generic.List[PSCustomObject]]($fcList | Sort-Object GroupType, GroupName)

        $cgLevel = if ($commonIds.Count -gt 0) { 'Success' } else { 'Warning' }
        Write-VerboseLog "Common groups total: $($commonIds.Count)" -Level $cgLevel
        if ($failedEntries.Count -gt 0) {
            Write-VerboseLog "  Skipped (unresolved): $($failedEntries -join ', ')" -Level Warning
        }

    } -Done {
        $script:Window.FindName('PnlFCProgress').Visibility = 'Collapsed'
        $resultList = $script:FCResult
        $count = if ($resultList) { $resultList.Count } else { 0 }

        if ($count -gt 0) {
            $script:Window.FindName('DgFCResults').ItemsSource    = $resultList
            $script:Window.FindName('TxtFCCount').Text            = "$count common group(s) found"
            $script:Window.FindName('PnlFCResults').Visibility    = 'Visible'
            $script:Window.FindName('TxtFCNoResults').Visibility  = 'Collapsed'
            $script:Window.FindName('BtnFCCopyAll').IsEnabled     = $true
            $script:Window.FindName('TxtFCFilter').IsEnabled  = $true
            $script:Window.FindName('BtnFCFilter').IsEnabled  = $true
            $script:Window.FindName('BtnFCFilterClear').IsEnabled = $true
            $script:Window.FindName('BtnFCExportXlsx').IsEnabled  = $true
            Write-VerboseLog "Find Common Groups complete. Found $count group(s)." -Level Success
        } else {
            $script:Window.FindName('PnlFCResults').Visibility    = 'Collapsed'
            $script:Window.FindName('TxtFCNoResults').Visibility  = 'Visible'
            Write-VerboseLog "Find Common Groups complete. No common groups found." -Level Warning
        }
    }
})

# ── DgFCResults: cell selection ───────────────────────────────────────────────
$script:Window.FindName('DgFCResults').Add_SelectedCellsChanged({
    $dg       = $script:Window.FindName('DgFCResults')
    $selCells = $dg.SelectedCells.Count
    $script:Window.FindName('BtnFCCopyValue').IsEnabled = ($selCells -ge 1)
    $script:Window.FindName('BtnFCCopyRow').IsEnabled   = ($selCells -gt 0)
})

# ── BtnFCCopyValue ────────────────────────────────────────────────────────────
# ── BtnFCFilter ──
$script:Window.FindName('BtnFCFilter').Add_Click({
    $dg      = $script:Window.FindName('DgFCResults')
    $keyword = $script:Window.FindName('TxtFCFilter').Text.Trim()
    if ($null -eq $script:FCAllData) {
        $script:FCAllData = @($dg.ItemsSource)
    }
    if ($keyword -eq '') {
        $dg.ItemsSource = $script:FCAllData
        Show-Notification "Filter cleared - $($script:FCAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    } else {
        $filtered = @($script:FCAllData | Where-Object {
            $row = $_
            $match = $false
            foreach ($prop in $row.PSObject.Properties) {
                if ([string]$prop.Value -like "*$keyword*") { $match = $true; break }
            }
            $match
        })

# ── BtnFCFilterClear ──
$script:Window.FindName('BtnFCFilterClear').Add_Click({
    $script:Window.FindName('TxtFCFilter').Text = ''
    $dg = $script:Window.FindName('DgFCResults')
    if ($null -ne $script:FCAllData) {
        $dg.ItemsSource = $script:FCAllData
        Show-Notification "Filter cleared - $($script:FCAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    }
})
        $dg.ItemsSource = $filtered
        Show-Notification "$($filtered.Count) of $($script:FCAllData.Count) result(s) shown after filter." -BgColor '#D4EDDA' -FgColor '#155724'
    }
    Write-VerboseLog "FC: Filter applied - keyword='$keyword'" -Level Info
})

$script:Window.FindName('BtnFCCopyValue').Add_Click({
    $dg = $script:Window.FindName('DgFCResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $vals = [System.Collections.Generic.List[string]]::new()
    foreach ($cellInfo in @($dg.SelectedCells)) {
        $col = $cellInfo.Column -as [System.Windows.Controls.DataGridBoundColumn]
        if ($col) { $vals.Add(([string]$cellInfo.Item.($col.Binding.Path.Path)).Trim()) }
    }
    [System.Windows.Clipboard]::SetText($vals -join "`r`n")
    Show-Notification "$($vals.Count) value(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "FC: Copied cell value to clipboard." -Level Success
})

# ── BtnFCCopyRow ──────────────────────────────────────────────────────────────
$script:Window.FindName('BtnFCCopyRow').Add_Click({
    $dg = $script:Window.FindName('DgFCResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $seen = [System.Collections.Generic.HashSet[object]]::new()
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($cell in @($dg.SelectedCells)) { if ($seen.Add($cell.Item)) { $rows.Add($cell.Item) } }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc    = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add(($colDefs | ForEach-Object { $_.H }) -join "`t")
    foreach ($row in $rows) {
        $vals = $colDefs | ForEach-Object { [string]$row.($_.P) }
        $lines.Add($vals -join "`t")
    }
    [System.Windows.Clipboard]::SetText($lines -join "`r`n")
    Show-Notification "$($rows.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "FC: Copied $($rows.Count) row(s) to clipboard." -Level Success
    # Highlight full rows
    $dg.UnselectAllCells()
    foreach ($r in $rows) { $null = $dg.SelectedItems.Add($r) }
})

# ── BtnFCCopyAll ──────────────────────────────────────────────────────────────
$script:Window.FindName('BtnFCCopyAll').Add_Click({
    $dg       = $script:Window.FindName('DgFCResults')
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) { return }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc    = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add(($colDefs | ForEach-Object { $_.H }) -join "`t")
    foreach ($row in $allItems) {
        $vals = $colDefs | ForEach-Object { [string]$row.($_.P) }
        $lines.Add($vals -join "`t")
    }
    [System.Windows.Clipboard]::SetText($lines -join "`r`n")
    Show-Notification "$($allItems.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "FC: Copied all $($allItems.Count) row(s) to clipboard." -Level Success
})

# ── BtnFCExportXlsx ───────────────────────────────────────────────────────────
$script:Window.FindName('BtnFCExportXlsx').Add_Click({
    $dg       = $script:Window.FindName('DgFCResults')
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) {
        Show-Notification "No results to export." -BgColor '#FFF3CD' -FgColor '#856404'
        return
    }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc    = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title            = "Export Find Common Groups Results"
    $dlg.Filter           = "Excel Workbook (*.xlsx)|*.xlsx|CSV File (*.csv)|*.csv|All Files (*.*)|*.*"
    $dlg.FileName         = "FindCommonGroups_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($dlg.ShowDialog() -ne $true) { return }
    $outPath = $dlg.FileName
    try {
        $folder = Split-Path $outPath -Parent
        if ($folder -and -not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        $exportRows = $allItems | ForEach-Object {
            $row = $_; $ht = [ordered]@{}
            foreach ($cd in $colDefs) { $ht[$cd['H']] = [string]$row.($cd['P']) }
            [PSCustomObject]$ht
        }
        if (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue) {
            Import-Module ImportExcel -ErrorAction SilentlyContinue
            $exportRows | Export-Excel -Path $outPath -WorksheetName "CommonGroups" -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium6
            Show-Notification "Exported $($allItems.Count) rows: $outPath" -BgColor '#D4EDDA' -FgColor '#155724'
        } else {
            $exportRows | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8
            Show-Notification "ImportExcel not found — saved as CSV: $(Split-Path $outPath -Leaf)" -BgColor '#FFF3CD' -FgColor '#856404'
        }
        Write-VerboseLog "FC: Exported $($allItems.Count) rows to: $outPath" -Level Success
    } catch {
        Show-Notification "Export failed: $($_.Exception.Message)" -BgColor '#F8D7DA' -FgColor '#721C24'
        Write-VerboseLog "FC: Export failed: $($_.Exception.Message)" -Level Error
    }
})



# ── FIND DISTINCT GROUPS ────────────────────────────────────────────────────

# Update the input label when the input type RadioButton changes
$script:Window.FindName('RbFDUsers').Add_Checked({
    $script:Window.FindName('TxtFDInputLabel').Text = 'USER UPNs  -  one per line'
    $script:Window.FindName('TxtFDInputList').Tag   = 'user1@domain.com'
})
$script:Window.FindName('RbFDDevices').Add_Checked({
    $script:Window.FindName('TxtFDInputLabel').Text = 'DEVICE NAMES or Object IDs  -  one per line'
    $script:Window.FindName('TxtFDInputList').Tag   = 'LAPTOP-001'
})
$script:Window.FindName('RbFDGroups').Add_Checked({
    $script:Window.FindName('TxtFDInputLabel').Text = 'GROUP NAMES or Object IDs  -  one per line'
    $script:Window.FindName('TxtFDInputList').Tag   = 'SG-Finance-All'
})

$script:Window.FindName('BtnOpFindDistinct').Add_Click({
    $script:Window.FindName('TxtFDInputList').Text  = ''
    $script:Window.FindName('RbFDUsers').IsChecked  = $true
    $script:Window.FindName('TxtFDInputLabel').Text = 'USER UPNs  -  one per line'
    $script:Window.FindName('TxtFDInputList').Tag   = 'user1@domain.com'
    Show-Panel 'PanelFindDistinct'
    Hide-Notification
    Write-VerboseLog 'Panel: Find Distinct Groups' -Level Info
})

$script:Window.FindName('BtnFDFind').Add_Click({
    Hide-Notification

    $inputTxt = $script:Window.FindName('TxtFDInputList').Text
    if ([string]::IsNullOrWhiteSpace($inputTxt)) {
        Show-Notification 'Please enter at least one identifier.' -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }

    $inputType = if ($script:Window.FindName('RbFDUsers').IsChecked)      { 'Users' }
                 elseif ($script:Window.FindName('RbFDDevices').IsChecked) { 'Devices' }
                 else { 'Groups' }

    $script:FDParams = @{ InputTxt = $inputTxt; InputType = $inputType }
    $btn = $script:Window.FindName('BtnFDFind')
    $script:Window.FindName('PnlFDResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtFDNoResults').Visibility = 'Collapsed'
    $script:Window.FindName('PnlFDProgress').Visibility  = 'Visible'
    $script:Window.FindName('TxtFDProgressMsg').Text     = "Resolving $inputType..."
    $script:Window.FindName('TxtFDProgressDetail').Text  = ''
    $script:Window.FindName('BtnFDCopyValue').IsEnabled  = $false
    $script:Window.FindName('BtnFDCopyRow').IsEnabled    = $false
    $script:Window.FindName('BtnFDCopyAll').IsEnabled    = $false
    $script:Window.FindName('BtnFDFilter').IsEnabled  = $false
    $script:Window.FindName('BtnFDFilterClear').IsEnabled = $false
    $script:Window.FindName('TxtFDFilter').Text       = ''
    $script:FDAllData = $null
    $script:Window.FindName('TxtFDFilter').IsEnabled  = $false
    $script:Window.FindName('BtnFDExportXlsx').IsEnabled = $false

    Invoke-OnBackground -DisableButton $btn -BusyMessage "--- Find Distinct Groups: resolving $inputType ---" -Work {
        $params    = $script:FDParams
        $inputType = $params['InputType']
        $entries   = $params['InputTxt'] -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }

        Write-VerboseLog "Input type : $inputType" -Level Info
        Write-VerboseLog "Entries    : $($entries.Count)" -Level Info

        if ($entries.Count -lt 2) {
            throw 'Please enter at least two identifiers to find distinct groups.'
        }

        # ── Step 1: resolve each entry to an Entra Object ID ─────────────────
        $resolvedIds    = [System.Collections.Generic.List[string]]::new()
        $resolvedLabels = [System.Collections.Generic.List[string]]::new()
        $failedEntries  = [System.Collections.Generic.List[string]]::new()
        $isGuidRx = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'

        foreach ($entry in $entries) {
            $objId = $null; $label = $entry
            try {
                switch ($inputType) {
                    'Users' {
                        $u = Invoke-MgGraphRequest -Method GET `
                            -Uri "https://graph.microsoft.com/v1.0/users/$([Uri]::EscapeDataString($entry))?`$select=id,displayName,userPrincipalName"
                        $objId = $u['id']; $label = $u['userPrincipalName']
                    }
                    'Devices' {
                        if ($entry -match $isGuidRx) {
                            $d = Invoke-MgGraphRequest -Method GET `
                                -Uri "https://graph.microsoft.com/v1.0/devices/$entry`?`$select=id,displayName"
                            $objId = $d['id']; $label = $d['displayName']
                        } else {
                            $safe = $entry -replace "'","''"
                            $dr = Invoke-MgGraphRequest -Method GET `
                                -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=displayName eq '$safe'`&`$select=id,displayName`&`$top=1"
                            if ($dr['value'] -and $dr['value'].Count -gt 0) {
                                $objId = $dr['value'][0]['id']; $label = $dr['value'][0]['displayName']
                            } else { throw "Device not found: $entry" }
                        }
                    }
                    'Groups' {
                        if ($entry -match $isGuidRx) {
                            $g = Invoke-MgGraphRequest -Method GET `
                                -Uri "https://graph.microsoft.com/v1.0/groups/$entry`?`$select=id,displayName"
                            $objId = $g['id']; $label = $g['displayName']
                        } else {
                            $safe = $entry -replace "'","''"
                            $gr = Invoke-MgGraphRequest -Method GET `
                                -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe'`&`$select=id,displayName`&`$top=1"
                            if ($gr['value'] -and $gr['value'].Count -gt 0) {
                                $objId = $gr['value'][0]['id']; $label = $gr['value'][0]['displayName']
                            } else { throw "Group not found: $entry" }
                        }
                    }
                }
                Write-VerboseLog "  $([char]0x2713) $label  ($objId)" -Level Success
                $resolvedIds.Add($objId)
                $resolvedLabels.Add($label)
            } catch {
                Write-VerboseLog "  $([char]0x2717) Could not resolve: $entry  -  $($_.Exception.Message)" -Level Warning
                $failedEntries.Add($entry)
            }
        }

        if ($resolvedIds.Count -eq 0) {
            throw 'None of the entries could be resolved in Entra. Check identifiers and try again.'
        }
        if ($resolvedIds.Count -lt 2) {
            throw ('Only one entry was resolved (' + $resolvedLabels[0] + '). At least two are needed to find distinct groups.')
        }

        # ── Step 2: fetch group memberships for each object ──────────────────
        $membershipSets = [System.Collections.Generic.List[System.Collections.Generic.HashSet[string]]]::new()
        $groupMeta      = [System.Collections.Generic.Dictionary[string,hashtable]]::new()

        for ($i = 0; $i -lt $resolvedIds.Count; $i++) {
            $objId  = $resolvedIds[$i]
            $objLbl = $resolvedLabels[$i]
            Write-VerboseLog "Fetching memberships for: $objLbl" -Level Info

            $groupIds     = [System.Collections.Generic.HashSet[string]]::new()
            $baseEndpoint = if ($inputType -eq 'Users') {
                "https://graph.microsoft.com/v1.0/users/$objId/transitiveMemberOf"
            } else {
                "https://graph.microsoft.com/v1.0/directoryObjects/$objId/transitiveMemberOf"
            }
            $nextUri = $baseEndpoint + "?`$select=id,displayName,groupTypes,mailEnabled,securityEnabled,membershipRuleProcessingState`&`$top=999"
            do {
                try {
                    $resp = Invoke-MgGraphRequest -Method GET -Uri $nextUri
                    foreach ($obj in @($resp['value'])) {
                        $odataType = [string]$obj['@odata.type']
                        if ($odataType -and $odataType -ne '#microsoft.graph.group') { continue }
                        $gid = $obj['id']
                        $null = $groupIds.Add($gid)
                        if (-not $groupMeta.ContainsKey($gid)) {
                            $groupMeta[$gid] = @{
                                DisplayName                   = [string]$obj['displayName']
                                GroupTypes                    = @($obj['groupTypes'])
                                MailEnabled                   = [bool]$obj['mailEnabled']
                                SecurityEnabled               = [bool]$obj['securityEnabled']
                                MembershipRuleProcessingState = [string]$obj['membershipRuleProcessingState']
                            }
                        }
                    }
                    $nextUri = if ($resp.ContainsKey('@odata.nextLink')) { $resp['@odata.nextLink'] } else { $null }
                } catch {
                    Write-VerboseLog "  [!] Membership fetch error for $objLbl : $($_.Exception.Message)" -Level Warning
                    $nextUri = $null
                }
            } while ($nextUri)

            Write-VerboseLog "  $($groupIds.Count) group(s)" -Level Info
            $membershipSets.Add($groupIds)
        }

        # ── Step 3: compute union minus intersection (distinct groups) ────────
        $unionIds = [System.Collections.Generic.HashSet[string]]::new($membershipSets[0])
        for ($i = 1; $i -lt $membershipSets.Count; $i++) {
            $unionIds.UnionWith($membershipSets[$i])
        }
        $intersectIds = [System.Collections.Generic.HashSet[string]]::new($membershipSets[0])
        for ($i = 1; $i -lt $membershipSets.Count; $i++) {
            $intersectIds.IntersectWith($membershipSets[$i])
        }
        # Distinct = groups in the union but NOT in the intersection
        # (i.e. groups that belong to SOME but NOT ALL input objects)
        $distinctIds = [System.Collections.Generic.HashSet[string]]::new($unionIds)
        $distinctIds.ExceptWith($intersectIds)


        # ── Step 3b: map distinct groups → input object labels ──────────────
        $groupOwnerMap = [System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[string]]]::new()
        foreach ($gid in $distinctIds) {
            $owners = [System.Collections.Generic.List[string]]::new()
            for ($i = 0; $i -lt $membershipSets.Count; $i++) {
                if ($membershipSets[$i].Contains($gid)) { $owners.Add($resolvedLabels[$i]) }
            }
            $groupOwnerMap[$gid] = $owners
        }
        # ── Step 4: build flat result list ────────────────────────────────────
        $fdList = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($gid in $distinctIds) {
            $meta = $groupMeta[$gid]
            if ($meta['GroupTypes'] -contains 'Unified') {
                $gType = 'M365'
            } elseif ($meta['MailEnabled'] -and $meta['SecurityEnabled']) {
                $gType = 'Mail-Enabled Security'
            } elseif ($meta['SecurityEnabled']) {
                $gType = 'Security'
            } else {
                $gType = 'Other'
            }
            $mType = if ($meta['MembershipRuleProcessingState'] -eq 'On') { 'Dynamic' } else { 'Assigned' }
            $fdList.Add([PSCustomObject]@{
                GroupName      = $meta['DisplayName']
                GroupID        = $gid
                GroupType      = $gType
                MembershipType = $mType
                MemberOf       = ($groupOwnerMap[$gid] -join ', ')
            })
        }
        $script:FDResult = [System.Collections.Generic.List[PSCustomObject]]($fdList | Sort-Object GroupType, GroupName)

        $fdLevel = if ($distinctIds.Count -gt 0) { 'Success' } else { 'Warning' }
        Write-VerboseLog "Distinct groups total: $($distinctIds.Count)  (union: $($unionIds.Count)  common: $($intersectIds.Count))" -Level $fdLevel
        if ($failedEntries.Count -gt 0) {
            Write-VerboseLog "  Skipped (unresolved): $($failedEntries -join ', ')" -Level Warning
        }

    } -Done {
        $script:Window.FindName('PnlFDProgress').Visibility = 'Collapsed'
        $resultList = $script:FDResult
        $count = if ($resultList) { $resultList.Count } else { 0 }

        if ($count -gt 0) {
            $script:Window.FindName('DgFDResults').ItemsSource    = $resultList
            $script:Window.FindName('TxtFDCount').Text            = "$count distinct group(s) found"
            $script:Window.FindName('PnlFDResults').Visibility    = 'Visible'
            $script:Window.FindName('TxtFDNoResults').Visibility  = 'Collapsed'
            $script:Window.FindName('BtnFDCopyAll').IsEnabled     = $true
            $script:Window.FindName('TxtFDFilter').IsEnabled  = $true
            $script:Window.FindName('BtnFDFilter').IsEnabled  = $true
            $script:Window.FindName('BtnFDFilterClear').IsEnabled = $true
            $script:Window.FindName('BtnFDExportXlsx').IsEnabled  = $true
            Write-VerboseLog "Find Distinct Groups complete. Found $count group(s)." -Level Success
        } else {
            $script:Window.FindName('PnlFDResults').Visibility    = 'Collapsed'
            $script:Window.FindName('TxtFDNoResults').Visibility  = 'Visible'
            Write-VerboseLog "Find Distinct Groups complete. No distinct groups found." -Level Warning
        }
    }
})

# ── DgFDResults: cell selection ───────────────────────────────────────────────
$script:Window.FindName('DgFDResults').Add_SelectedCellsChanged({
    $dg       = $script:Window.FindName('DgFDResults')
    $selCells = $dg.SelectedCells.Count
    $script:Window.FindName('BtnFDCopyValue').IsEnabled = ($selCells -ge 1)
    $script:Window.FindName('BtnFDCopyRow').IsEnabled   = ($selCells -gt 0)
})

# ── BtnFDCopyValue ────────────────────────────────────────────────────────────
# ── BtnFDFilter ──
$script:Window.FindName('BtnFDFilter').Add_Click({
    $dg      = $script:Window.FindName('DgFDResults')
    $keyword = $script:Window.FindName('TxtFDFilter').Text.Trim()
    if ($null -eq $script:FDAllData) {
        $script:FDAllData = @($dg.ItemsSource)
    }
    if ($keyword -eq '') {
        $dg.ItemsSource = $script:FDAllData
        Show-Notification "Filter cleared - $($script:FDAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    } else {
        $filtered = @($script:FDAllData | Where-Object {
            $row = $_
            $match = $false
            foreach ($prop in $row.PSObject.Properties) {
                if ([string]$prop.Value -like "*$keyword*") { $match = $true; break }
            }
            $match
        })

# ── BtnFDFilterClear ──
$script:Window.FindName('BtnFDFilterClear').Add_Click({
    $script:Window.FindName('TxtFDFilter').Text = ''
    $dg = $script:Window.FindName('DgFDResults')
    if ($null -ne $script:FDAllData) {
        $dg.ItemsSource = $script:FDAllData
        Show-Notification "Filter cleared - $($script:FDAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    }
})
        $dg.ItemsSource = $filtered
        Show-Notification "$($filtered.Count) of $($script:FDAllData.Count) result(s) shown after filter." -BgColor '#D4EDDA' -FgColor '#155724'
    }
    Write-VerboseLog "FD: Filter applied - keyword='$keyword'" -Level Info
})

$script:Window.FindName('BtnFDCopyValue').Add_Click({
    $dg = $script:Window.FindName('DgFDResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $vals = [System.Collections.Generic.List[string]]::new()
    foreach ($cellInfo in @($dg.SelectedCells)) {
        $col = $cellInfo.Column -as [System.Windows.Controls.DataGridBoundColumn]
        if ($col) { $vals.Add(([string]$cellInfo.Item.($col.Binding.Path.Path)).Trim()) }
    }
    [System.Windows.Clipboard]::SetText($vals -join "`r`n")
    Show-Notification "$($vals.Count) value(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "FD: Copied cell value to clipboard." -Level Success
})

# ── BtnFDCopyRow ──────────────────────────────────────────────────────────────
$script:Window.FindName('BtnFDCopyRow').Add_Click({
    $dg = $script:Window.FindName('DgFDResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $seen = [System.Collections.Generic.HashSet[object]]::new()
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($cell in @($dg.SelectedCells)) { if ($seen.Add($cell.Item)) { $rows.Add($cell.Item) } }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc    = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add(($colDefs | ForEach-Object { $_.H }) -join "`t")
    foreach ($row in $rows) {
        $vals = $colDefs | ForEach-Object { [string]$row.($_.P) }
        $lines.Add($vals -join "`t")
    }
    [System.Windows.Clipboard]::SetText($lines -join "`r`n")
    Show-Notification "$($rows.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "FD: Copied $($rows.Count) row(s) to clipboard." -Level Success
    # Highlight full rows
    $dg.UnselectAllCells()
    foreach ($r in $rows) { $null = $dg.SelectedItems.Add($r) }
})

# ── BtnFDCopyAll ──────────────────────────────────────────────────────────────
$script:Window.FindName('BtnFDCopyAll').Add_Click({
    $dg       = $script:Window.FindName('DgFDResults')
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) { return }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc    = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add(($colDefs | ForEach-Object { $_.H }) -join "`t")
    foreach ($row in $allItems) {
        $vals = $colDefs | ForEach-Object { [string]$row.($_.P) }
        $lines.Add($vals -join "`t")
    }
    [System.Windows.Clipboard]::SetText($lines -join "`r`n")
    Show-Notification "$($allItems.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "FD: Copied all $($allItems.Count) row(s) to clipboard." -Level Success
})

# ── BtnFDExportXlsx ───────────────────────────────────────────────────────────
$script:Window.FindName('BtnFDExportXlsx').Add_Click({
    $dg       = $script:Window.FindName('DgFDResults')
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) {
        Show-Notification "No results to export." -BgColor '#FFF3CD' -FgColor '#856404'
        return
    }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc    = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title            = "Export Find Distinct Groups Results"
    $dlg.Filter           = "Excel Workbook (*.xlsx)|*.xlsx|CSV File (*.csv)|*.csv|All Files (*.*)|*.*"
    $dlg.FileName         = "FindDistinctGroups_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($dlg.ShowDialog() -ne $true) { return }
    $outPath = $dlg.FileName
    try {
        $folder = Split-Path $outPath -Parent
        if ($folder -and -not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        $exportRows = $allItems | ForEach-Object {
            $row = $_; $ht = [ordered]@{}
            foreach ($cd in $colDefs) { $ht[$cd['H']] = [string]$row.($cd['P']) }
            [PSCustomObject]$ht
        }
        if (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue) {
            Import-Module ImportExcel -ErrorAction SilentlyContinue
            $exportRows | Export-Excel -Path $outPath -WorksheetName "DistinctGroups" -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium6
            Show-Notification "Exported $($allItems.Count) rows: $outPath" -BgColor '#D4EDDA' -FgColor '#155724'
        } else {
            $exportRows | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8
            Show-Notification "ImportExcel not found — saved as CSV: $(Split-Path $outPath -Leaf)" -BgColor '#FFF3CD' -FgColor '#856404'
        }
        Write-VerboseLog "FD: Exported $($allItems.Count) rows to: $outPath" -Level Success
    } catch {
        Show-Notification "Export failed: $($_.Exception.Message)" -BgColor '#F8D7DA' -FgColor '#721C24'
        Write-VerboseLog "FD: Export failed: $($_.Exception.Message)" -Level Error
    }
})


# ── DEVICE MANAGEMENT BLADE ─────────────────────────────────────────

# Blade toggle
$script:BladeDevExpanded = $false
$script:Window.FindName('BtnBladeDevMgmt').Add_Click({
    $script:BladeDevExpanded = -not $script:BladeDevExpanded
    $content = $script:Window.FindName('BladeDevMgmtContent')
    $divider = $script:Window.FindName('BladeDevDivider')
    $arrow   = $script:Window.FindName('TxtBladeDevArrow')
    if ($script:BladeDevExpanded) {
        $content.Visibility = 'Visible'
        $divider.Visibility = 'Visible'
        $arrow.Text = [char]0x2212
    } else {
        $content.Visibility = 'Collapsed'
        $divider.Visibility = 'Collapsed'
        $arrow.Text = '+'
    }
})

# ÄÄ Collapse Device Mgmt blade on startup ÄÄ
$script:Window.FindName('BladeDevMgmtContent').Visibility = 'Collapsed'
$script:Window.FindName('BladeDevDivider').Visibility = 'Collapsed'
$script:Window.FindName('TxtBladeDevArrow').Text = '+'

# Input type ComboBox → update label
$script:Window.FindName('CmbGDIInputType').Add_SelectionChanged({
    $sel = $script:Window.FindName('CmbGDIInputType').SelectedItem.Content
    $lbl = switch ($sel) {
        'User UPNs (resolve owned devices)' { 'USER UPNs  -  one per line' }
        'Device names'                       { 'DEVICE NAMES  -  one per line' }
        'Device serial numbers'              { 'SERIAL NUMBERS  -  one per line' }
        'Entra Device Object IDs'            { 'ENTRA DEVICE OBJECT IDs  -  one per line (GUIDs)' }
        'Intune Device IDs'                  { 'INTUNE DEVICE IDs  -  one per line (GUIDs)' }
        'Groups (Users and Devices)'         { 'GROUP NAMES or Object IDs  -  one per line' }
        default                              { 'INPUT  -  one per line' }
    }
    $script:Window.FindName('TxtGDIInputLabel').Text = $lbl
})

# Panel open
$script:Window.FindName('BtnOpGetDeviceInfo').Add_Click({
    $script:Window.FindName('TxtGDIInputList').Text          = ''
    $script:Window.FindName('CmbGDIInputType').SelectedIndex = 0
    $script:Window.FindName('CmbGDIOwnership').SelectedIndex = 0
    $script:Window.FindName('TxtGDIInputLabel').Text         = 'USER UPNs  -  one per line'
    foreach ($chk in @('ChkGDIWindows','ChkGDIAndroid','ChkGDIiOS','ChkGDIMacOS')) {
        $script:Window.FindName($chk).IsChecked = $false
    }
    $script:Window.FindName('PnlGDIProgress').Visibility  = 'Collapsed'
    $script:Window.FindName('PnlGDIResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtGDINoResults').Visibility = 'Collapsed'
    $script:Window.FindName('BtnGDICopyValue').IsEnabled  = $false
    $script:Window.FindName('BtnGDICopyRow').IsEnabled    = $false
    $script:Window.FindName('BtnGDICopyAll').IsEnabled    = $false
    $script:Window.FindName('TxtGDIFilter').IsEnabled  = $false
    $script:Window.FindName('BtnGDIFilter').IsEnabled  = $false
    $script:Window.FindName('BtnGDIFilterClear').IsEnabled = $false
    $script:Window.FindName('TxtGDIFilter').Text       = ''
    $script:GDIAllData = $null
    $script:Window.FindName('BtnGDIExportXlsx').IsEnabled = $false
    $dg = $script:Window.FindName('DgGDIResults')
    if ($dg) { $dg.ItemsSource = $null }
    Show-Panel 'PanelGetDeviceInfo'
    Hide-Notification
    Write-VerboseLog 'Panel: Get Device Info' -Level Info
})


# Panel open: Get App Info
$script:Window.FindName('BtnOpGetDiscoveredApps').Add_Click({
    $script:Window.FindName('TxtDAInputList').Text          = ''
    $script:Window.FindName('TxtDAKeyword').Text            = ''
    $script:Window.FindName('CmbDAInputType').SelectedIndex = 0
    $script:Window.FindName('CmbDAOwnership').SelectedIndex = 0
    $script:Window.FindName('TxtDAInputLabel').Text         = 'USER UPNs  -  one per line'
    foreach ($chk in @('ChkDAWindows','ChkDAAndroid','ChkDAiOS','ChkDAMacOS')) {
        $script:Window.FindName($chk).IsChecked = $false
    }
    $script:Window.FindName('ChkDADiscoveredApps').IsChecked = $true
    $script:Window.FindName('ChkDAManagedApps').IsChecked    = $true
    $script:Window.FindName('BtnDARun').IsEnabled             = $true
    $script:Window.FindName('BtnDARunAll').IsEnabled          = $true
    $script:Window.FindName('BtnDAGetManagedAssignments').IsEnabled = $true
    $script:Window.FindName('PnlDAProgress').Visibility  = 'Collapsed'
    $script:Window.FindName('PnlDAResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtDANoResults').Visibility = 'Collapsed'
    $script:Window.FindName('BtnDACopyValue').IsEnabled  = $false
    $script:Window.FindName('BtnDACopyRow').IsEnabled    = $false
    $script:Window.FindName('BtnDACopyAll').IsEnabled    = $false
    $script:Window.FindName('TxtDAFilter').IsEnabled  = $false
    $script:Window.FindName('BtnDAFilter').IsEnabled  = $false
    $script:Window.FindName('BtnDAFilterClear').IsEnabled = $false
    $script:Window.FindName('TxtDAFilter').Text       = ''
    $script:DAAllData = $null
    $script:Window.FindName('BtnDAExportXlsx').IsEnabled = $false
    $dg = $script:Window.FindName('DgDAResults')
    if ($dg) { $dg.ItemsSource = $null; $dg.Columns.Clear() }
    Show-Panel 'PanelGetDiscoveredApps'
    Hide-Notification
    Write-VerboseLog 'Panel: Get App Assignments' -Level Info
})

$script:Window.FindName('CmbDAInputType').Add_SelectionChanged({
    $sel = $script:Window.FindName('CmbDAInputType').SelectedItem.Content
    $label = switch ($sel) {
        'User UPNs (resolve owned devices)' { 'USER UPNs  -  one per line' }
        'Device names'                       { 'DEVICE NAMES  -  one per line' }
        'Device IDs (Entra)'                 { 'ENTRA DEVICE IDs  -  one per line' }
        'Serial numbers'                     { 'SERIAL NUMBERS  -  one per line' }
        'Groups (names or IDs)'              { 'GROUP names or IDs  -  one per line' }
        default                              { 'IDENTIFIERS  -  one per line' }
    }
    $script:Window.FindName('TxtDAInputLabel').Text = $label
})
# ── GET APP INFO  -  Run ────────────────────────────────────────
$script:Window.FindName('BtnDARun').Add_Click({
    Hide-Notification

    $inputTxt   = $script:Window.FindName('TxtDAInputList').Text
    $inputTypeSel = $script:Window.FindName('CmbDAInputType').SelectedItem.Content

    if ([string]::IsNullOrWhiteSpace($inputTxt)) {
        Show-Notification 'Please enter at least one identifier.' -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }

    # Collect filters on UI thread
    $platFilter = [System.Collections.Generic.List[string]]::new()
    if ($script:Window.FindName('ChkDAWindows').IsChecked) { $platFilter.Add('Windows') }
    if ($script:Window.FindName('ChkDAAndroid').IsChecked) { $platFilter.Add('Android') }
    if ($script:Window.FindName('ChkDAiOS').IsChecked)     { $platFilter.Add('iOS') }
    if ($script:Window.FindName('ChkDAMacOS').IsChecked)   { $platFilter.Add('macOS') }

    $ownershipSel = $script:Window.FindName('CmbDAOwnership').SelectedItem.Content
    $ownershipFilter = switch ($ownershipSel) {
        'Company only'  { 'Company' }
        'Personal only' { 'Personal' }
        default         { 'All' }
    }

    $keywordTxt = $script:Window.FindName('TxtDAKeyword').Text

    $script:DAParams = @{
        InputTxt        = $inputTxt
        InputType       = $inputTypeSel
        PlatformFilter  = $platFilter
        OwnershipFilter = $ownershipFilter
        Keyword         = $keywordTxt
        ShowDiscovered  = $script:Window.FindName('ChkDADiscoveredApps').IsChecked -eq $true
        ShowManaged     = $script:Window.FindName('ChkDAManagedApps').IsChecked -eq $true
    }

    # Reset UI
    $script:Window.FindName('PnlDAResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtDANoResults').Visibility = 'Collapsed'
    $script:Window.FindName('BtnDACopyValue').IsEnabled  = $false
    $script:Window.FindName('BtnDACopyRow').IsEnabled    = $false
    $script:Window.FindName('BtnDACopyAll').IsEnabled    = $false
    $script:Window.FindName('TxtDAFilter').Text       = ''
    $script:DAAllData = $null
    $script:Window.FindName('TxtDAFilter').IsEnabled  = $false
    $script:Window.FindName('BtnDAFilter').IsEnabled  = $false
    $script:Window.FindName('BtnDAFilterClear').IsEnabled = $false
    $script:Window.FindName('BtnDAExportXlsx').IsEnabled = $false
    $dg = $script:Window.FindName('DgDAResults')
    if ($dg) { $dg.ItemsSource = $null }
    $script:Window.FindName('PnlDAProgress').Visibility      = 'Visible'
    $script:Window.FindName('TxtDAProgressMsg').Text         = 'Resolving devices...'
    $script:Window.FindName('TxtDAProgressDetail').Text      = ''

    $script:Window.FindName('BtnDARun').IsEnabled = $false
    $script:Window.FindName('BtnDARunAll').IsEnabled = $false
    $script:Window.FindName('BtnDAGetManagedAssignments').IsEnabled = $false
    $script:Window.FindName('BtnStop').Visibility = 'Visible'
    $script:Window.FindName('BtnStop').IsEnabled  = $true
    $script:Window.FindName('BtnStop').Content = (New-StopBtnContent 'Stop')
    $script:BgStopped = $false
    Write-VerboseLog '--- Get App Assignments: starting ---' -Level Action

    $script:DaQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $script:DaStop  = [hashtable]::Synchronized(@{ Stop = $false })

    $script:DaRs = [runspacefactory]::CreateRunspace()
    $script:DaRs.ApartmentState = 'STA'
    $script:DaRs.ThreadOptions  = 'ReuseThread'
    $script:DaRs.Open()
    $script:DaRs.SessionStateProxy.SetVariable('DaQueue',  $script:DaQueue)
    $script:DaRs.SessionStateProxy.SetVariable('DaStop',   $script:DaStop)
    $script:DaRs.SessionStateProxy.SetVariable('DaParams', $script:DAParams)
    $script:DaRs.SessionStateProxy.SetVariable('LogFile',  $script:LogFile)

    $script:DaPs = [powershell]::Create()
    $script:DaPs.Runspace = $script:DaRs

    $null = $script:DaPs.AddScript({
        function DLog {
            param([string]$Msg, [string]$Level = 'Info')
            $prefix = switch ($Level) {
                "Success" { "$([char]0x2713) " } "Warning" { "[!] " } "Error" { "$([char]0x2717) " }
                'Action'  { '> ' } default   { '  ' }
            }
            $line = "[$(Get-Date -Format 'HH:mm:ss')] $prefix$Msg"
            $DaQueue.Enqueue(@{ Type='log'; Line=$line; Level=$Level })
            if ($LogFile) {
                try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
            }
        }

        function Get-GraphAll([string]$Uri) {
            $items = [System.Collections.Generic.List[object]]::new()
            $next  = $Uri
            do {
                $r = Invoke-MgGraphRequest -Method GET -Uri $next
                if ($r['value']) { foreach ($i in @($r['value'])) { $items.Add($i) } }
                $next = if ($r.ContainsKey('@odata.nextLink')) { $r['@odata.nextLink'] } else { $null }
            } while ($next)
            return ,$items
        }

        $isGuidRx = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'

        function Normalize-OS([string]$os) {
            switch -Wildcard ($os) {
                'Windows*' { 'Windows' } 'Android*' { 'Android' }
                'iPhone'   { 'iOS'     } 'iPadOS'   { 'iOS'     }
                'iOS'      { 'iOS'     } 'MacMDM'   { 'macOS'   }
                'Mac OS X' { 'macOS'   } 'macOS'    { 'macOS'   }
                default    { $os }
            }
        }

        try {
            $p = $DaParams
            $entries = @($p['InputTxt'] -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
            $inputType = $p['InputType']
            $platFilter = $p['PlatformFilter']
            $ownerFilter = $p['OwnershipFilter']
            $keyword = $p['Keyword']
            $showDiscovered = $p['ShowDiscovered']
            $showManaged    = $p['ShowManaged']
            $intuneCache = @{}

            DLog "Processing $($entries.Count) input(s) - Type: $inputType" 'Action'

            # ── STEP 1: Resolve inputs to Intune device records ──
            $deviceRecords = [System.Collections.Generic.List[hashtable]]::new()
            $seenIntuneIds = [System.Collections.Generic.HashSet[string]]::new()
            $idx = 0

            foreach ($entry in $entries) {
                if ($DaStop['Stop']) { DLog 'Stop requested - halting.' 'Warning'; break }
                $idx++
                $DaQueue.Enqueue(@{ Type='progress'; Msg="Step 1/3: Resolving input $idx of $($entries.Count)..."; Detail=$entry })
                DLog "Resolving: $entry" 'Info'

                switch ($inputType) {
                    'User UPNs (resolve owned devices)' {
                        try {
                            $enc = [Uri]::EscapeDataString($entry)
                            $uDevs = Get-GraphAll "https://graph.microsoft.com/v1.0/users/$enc/ownedDevices?`$select=deviceId,displayName`&`$top=999"
                            DLog "  User owns $($uDevs.Count) device(s)" 'Info'
                            foreach ($ud in $uDevs) {
                                $azDevId = [string]$ud['deviceId']
                                if ([string]::IsNullOrWhiteSpace($azDevId)) { continue }
                                try {
                                    $encDev = [Uri]::EscapeDataString($azDevId)
                                    $iResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=azureADDeviceId eq '$encDev'`&`$select=id,deviceName,operatingSystem,osVersion,managedDeviceOwnerType,azureADDeviceId,userPrincipalName`&`$top=1"
                                    if ($iResp['value'] -and $iResp['value'].Count -gt 0) {
                                        $iDev = $iResp['value'][0]
                                        $iId = [string]$iDev['id']
                                        if ($seenIntuneIds.Add($iId)) {
                                            $deviceRecords.Add(@{
                                                IntuneId      = $iId
                                                DeviceName    = [string]$iDev['deviceName']
                                                EntraDeviceId = $azDevId
                                                OS            = [string]$iDev['operatingSystem']
                                                OSVersion     = [string]$iDev['osVersion']
                                                Ownership     = [string]$iDev['managedDeviceOwnerType']
                                                UPN           = [string]$iDev['userPrincipalName']
                                            })
                                        }
                                    }
                                } catch {}
                            }
                        } catch { DLog "  User resolve failed: $($_.Exception.Message)" 'Warning' }
                    }
                    'Device names' {
                        try {
                            $safe = $entry -replace "'","''"
                            $dResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=displayName eq '$safe'`&`$select=deviceId,displayName`&`$top=5"
                            foreach ($d in @($dResp['value'])) {
                                $azDevId = [string]$d['deviceId']
                                if ([string]::IsNullOrWhiteSpace($azDevId)) { continue }
                                $encDev = [Uri]::EscapeDataString($azDevId)
                                try {
                                    $iResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=azureADDeviceId eq '$encDev'`&`$select=id,deviceName,operatingSystem,osVersion,managedDeviceOwnerType,azureADDeviceId,userPrincipalName`&`$top=1"
                                    if ($iResp['value'] -and $iResp['value'].Count -gt 0) {
                                        $iDev = $iResp['value'][0]
                                        $iId = [string]$iDev['id']
                                        if ($seenIntuneIds.Add($iId)) {
                                            $deviceRecords.Add(@{
                                                IntuneId      = $iId
                                                DeviceName    = [string]$iDev['deviceName']
                                                EntraDeviceId = $azDevId
                                                OS            = [string]$iDev['operatingSystem']
                                                OSVersion     = [string]$iDev['osVersion']
                                                Ownership     = [string]$iDev['managedDeviceOwnerType']
                                                UPN           = [string]$iDev['userPrincipalName']
                                            })
                                        }
                                    }
                                } catch {}
                            }
                        } catch { DLog "  Device name resolve failed: $($_.Exception.Message)" 'Warning' }
                    }
                    'Device IDs (Entra)' {
                        try {
                            $encDev = [Uri]::EscapeDataString($entry)
                            $iResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=azureADDeviceId eq '$encDev'`&`$select=id,deviceName,operatingSystem,osVersion,managedDeviceOwnerType,azureADDeviceId,userPrincipalName`&`$top=1"
                            if ($iResp['value'] -and $iResp['value'].Count -gt 0) {
                                $iDev = $iResp['value'][0]
                                $iId = [string]$iDev['id']
                                if ($seenIntuneIds.Add($iId)) {
                                    $deviceRecords.Add(@{
                                        IntuneId      = $iId
                                        DeviceName    = [string]$iDev['deviceName']
                                        EntraDeviceId = [string]$iDev['azureADDeviceId']
                                        OS            = [string]$iDev['operatingSystem']
                                        OSVersion     = [string]$iDev['osVersion']
                                        Ownership     = [string]$iDev['managedDeviceOwnerType']
                                        UPN           = [string]$iDev['userPrincipalName']
                                    })
                                }
                            } else { DLog "  No Intune record for Entra ID: $entry" 'Warning' }
                        } catch { DLog "  Entra ID resolve failed: $($_.Exception.Message)" 'Warning' }
                    }
                    'Serial numbers' {
                        try {
                            $encSer = [Uri]::EscapeDataString($entry)
                            $iResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=serialNumber eq '$encSer'`&`$select=id,deviceName,operatingSystem,osVersion,managedDeviceOwnerType,azureADDeviceId,userPrincipalName`&`$top=1"
                            if ($iResp['value'] -and $iResp['value'].Count -gt 0) {
                                $iDev = $iResp['value'][0]
                                $iId = [string]$iDev['id']
                                if ($seenIntuneIds.Add($iId)) {
                                    $deviceRecords.Add(@{
                                        IntuneId      = $iId
                                        DeviceName    = [string]$iDev['deviceName']
                                        EntraDeviceId = [string]$iDev['azureADDeviceId']
                                        OS            = [string]$iDev['operatingSystem']
                                        OSVersion     = [string]$iDev['osVersion']
                                        Ownership     = [string]$iDev['managedDeviceOwnerType']
                                        UPN           = [string]$iDev['userPrincipalName']
                                    })
                                }
                            } else { DLog "  No Intune record for serial: $entry" 'Warning' }
                        } catch { DLog "  Serial resolve failed: $($_.Exception.Message)" 'Warning' }
                    }
                    'Groups (names or IDs)' {
                        try {
                            $groupId = $null
                            if ($entry -match $isGuidRx) {
                                $groupId = $entry
                            } else {
                                $safe = $entry -replace "'","''"
                                $gr = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe'`&`$select=id`&`$top=1"
                                if ($gr['value'] -and $gr['value'].Count -gt 0) {
                                    $groupId = [string]$gr['value'][0]['id']
                                }
                            }
                            if (-not $groupId) { DLog "  Group not found: $entry" 'Warning'; continue }

                            DLog "  Scanning group members (transitive)..." 'Info'
                            $members = Get-GraphAll "https://graph.microsoft.com/v1.0/groups/$groupId/transitiveMembers?`$select=id&`$top=999"
                            DLog "  $($members.Count) member(s) found" 'Info'

                            foreach ($mem in $members) {
                                if ($DaStop['Stop']) { break }
                                $odataType = [string]$mem['@odata.type']
                                if ($odataType -eq '#microsoft.graph.device') {
                                    # Direct device member - resolve via Entra deviceId
                                    $devObjId = [string]$mem['id']
                                    try {
                                        $entraDevRaw = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/devices/$devObjId`?`$select=deviceId"
                                        $azDevId = [string]$entraDevRaw['deviceId']
                                        if ([string]::IsNullOrWhiteSpace($azDevId)) { continue }
                                        $encDev = [Uri]::EscapeDataString($azDevId)
                                        $iResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=azureADDeviceId eq '$encDev'&`$select=id,deviceName,operatingSystem,osVersion,managedDeviceOwnerType,azureADDeviceId,userPrincipalName&`$top=1"
                                        if ($iResp['value'] -and $iResp['value'].Count -gt 0) {
                                            $iDev = $iResp['value'][0]
                                            $iId = [string]$iDev['id']
                                            if ($seenIntuneIds.Add($iId)) {
                                                $deviceRecords.Add(@{
                                                    IntuneId      = $iId
                                                    DeviceName    = [string]$iDev['deviceName']
                                                    EntraDeviceId = $azDevId
                                                    OS            = [string]$iDev['operatingSystem']
                                                    OSVersion     = [string]$iDev['osVersion']
                                                    Ownership     = [string]$iDev['managedDeviceOwnerType']
                                                    UPN           = [string]$iDev['userPrincipalName']
                                                })
                                            }
                                        }
                                    } catch {}
                                } elseif ($odataType -eq '#microsoft.graph.user') {
                                    # User member - resolve their owned devices
                                    $uid = [string]$mem['id']
                                    try {
                                        $uDevs = Get-GraphAll "https://graph.microsoft.com/v1.0/users/$uid/ownedDevices?`$select=deviceId,displayName&`$top=999"
                                        DLog "    User $uid owns $($uDevs.Count) device(s)" 'Info'
                                        foreach ($ud in $uDevs) {
                                            if ($DaStop['Stop']) { break }
                                            $azDevId = [string]$ud['deviceId']
                                            if ([string]::IsNullOrWhiteSpace($azDevId)) { continue }
                                            try {
                                                $encDev = [Uri]::EscapeDataString($azDevId)
                                                $iResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=azureADDeviceId eq '$encDev'&`$select=id,deviceName,operatingSystem,osVersion,managedDeviceOwnerType,azureADDeviceId,userPrincipalName&`$top=1"
                                                if ($iResp['value'] -and $iResp['value'].Count -gt 0) {
                                                    $iDev = $iResp['value'][0]
                                                    $iId = [string]$iDev['id']
                                                    if ($seenIntuneIds.Add($iId)) {
                                                        $deviceRecords.Add(@{
                                                            IntuneId      = $iId
                                                            DeviceName    = [string]$iDev['deviceName']
                                                            EntraDeviceId = $azDevId
                                                            OS            = [string]$iDev['operatingSystem']
                                                            OSVersion     = [string]$iDev['osVersion']
                                                            Ownership     = [string]$iDev['managedDeviceOwnerType']
                                                            UPN           = [string]$iDev['userPrincipalName']
                                                        })
                                                    }
                                                }
                                            } catch {}
                                        }
                                    } catch { DLog "    User device resolve failed for $uid : $($_.Exception.Message)" 'Warning' }
                                }
                            }
                        } catch { DLog "  Group resolve failed: $($_.Exception.Message)" 'Warning' }
                    }
                }
            }

            DLog "Step 1 complete: $($deviceRecords.Count) unique Intune device(s) resolved" 'Info'

            if ($deviceRecords.Count -eq 0) {
                $DaQueue.Enqueue(@{ Type='done'; Rows=@(); Error='No devices resolved from the provided input.' })
                return
            }

            # ── STEP 2: Apply platform and ownership filters ──
            $DaQueue.Enqueue(@{ Type='progress'; Msg="Step 2/3: Applying filters..."; Detail='' })
            $filtered = [System.Collections.Generic.List[hashtable]]::new()
            foreach ($rec in $deviceRecords) {
                $osNorm = Normalize-OS $rec['OS']
                if ($platFilter.Count -gt 0 -and $osNorm -notin $platFilter) { continue }
                if ($ownerFilter -ne 'All' -and $rec['Ownership'] -ne $ownerFilter) { continue }
                $filtered.Add($rec)
            }
            DLog "Step 2 complete: $($filtered.Count) device(s) after filtering (from $($deviceRecords.Count))" 'Info'

            if ($filtered.Count -eq 0) {
                $DaQueue.Enqueue(@{ Type='done'; Rows=@(); Error='No devices matched the platform/ownership filters.' })
                return
            }





            # ── STEP 3: Fetch apps per device ──
            $allRows = [System.Collections.Generic.List[PSCustomObject]]::new()

            # Intent and state mappings for managed apps
            $intentMap = @{
                'available'                       = 'Available'
                'notAvailable'                    = 'Not Available'
                'requiredInstall'                 = 'Required Install'
                'requiredUninstall'               = 'Required Uninstall'
                'requiredAndAvailableInstall'     = 'Required and Available Install'
                'availableInstallWithoutEnrollment' = 'Available (No Enrollment)'
                'exclude'                         = 'Exclude'
            }
            $stateMap = @{
                'installed'      = 'Installed'
                'failed'         = 'Failed'
                'notInstalled'   = 'Not Installed'
                'uninstallFailed'= 'Uninstall Failed'
                'pendingInstall' = 'Pending Install'
                'unknown'        = 'Unknown'
                'notApplicable'  = 'Not Applicable'
            }

            # User ID cache (UPN -> GUID) to avoid redundant lookups
            $userIdCache = @{}

            # App description cache (applicationId -> description) to avoid redundant lookups
            $appDescCache = @{}

            $devIdx = 0
            foreach ($rec in $filtered) {
                if ($DaStop['Stop']) { DLog 'Stop requested - halting.' 'Warning'; break }
                $devIdx++
                $DaQueue.Enqueue(@{ Type='progress'; Msg="Step 3/3: Fetching apps for device $devIdx of $($filtered.Count)..."; Detail=$rec['DeviceName'] })
                DLog "[$devIdx] Fetching apps for: $($rec['DeviceName'])" 'Info'

                $collectTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

                # ── 3A: Discovered Apps (detectedApps API - per device) ──
                if ($showDiscovered) {
                    try {
                        $appsUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($rec['IntuneId'])/detectedApps?`$select=displayName,version,platform,publisher"
                        $apps = Get-GraphAll $appsUri
                        DLog "  Discovered: $($apps.Count) app(s)" 'Info'

                        foreach ($app in $apps) {
                            $appName = [string]$app['displayName']
                            if (-not [string]::IsNullOrWhiteSpace($keyword)) {
                                if ($appName -notlike "*$keyword*") { continue }
                            }
                            $allRows.Add([PSCustomObject]@{
                                DeviceName         = [string]$rec['DeviceName']
                                EntraDeviceId      = [string]$rec['EntraDeviceId']
                                PrimaryUserUPN     = [string]$rec['UPN']
                                OSVersion          = [string]$rec['OSVersion']
                                DeviceOwnership    = [string]$rec['Ownership']
                                AppName            = $appName
                                AppVersion         = [string]$app['version']
                                AppPublisher       = [string]$app['publisher']
                                ResolvedIntent     = ''
                                InstallationStatus = ''
                                AppDescription     = ''
                                AppSource          = 'Discovered'
                                CollectedAt        = $collectTime
                            })
                        }
                    } catch {
                        DLog "  Discovered apps fetch failed for $($rec['DeviceName']): $($_.Exception.Message)" 'Warning'
                    }
                }

                # ── 3B: Managed Apps (via mobileAppIntentAndStates direct query per device) ──
                if ($showManaged) {
                    $upn = [string]$rec['UPN']
                    $intuneId = [string]$rec['IntuneId']

                    if ([string]::IsNullOrWhiteSpace($upn)) {
                        DLog "  Managed: skipped (no primary user UPN)" 'Warning'
                    } else {
                        try {
                            # Resolve UPN to user GUID (cached)
                            if (-not $userIdCache.ContainsKey($upn)) {
                                $encUpn = [Uri]::EscapeDataString($upn)
                                $userResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/${encUpn}?`$select=id"
                                $userIdCache[$upn] = [string]$userResp['id']
                            }
                            $userId = $userIdCache[$upn]

                            # Direct query: get all managed apps + intent + state for this device
                            $intentUri = "https://graph.microsoft.com/beta/users/${userId}/mobileAppIntentAndStates/${intuneId}"
                            $intentResp = Invoke-MgGraphRequest -Method GET -Uri $intentUri
                            $mobileAppList = @()
                            if ($intentResp.ContainsKey('mobileAppList')) {
                                $mobileAppList = @($intentResp['mobileAppList'])
                            }
                            DLog "  Managed: $($mobileAppList.Count) app(s)" 'Info'

                            foreach ($mApp in $mobileAppList) {
                                $appName = [string]$mApp['displayName']
                                if (-not [string]::IsNullOrWhiteSpace($keyword)) {
                                    if ($appName -notlike "*$keyword*") { continue }
                                }
                                # Fetch app description (cached)
                                $appId = [string]$mApp['applicationId']
                                if ($appId -and -not $appDescCache.ContainsKey($appId)) {
                                    try {
                                        $appDetail = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/${appId}?`$select=description"
                                        $appDescCache[$appId] = [string]$appDetail['description']
                                    } catch {
                                        $appDescCache[$appId] = ''
                                    }
                                }
                                $appDesc = if ($appId -and $appDescCache.ContainsKey($appId)) { $appDescCache[$appId] } else { '' }

                                $rawIntent = [string]$mApp['mobileAppIntent']
                                $rawState  = [string]$mApp['installState']
                                $friendlyIntent = if ($intentMap.ContainsKey($rawIntent)) { $intentMap[$rawIntent] } else { $rawIntent }
                                $friendlyState  = if ($stateMap.ContainsKey($rawState))   { $stateMap[$rawState]   } else { $rawState }

                                $allRows.Add([PSCustomObject]@{
                                    DeviceName         = [string]$rec['DeviceName']
                                    EntraDeviceId      = [string]$rec['EntraDeviceId']
                                    PrimaryUserUPN     = [string]$rec['UPN']
                                    OSVersion          = [string]$rec['OSVersion']
                                    DeviceOwnership    = [string]$rec['Ownership']
                                    AppName            = $appName
                                    AppVersion         = [string]$mApp['displayVersion']
                                    AppPublisher       = ''
                                    ResolvedIntent     = $friendlyIntent
                                    InstallationStatus = $friendlyState
                                    AppDescription     = $appDesc
                                    AppSource          = 'Managed'
                                    CollectedAt        = $collectTime
                                })
                            }
                        } catch {
                            DLog "  Managed apps fetch failed for $($rec['DeviceName']): $($_.Exception.Message)" 'Warning'
                        }
                    }
                }
            }

            DLog "Completed: $($allRows.Count) app record(s) across $($filtered.Count) device(s)." 'Success'
            $DaQueue.Enqueue(@{ Type='done'; Rows=$allRows; Error=$null })

        } catch {
            $DaQueue.Enqueue(@{ Type='done'; Rows=@(); Error=$_.Exception.Message })
        }
    })

    $script:DaHandle = $script:DaPs.BeginInvoke()

    # Timer to drain DaQueue on the UI thread
    $script:DaTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:DaTimer.Interval = [TimeSpan]::FromMilliseconds(300)
    $script:DaTimer.Add_Tick({
        $entry = $null
        while ($script:DaQueue.TryDequeue([ref]$entry)) {
            switch ($entry['Type']) {
                'log' {
                    $lvl = $entry['Level']
                    $col = switch ($lvl) {
                        'Success' { $DevConfig.LogColorSuccess }
                        'Warning' { $DevConfig.LogColorWarning }
                        'Error'   { $DevConfig.LogColorError   }
                        'Action'  { $DevConfig.LogColorAction  }
                        default   { $DevConfig.LogColorInfo    }
                    }
                    try {
                        $para = [System.Windows.Documents.Paragraph]::new()
                        $run  = [System.Windows.Documents.Run]::new($entry['Line'])
                        $run.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                            [System.Windows.Media.ColorConverter]::ConvertFromString($col))
                        $para.Margin = [System.Windows.Thickness]::new(0)
                        $para.Inlines.Add($run)
                        $script:RtbLog.Document.Blocks.Add($para)
                        $script:RtbLog.ScrollToEnd()
                    } catch {}
                }
                'progress' {
                    $script:Window.FindName('PnlDAProgress').Visibility       = 'Visible'
                    $script:Window.FindName('TxtDAProgressMsg').Text          = $entry['Msg']
                    $script:Window.FindName('TxtDAProgressDetail').Text       = $entry['Detail']
                    Update-StatusBar -Text "Get App Info: $($entry['Msg'])"
                }
                'done' {
                    $script:DaTimer.Stop()
                    try { $script:DaPs.EndInvoke($script:DaHandle) } catch {}
                    try { $script:DaRs.Close() }                      catch {}
                    try { $script:DaPs.Dispose() }                    catch {}
                    $script:DaPs = $null; $script:DaRs = $null
                    $script:Window.FindName('BtnStop').Visibility = 'Collapsed'
                    $script:Window.FindName('BtnStop').IsEnabled  = $true
                    $script:Window.FindName('BtnStop').Content = (New-StopBtnContent 'Stop')
                    $script:Window.FindName('BtnDARun').IsEnabled = $true
                    $script:Window.FindName('BtnDARunAll').IsEnabled = $true
                    $script:Window.FindName('BtnDAGetManagedAssignments').IsEnabled = $true
                    $script:Window.FindName('PnlDAProgress').Visibility = 'Collapsed'

                    $errMsg = $entry['Error']
                    if ($errMsg) {
                        Write-VerboseLog "Get App Info error: $errMsg" -Level Error
                        Show-Notification "Error: $errMsg" -BgColor '#F8D7DA' -FgColor '#721C24'
                        Update-StatusBar -Text 'Connected'
                        return
                    }

                    $rows = @($entry['Rows'])
                    if ($rows.Count -eq 0) {
                        Write-VerboseLog 'Get App Info: no results.' -Level Warning
                        $script:Window.FindName('TxtDANoResults').Visibility = 'Visible'
                        Show-Notification 'No discovered apps matched the filters.' -BgColor '#FFF3CD' -FgColor '#7A4800'
                        Update-StatusBar -Text 'Connected'
                        return
                    }

                    $dg = $script:Window.FindName('DgDAResults')
                    $dg.Columns.Clear()
                    foreach ($prop in ($rows[0].PSObject.Properties)) {
                        $col = [System.Windows.Controls.DataGridTextColumn]::new()
                        $col.Header = $prop.Name
                        $col.Binding = [System.Windows.Data.Binding]::new($prop.Name)
                        if ($prop.Name -in @('AppDescription','Description')) {
                            $col.MaxWidth = 250
                            $col.Width = [System.Windows.Controls.DataGridLength]::new(200)
                            $style = [System.Windows.Style]::new([System.Windows.Controls.TextBlock])
                            $style.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.TextBlock]::TextWrappingProperty, [System.Windows.TextWrapping]::NoWrap))
                            $style.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.TextBlock]::TextTrimmingProperty, [System.Windows.TextTrimming]::CharacterEllipsis))
                            $col.ElementStyle = $style
                        }
                        $dg.Columns.Add($col)
                    }
                    $dg.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSCustomObject]]::new($rows)
                    $script:Window.FindName('TxtDACount').Text            = "$($rows.Count) app record(s) found"
                    $script:Window.FindName('PnlDAResults').Visibility     = 'Visible'
                    $script:Window.FindName('BtnDACopyAll').IsEnabled      = $true
                    $script:Window.FindName('TxtDAFilter').IsEnabled  = $true
                    $script:Window.FindName('BtnDAFilter').IsEnabled  = $true
                    $script:Window.FindName('BtnDAFilterClear').IsEnabled = $true
                    $script:Window.FindName('BtnDAExportXlsx').IsEnabled   = $true
                    Write-VerboseLog "Get App Info: $($rows.Count) record(s) loaded." -Level Success
                    Show-Notification "$($rows.Count) app record(s) ready." -BgColor '#D4EDDA' -FgColor '#155724'
                    Update-StatusBar -Text 'Connected'
                }
            }
        }
    })
    $script:DaTimer.Start()
})

# -- BtnDARunAll: Get Apps from All Devices --
$script:Window.FindName('BtnDARunAll').Add_Click({
    Hide-Notification

    # Collect filters on UI thread (no input list or keyword needed)
    $platFilter = [System.Collections.Generic.List[string]]::new()
    if ($script:Window.FindName('ChkDAWindows').IsChecked) { $platFilter.Add('Windows') }
    if ($script:Window.FindName('ChkDAAndroid').IsChecked) { $platFilter.Add('Android') }
    if ($script:Window.FindName('ChkDAiOS').IsChecked)     { $platFilter.Add('iOS') }
    if ($script:Window.FindName('ChkDAMacOS').IsChecked)   { $platFilter.Add('macOS') }

    $ownershipSel = $script:Window.FindName('CmbDAOwnership').SelectedItem.Content
    $ownershipFilter = switch ($ownershipSel) {
        'Company only'  { 'Company' }
        'Personal only' { 'Personal' }
        default         { 'All' }
    }

    $script:DAParams = @{
        AllDevices      = $true
        PlatformFilter  = $platFilter
        OwnershipFilter = $ownershipFilter
        ShowDiscovered  = $script:Window.FindName('ChkDADiscoveredApps').IsChecked -eq $true
        ShowManaged     = $script:Window.FindName('ChkDAManagedApps').IsChecked -eq $true
    }

    # Reset UI
    $script:Window.FindName('PnlDAResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtDANoResults').Visibility = 'Collapsed'
    $script:Window.FindName('BtnDACopyValue').IsEnabled  = $false
    $script:Window.FindName('BtnDACopyRow').IsEnabled    = $false
    $script:Window.FindName('BtnDACopyAll').IsEnabled    = $false
    $script:Window.FindName('TxtDAFilter').IsEnabled  = $false
    $script:Window.FindName('BtnDAFilter').IsEnabled  = $false
    $script:Window.FindName('BtnDAFilterClear').IsEnabled = $false
    $script:Window.FindName('TxtDAFilter').Text       = ''
    $script:DAAllData = $null
    $script:Window.FindName('BtnDAExportXlsx').IsEnabled = $false
    $dg = $script:Window.FindName('DgDAResults')
    if ($dg) { $dg.ItemsSource = $null }
    $script:Window.FindName('PnlDAProgress').Visibility      = 'Visible'
    $script:Window.FindName('TxtDAProgressMsg').Text         = 'Fetching all managed devices...'
    $script:Window.FindName('TxtDAProgressDetail').Text      = ''

    $script:Window.FindName('BtnDARun').IsEnabled    = $false
    $script:Window.FindName('BtnDARunAll').IsEnabled = $false
    $script:Window.FindName('BtnDAGetManagedAssignments').IsEnabled = $false
    $script:Window.FindName('BtnStop').Visibility = 'Visible'
    $script:Window.FindName('BtnStop').IsEnabled  = $true
    $script:Window.FindName('BtnStop').Content = (New-StopBtnContent 'Stop')
    $script:BgStopped = $false
    Write-VerboseLog '--- Get App Assignments (All Devices): starting ---' -Level Action

    $script:DaQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $script:DaStop  = [hashtable]::Synchronized(@{ Stop = $false })

    $script:DaRs = [runspacefactory]::CreateRunspace()
    $script:DaRs.ApartmentState = 'STA'
    $script:DaRs.ThreadOptions  = 'ReuseThread'
    $script:DaRs.Open()
    $script:DaRs.SessionStateProxy.SetVariable('DaQueue',  $script:DaQueue)
    $script:DaRs.SessionStateProxy.SetVariable('DaStop',   $script:DaStop)
    $script:DaRs.SessionStateProxy.SetVariable('DaParams', $script:DAParams)
    $script:DaRs.SessionStateProxy.SetVariable('LogFile',  $script:LogFile)

    $script:DaPs = [powershell]::Create()
    $script:DaPs.Runspace = $script:DaRs
    $null = $script:DaPs.AddScript({
        function DLog {
            param([string]$Msg, [string]$Level = 'Info')
            $prefix = switch ($Level) {
                "Success" { "$([char]0x2713) " } "Warning" { "[!] " } "Error" { "$([char]0x2717) " }
                'Action'  { '> ' } default   { '  ' }
            }
            $line = "[$(Get-Date -Format 'HH:mm:ss')] $prefix$Msg"
            $DaQueue.Enqueue(@{ Type='log'; Line=$line; Level=$Level })
            if ($LogFile) {
                try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
            }
        }

        function Get-GraphAll([string]$Uri) {
            $items = [System.Collections.Generic.List[object]]::new()
            $next  = $Uri
            do {
                $r = Invoke-MgGraphRequest -Method GET -Uri $next
                if ($r['value']) { foreach ($i in @($r['value'])) { $items.Add($i) } }
                $next = if ($r.ContainsKey('@odata.nextLink')) { $r['@odata.nextLink'] } else { $null }
            } while ($next)
            return ,$items
        }

        function Normalize-OS([string]$os) {
            switch -Wildcard ($os) {
                'Windows*' { 'Windows' } 'Android*' { 'Android' }
                'iPhone'   { 'iOS'     } 'iPadOS'   { 'iOS'     }
                'iOS'      { 'iOS'     } 'MacMDM'   { 'macOS'   }
                'Mac OS X' { 'macOS'   } 'macOS'    { 'macOS'   }
                default    { $os }
            }
        }

        try {
            $p = $DaParams
            $platFilter  = $p['PlatformFilter']
            $ownerFilter = $p['OwnershipFilter']
            $showDiscovered = $p['ShowDiscovered']
            $showManaged    = $p['ShowManaged']

            DLog "Fetching all managed devices from Intune..." 'Action'
            $DaQueue.Enqueue(@{ Type='progress'; Msg="Step 1/3: Fetching all managed devices..."; Detail='' })

            $allDevices = Get-GraphAll "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=id,deviceName,operatingSystem,osVersion,managedDeviceOwnerType,azureADDeviceId,userPrincipalName&`$top=999"
            DLog "Step 1 complete: $($allDevices.Count) managed device(s) retrieved" 'Info'

            if ($allDevices.Count -eq 0) {
                $DaQueue.Enqueue(@{ Type='done'; Rows=@(); Error='No managed devices found in Intune.' })
                return
            }

            $deviceRecords = [System.Collections.Generic.List[hashtable]]::new()
            foreach ($iDev in $allDevices) {
                $deviceRecords.Add(@{
                    IntuneId      = [string]$iDev['id']
                    DeviceName    = [string]$iDev['deviceName']
                    EntraDeviceId = [string]$iDev['azureADDeviceId']
                    OS            = [string]$iDev['operatingSystem']
                    OSVersion     = [string]$iDev['osVersion']
                    Ownership     = [string]$iDev['managedDeviceOwnerType']
                    UPN           = [string]$iDev['userPrincipalName']
                })
            }

            $DaQueue.Enqueue(@{ Type='progress'; Msg="Step 2/3: Applying filters..."; Detail='' })
            $filtered = [System.Collections.Generic.List[hashtable]]::new()
            foreach ($rec in $deviceRecords) {
                $osNorm = Normalize-OS $rec['OS']
                if ($platFilter.Count -gt 0 -and $osNorm -notin $platFilter) { continue }
                if ($ownerFilter -ne 'All' -and $rec['Ownership'] -ne $ownerFilter) { continue }
                $filtered.Add($rec)
            }
            DLog "Step 2 complete: $($filtered.Count) device(s) after filtering (from $($deviceRecords.Count))" 'Info'

            if ($filtered.Count -eq 0) {
                $DaQueue.Enqueue(@{ Type='done'; Rows=@(); Error='No devices matched the platform/ownership filters.' })
                return
            }




            # ── STEP 3: Fetch apps per device ──
            $allRows = [System.Collections.Generic.List[PSCustomObject]]::new()

            # Intent and state mappings for managed apps
            $intentMap = @{
                'available'                       = 'Available'
                'notAvailable'                    = 'Not Available'
                'requiredInstall'                 = 'Required Install'
                'requiredUninstall'               = 'Required Uninstall'
                'requiredAndAvailableInstall'     = 'Required and Available Install'
                'availableInstallWithoutEnrollment' = 'Available (No Enrollment)'
                'exclude'                         = 'Exclude'
            }
            $stateMap = @{
                'installed'      = 'Installed'
                'failed'         = 'Failed'
                'notInstalled'   = 'Not Installed'
                'uninstallFailed'= 'Uninstall Failed'
                'pendingInstall' = 'Pending Install'
                'unknown'        = 'Unknown'
                'notApplicable'  = 'Not Applicable'
            }

            # User ID cache (UPN -> GUID) to avoid redundant lookups
            $userIdCache = @{}

            # App description cache (applicationId -> description) to avoid redundant lookups
            $appDescCache = @{}

            $devIdx = 0
            foreach ($rec in $filtered) {
                if ($DaStop['Stop']) { DLog 'Stop requested - halting.' 'Warning'; break }
                $devIdx++
                $DaQueue.Enqueue(@{ Type='progress'; Msg="Step 3/3: Fetching apps for device $devIdx of $($filtered.Count)..."; Detail=$rec['DeviceName'] })
                DLog "[$devIdx] Fetching apps for: $($rec['DeviceName'])" 'Info'

                $collectTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

                # ── 3A: Discovered Apps (detectedApps API - per device) ──
                if ($showDiscovered) {
                    try {
                        $appsUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($rec['IntuneId'])/detectedApps?`$select=displayName,version,platform,publisher"
                        $apps = Get-GraphAll $appsUri
                        DLog "  Discovered: $($apps.Count) app(s)" 'Info'

                        foreach ($app in $apps) {
                            $appName = [string]$app['displayName']
                            $allRows.Add([PSCustomObject]@{
                                DeviceName         = [string]$rec['DeviceName']
                                EntraDeviceId      = [string]$rec['EntraDeviceId']
                                PrimaryUserUPN     = [string]$rec['UPN']
                                OSVersion          = [string]$rec['OSVersion']
                                DeviceOwnership    = [string]$rec['Ownership']
                                AppName            = $appName
                                AppVersion         = [string]$app['version']
                                AppPublisher       = [string]$app['publisher']
                                ResolvedIntent     = ''
                                InstallationStatus = ''
                                AppDescription     = ''
                                AppSource          = 'Discovered'
                                CollectedAt        = $collectTime
                            })
                        }
                    } catch {
                        DLog "  Discovered apps fetch failed for $($rec['DeviceName']): $($_.Exception.Message)" 'Warning'
                    }
                }

                # ── 3B: Managed Apps (via mobileAppIntentAndStates direct query per device) ──
                if ($showManaged) {
                    $upn = [string]$rec['UPN']
                    $intuneId = [string]$rec['IntuneId']

                    if ([string]::IsNullOrWhiteSpace($upn)) {
                        DLog "  Managed: skipped (no primary user UPN)" 'Warning'
                    } else {
                        try {
                            # Resolve UPN to user GUID (cached)
                            if (-not $userIdCache.ContainsKey($upn)) {
                                $encUpn = [Uri]::EscapeDataString($upn)
                                $userResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/${encUpn}?`$select=id"
                                $userIdCache[$upn] = [string]$userResp['id']
                            }
                            $userId = $userIdCache[$upn]

                            # Direct query: get all managed apps + intent + state for this device
                            $intentUri = "https://graph.microsoft.com/beta/users/${userId}/mobileAppIntentAndStates/${intuneId}"
                            $intentResp = Invoke-MgGraphRequest -Method GET -Uri $intentUri
                            $mobileAppList = @()
                            if ($intentResp.ContainsKey('mobileAppList')) {
                                $mobileAppList = @($intentResp['mobileAppList'])
                            }
                            DLog "  Managed: $($mobileAppList.Count) app(s)" 'Info'

                            foreach ($mApp in $mobileAppList) {
                                $appName = [string]$mApp['displayName']
                                $rawIntent = [string]$mApp['mobileAppIntent']
                                # Fetch app description (cached)
                                $appId = [string]$mApp['applicationId']
                                if ($appId -and -not $appDescCache.ContainsKey($appId)) {
                                    try {
                                        $appDetail = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/${appId}?`$select=description"
                                        $appDescCache[$appId] = [string]$appDetail['description']
                                    } catch {
                                        $appDescCache[$appId] = ''
                                    }
                                }
                                $appDesc = if ($appId -and $appDescCache.ContainsKey($appId)) { $appDescCache[$appId] } else { '' }

                                $rawState  = [string]$mApp['installState']
                                $friendlyIntent = if ($intentMap.ContainsKey($rawIntent)) { $intentMap[$rawIntent] } else { $rawIntent }
                                $friendlyState  = if ($stateMap.ContainsKey($rawState))   { $stateMap[$rawState]   } else { $rawState }

                                $allRows.Add([PSCustomObject]@{
                                    DeviceName         = [string]$rec['DeviceName']
                                    EntraDeviceId      = [string]$rec['EntraDeviceId']
                                    PrimaryUserUPN     = [string]$rec['UPN']
                                    OSVersion          = [string]$rec['OSVersion']
                                    DeviceOwnership    = [string]$rec['Ownership']
                                    AppName            = $appName
                                    AppVersion         = [string]$mApp['displayVersion']
                                    AppPublisher       = ''
                                    ResolvedIntent     = $friendlyIntent
                                    InstallationStatus = $friendlyState
                                    AppDescription     = $appDesc
                                    AppSource          = 'Managed'
                                    CollectedAt        = $collectTime
                                })
                            }
                        } catch {
                            DLog "  Managed apps fetch failed for $($rec['DeviceName']): $($_.Exception.Message)" 'Warning'
                        }
                    }
                }
            }

            DLog "Completed: $($allRows.Count) app record(s) across $($filtered.Count) device(s)." 'Success'
            $DaQueue.Enqueue(@{ Type='done'; Rows=$allRows; Error=$null })

        } catch {
            $DaQueue.Enqueue(@{ Type='done'; Rows=@(); Error=$_.Exception.Message })
        }
    })

    $script:DaHandle = $script:DaPs.BeginInvoke()
    # Timer to drain DaQueue on the UI thread
    $script:DaTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:DaTimer.Interval = [TimeSpan]::FromMilliseconds(300)
    $script:DaTimer.Add_Tick({
        $entry = $null
        while ($script:DaQueue.TryDequeue([ref]$entry)) {
            switch ($entry['Type']) {
                'log' {
                    $lvl = $entry['Level']
                    $col = switch ($lvl) {
                        'Success' { $DevConfig.LogColorSuccess }
                        'Warning' { $DevConfig.LogColorWarning }
                        'Error'   { $DevConfig.LogColorError   }
                        'Action'  { $DevConfig.LogColorAction  }
                        default   { $DevConfig.LogColorInfo    }
                    }
                    try {
                        $para = [System.Windows.Documents.Paragraph]::new()
                        $run  = [System.Windows.Documents.Run]::new($entry['Line'])
                        $run.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                            [System.Windows.Media.ColorConverter]::ConvertFromString($col))
                        $para.Margin = [System.Windows.Thickness]::new(0)
                        $para.Inlines.Add($run)
                        $script:RtbLog.Document.Blocks.Add($para)
                        $script:RtbLog.ScrollToEnd()
                    } catch {}
                }
                'progress' {
                    $script:Window.FindName('PnlDAProgress').Visibility       = 'Visible'
                    $script:Window.FindName('TxtDAProgressMsg').Text          = $entry['Msg']
                    $script:Window.FindName('TxtDAProgressDetail').Text       = $entry['Detail']
                    Update-StatusBar -Text "Get App Info (All): $($entry['Msg'])"
                }
                'done' {
                    $script:DaTimer.Stop()
                    try { $script:DaPs.EndInvoke($script:DaHandle) } catch {}
                    try { $script:DaRs.Close() }                      catch {}
                    try { $script:DaPs.Dispose() }                    catch {}
                    $script:DaPs = $null; $script:DaRs = $null
                    $script:Window.FindName('BtnStop').Visibility = 'Collapsed'
                    $script:Window.FindName('BtnStop').IsEnabled  = $true
                    $script:Window.FindName('BtnStop').Content = (New-StopBtnContent 'Stop')
                    $script:Window.FindName('BtnDARun').IsEnabled    = $true
                    $script:Window.FindName('BtnDARunAll').IsEnabled = $true
                    $script:Window.FindName('BtnDAGetManagedAssignments').IsEnabled = $true
                    $script:Window.FindName('PnlDAProgress').Visibility = 'Collapsed'

                    $errMsg = $entry['Error']
                    if ($errMsg) {
                        Write-VerboseLog "Get App Info (All) error: $errMsg" -Level Error
                        Show-Notification "Error: $errMsg" -BgColor '#F8D7DA' -FgColor '#721C24'
                        Update-StatusBar -Text 'Connected'
                        return
                    }

                    $rows = @($entry['Rows'])
                    if ($rows.Count -eq 0) {
                        Write-VerboseLog 'Get App Info (All): no results.' -Level Warning
                        $script:Window.FindName('TxtDANoResults').Visibility = 'Visible'
                        Show-Notification 'No discovered apps matched the filters.' -BgColor '#FFF3CD' -FgColor '#7A4800'
                        Update-StatusBar -Text 'Connected'
                        return
                    }

                    $dg = $script:Window.FindName('DgDAResults')
                    $dg.Columns.Clear()
                    foreach ($prop in ($rows[0].PSObject.Properties)) {
                        $col = [System.Windows.Controls.DataGridTextColumn]::new()
                        $col.Header = $prop.Name
                        $col.Binding = [System.Windows.Data.Binding]::new($prop.Name)
                        if ($prop.Name -in @('AppDescription','Description')) {
                            $col.MaxWidth = 250
                            $col.Width = [System.Windows.Controls.DataGridLength]::new(200)
                            $style = [System.Windows.Style]::new([System.Windows.Controls.TextBlock])
                            $style.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.TextBlock]::TextWrappingProperty, [System.Windows.TextWrapping]::NoWrap))
                            $style.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.TextBlock]::TextTrimmingProperty, [System.Windows.TextTrimming]::CharacterEllipsis))
                            $col.ElementStyle = $style
                        }
                        $dg.Columns.Add($col)
                    }
                    $dg.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSCustomObject]]::new($rows)
                    $script:Window.FindName('TxtDACount').Text            = "$($rows.Count) app record(s) found (all devices)"
                    $script:Window.FindName('PnlDAResults').Visibility     = 'Visible'
                    $script:Window.FindName('BtnDACopyAll').IsEnabled      = $true
                    $script:Window.FindName('TxtDAFilter').IsEnabled  = $true
                    $script:Window.FindName('BtnDAFilter').IsEnabled  = $true
                    $script:Window.FindName('BtnDAFilterClear').IsEnabled = $true
                    $script:Window.FindName('BtnDAExportXlsx').IsEnabled   = $true
                    Write-VerboseLog "Get App Info (All): $($rows.Count) record(s) loaded." -Level Success
                    Show-Notification "$($rows.Count) app record(s) ready (all devices)." -BgColor '#D4EDDA' -FgColor '#155724'
                    Update-StatusBar -Text 'Connected'
                }
            }
        }
    })
    $script:DaTimer.Start()
})


# -- BtnDAGetManagedAssignments: Tenant-wide Managed App Assignments --
$script:Window.FindName('BtnDAGetManagedAssignments').Add_Click({
    Hide-Notification

    # Respect OS Platform checkboxes and App Keyword filter
    $platFilter = [System.Collections.Generic.List[string]]::new()
    if ($script:Window.FindName('ChkDAWindows').IsChecked) { $platFilter.Add('Windows') }
    if ($script:Window.FindName('ChkDAAndroid').IsChecked) { $platFilter.Add('Android') }
    if ($script:Window.FindName('ChkDAiOS').IsChecked)     { $platFilter.Add('iOS') }
    if ($script:Window.FindName('ChkDAMacOS').IsChecked)   { $platFilter.Add('macOS') }

    $script:DAParams = @{
        ManagedAssignments = $true
        PlatformFilter     = $platFilter
        Keyword            = [string]$script:Window.FindName('TxtDAKeyword').Text
    }

    # Reset UI
    $script:Window.FindName('PnlDAResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtDANoResults').Visibility = 'Collapsed'
    $script:Window.FindName('BtnDACopyValue').IsEnabled  = $false
    $script:Window.FindName('BtnDACopyRow').IsEnabled    = $false
    $script:Window.FindName('BtnDACopyAll').IsEnabled    = $false
    $script:Window.FindName('TxtDAFilter').IsEnabled  = $false
    $script:Window.FindName('BtnDAFilter').IsEnabled  = $false
    $script:Window.FindName('BtnDAFilterClear').IsEnabled = $false
    $script:Window.FindName('TxtDAFilter').Text       = ''
    $script:DAAllData = $null
    $script:Window.FindName('BtnDAExportXlsx').IsEnabled = $false
    $dg = $script:Window.FindName('DgDAResults')
    if ($dg) { $dg.ItemsSource = $null }
    $script:Window.FindName('PnlDAProgress').Visibility      = 'Visible'
    $script:Window.FindName('TxtDAProgressMsg').Text         = 'Fetching managed app assignments...'
    $script:Window.FindName('TxtDAProgressDetail').Text      = ''

    $script:Window.FindName('BtnDARun').IsEnabled    = $false
    $script:Window.FindName('BtnDARunAll').IsEnabled = $false
    $script:Window.FindName('BtnDAGetManagedAssignments').IsEnabled = $false
    $script:Window.FindName('BtnStop').Visibility = 'Visible'
    $script:Window.FindName('BtnStop').IsEnabled  = $true
    $script:Window.FindName('BtnStop').Content = (New-StopBtnContent 'Stop')
    $script:BgStopped = $false
    Write-VerboseLog '--- All Managed App Assignments: starting ---' -Level Action

    $script:DaQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $script:DaStop  = [hashtable]::Synchronized(@{ Stop = $false })

    $script:DaRs = [runspacefactory]::CreateRunspace()
    $script:DaRs.ApartmentState = 'STA'
    $script:DaRs.ThreadOptions  = 'ReuseThread'
    $script:DaRs.Open()
    $script:DaRs.SessionStateProxy.SetVariable('DaQueue',  $script:DaQueue)
    $script:DaRs.SessionStateProxy.SetVariable('DaStop',   $script:DaStop)
    $script:DaRs.SessionStateProxy.SetVariable('DaParams', $script:DAParams)
    $script:DaRs.SessionStateProxy.SetVariable('LogFile',  $script:LogFile)

    $script:DaPs = [powershell]::Create()
    $script:DaPs.Runspace = $script:DaRs

    $null = $script:DaPs.AddScript({
        function DLog {
            param([string]$Msg, [string]$Level = 'Info')
            $prefix = switch ($Level) {
                "Success" { "$([char]0x2713) " } "Warning" { "[!] " } "Error" { "$([char]0x2717) " }
                'Action'  { '> ' } default   { '  ' }
            }
            $line = "[$(Get-Date -Format 'HH:mm:ss')] $prefix$Msg"
            $DaQueue.Enqueue(@{ Type='log'; Line=$line; Level=$Level })
            if ($LogFile) {
                try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
            }
        }

        function Get-GraphAll([string]$Uri) {
            $items = [System.Collections.Generic.List[object]]::new()
            $next  = $Uri
            do {
                $r = Invoke-MgGraphRequest -Method GET -Uri $next
                if ($r['value']) { foreach ($i in @($r['value'])) { $items.Add($i) } }
                $next = if ($r.ContainsKey('@odata.nextLink')) { $r['@odata.nextLink'] } else { $null }
            } while ($next)
            return ,$items
        }

        $typeMap = @{
            '#microsoft.graph.win32LobApp'                = 'Win32 LOB'
            '#microsoft.graph.windowsMobileMSI'           = 'Windows MSI'
            '#microsoft.graph.microsoftStoreForBusinessApp'= 'MS Store for Business'
            '#microsoft.graph.winGetApp'                  = 'WinGet'
            '#microsoft.graph.officeSuiteApp'             = 'Microsoft 365 Apps'
            '#microsoft.graph.windowsWebApp'              = 'Windows Web App'
            '#microsoft.graph.windowsUniversalAppX'       = 'Windows AppX/MSIX'
            '#microsoft.graph.windowsStoreApp'            = 'Windows Store App'
            '#microsoft.graph.webApp'                     = 'Web Link'
            '#microsoft.graph.iosVppApp'                  = 'iOS VPP'
            '#microsoft.graph.iosStoreApp'                = 'iOS Store'
            '#microsoft.graph.iosLobApp'                  = 'iOS LOB'
            '#microsoft.graph.managedIOSStoreApp'         = 'Managed iOS Store'
            '#microsoft.graph.managedIOSLobApp'           = 'Managed iOS LOB'
            '#microsoft.graph.managedAndroidStoreApp'     = 'Managed Android Store'
            '#microsoft.graph.managedAndroidLobApp'       = 'Managed Android LOB'
            '#microsoft.graph.androidStoreApp'            = 'Android Store'
            '#microsoft.graph.androidLobApp'              = 'Android LOB'
            '#microsoft.graph.androidManagedStoreApp'     = 'Android Enterprise'
            '#microsoft.graph.androidForWorkApp'          = 'Android for Work'
            '#microsoft.graph.macOSLobApp'                = 'macOS LOB'
            '#microsoft.graph.macOSDmgApp'                = 'macOS DMG'
            '#microsoft.graph.macOSPkgApp'                = 'macOS PKG'
            '#microsoft.graph.macOSMicrosoftEdgeApp'      = 'macOS Edge'
            '#microsoft.graph.macOSMicrosoftDefenderApp'  = 'macOS Defender'
            '#microsoft.graph.macOSOfficeSuiteApp'        = 'macOS Office Suite'
            '#microsoft.graph.macOSWebClip'               = 'macOS Web Clip'
        }
        $platMap = @{
            '#microsoft.graph.win32LobApp'='Windows';'#microsoft.graph.windowsMobileMSI'='Windows';'#microsoft.graph.microsoftStoreForBusinessApp'='Windows'
            '#microsoft.graph.winGetApp'='Windows';'#microsoft.graph.officeSuiteApp'='Windows';'#microsoft.graph.windowsWebApp'='Windows'
            '#microsoft.graph.windowsUniversalAppX'='Windows';'#microsoft.graph.windowsStoreApp'='Windows';'#microsoft.graph.webApp'='Cross-platform'
            '#microsoft.graph.iosVppApp'='iOS';'#microsoft.graph.iosStoreApp'='iOS';'#microsoft.graph.iosLobApp'='iOS'
            '#microsoft.graph.managedIOSStoreApp'='iOS';'#microsoft.graph.managedIOSLobApp'='iOS'
            '#microsoft.graph.managedAndroidStoreApp'='Android';'#microsoft.graph.managedAndroidLobApp'='Android'
            '#microsoft.graph.androidStoreApp'='Android';'#microsoft.graph.androidLobApp'='Android'
            '#microsoft.graph.androidManagedStoreApp'='Android';'#microsoft.graph.androidForWorkApp'='Android'
            '#microsoft.graph.macOSLobApp'='macOS';'#microsoft.graph.macOSDmgApp'='macOS';'#microsoft.graph.macOSPkgApp'='macOS'
            '#microsoft.graph.macOSMicrosoftEdgeApp'='macOS';'#microsoft.graph.macOSMicrosoftDefenderApp'='macOS'
            '#microsoft.graph.macOSOfficeSuiteApp'='macOS';'#microsoft.graph.macOSWebClip'='macOS'
        }

        try {
            $p = $DaParams
            $platFilter = $p['PlatformFilter']
            $keyword = [string]$p['Keyword']
            $groupCache  = @{}
            $filterCache = @{}

            function Resolve-GroupName([string]$GroupId) {
                if ($groupCache.ContainsKey($GroupId)) { return $groupCache[$GroupId] }
                try {
                    $gr = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId`?`$select=displayName"
                    $gn = [string]$gr['displayName']; $groupCache[$GroupId] = $gn; return $gn
                } catch { $groupCache[$GroupId] = $GroupId; return $GroupId }
            }

            function Resolve-FilterName([string]$FilterId) {
                if (-not $FilterId -or $FilterId -eq '00000000-0000-0000-0000-000000000000') { return '' }
                if ($filterCache.ContainsKey($FilterId)) { return $filterCache[$FilterId] }
                try {
                    $fr = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/assignmentFilters/$FilterId`?`$select=displayName"
                    $fn = [string]$fr['displayName']; $filterCache[$FilterId] = $fn; return $fn
                } catch { $filterCache[$FilterId] = $FilterId; return $FilterId }
            }

            # -- STEP 1: Fetch all managed apps --
            DLog "Step 1/3: Fetching managed apps from Intune..." 'Action'
            $DaQueue.Enqueue(@{ Type='progress'; Msg="Step 1/3: Fetching managed apps..."; Detail='' })
            $mobileApps = Get-GraphAll "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$select=id,displayName,publisher,description,createdDateTime,lastModifiedDateTime&`$top=999"
            DLog "Step 1 complete: $($mobileApps.Count) managed app(s)" 'Info'

            if ($mobileApps.Count -eq 0) {
                $DaQueue.Enqueue(@{ Type='done'; Rows=@(); Error='No managed apps found in Intune.' })
                return
            }

            # -- STEP 2: Apply platform filter --
            $DaQueue.Enqueue(@{ Type='progress'; Msg="Step 2/3: Filtering by platform..."; Detail='' })
            $filteredApps = [System.Collections.Generic.List[object]]::new()
            foreach ($app in $mobileApps) {
                $odataType = [string]$app['@odata.type']
                $appPlat = if ($platMap.ContainsKey($odataType)) { $platMap[$odataType] } else { 'Unknown' }
                if ($platFilter.Count -gt 0 -and $appPlat -ne 'Cross-platform' -and $appPlat -notin $platFilter) { continue }
                if (-not [string]::IsNullOrWhiteSpace($keyword)) { if ([string]$app['displayName'] -notlike "*$keyword*") { continue } }
                $filteredApps.Add($app)
            }
            DLog "Step 2 complete: $($filteredApps.Count) app(s) after platform/keyword filter" 'Info'

            if ($filteredApps.Count -eq 0) {
                $DaQueue.Enqueue(@{ Type='done'; Rows=@(); Error='No managed apps matched the platform/keyword filter.' })
                return
            }

            # -- STEP 3: Fetch assignments per app --
            $allRows = [System.Collections.Generic.List[PSCustomObject]]::new()
            $appIdx = 0

            foreach ($app in $filteredApps) {
                if ($DaStop['Stop']) { DLog 'Stop requested - halting.' 'Warning'; break }
                $appIdx++
                $appId     = [string]$app['id']
                $appName   = [string]$app['displayName']
                $odataType = [string]$app['@odata.type']
                $appType   = if ($typeMap.ContainsKey($odataType)) { $typeMap[$odataType] } else { $odataType -replace '^#microsoft\.graph\.','' }
                $appPlat   = if ($platMap.ContainsKey($odataType)) { $platMap[$odataType] } else { 'Unknown' }
                $appPub    = [string]$app['publisher']
                $appDesc     = [string]$app['description']
                $appVersion  = [string]$app['displayVersion']
                $appCreated  = if ($app['createdDateTime'])      { ([DateTime]$app['createdDateTime']).ToString('yyyy-MM-dd HH:mm') }      else { '' }
                $appModified = if ($app['lastModifiedDateTime']) { ([DateTime]$app['lastModifiedDateTime']).ToString('yyyy-MM-dd HH:mm') } else { '' }

                if ($appIdx % 25 -eq 1 -or $appIdx -eq $filteredApps.Count) {
                    $DaQueue.Enqueue(@{ Type='progress'; Msg="Step 3/3: Fetching assignments ($appIdx of $($filteredApps.Count))..."; Detail=$appName })
                }

                try {
                    $assignUri  = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$appId/assignments"
                    $assignResp = Invoke-MgGraphRequest -Method GET -Uri $assignUri
                    $assignments = @($assignResp['value'])

                    if ($assignments.Count -eq 0) {
                        $allRows.Add([PSCustomObject]@{
                            AppName        = $appName
                            AppType        = $appType
                            Platform       = $appPlat
                            Publisher      = $appPub
                            AppVersion     = $appVersion
                            Description    = $appDesc
                            DateCreated    = $appCreated
                            DateModified   = $appModified
                            Intent         = '(Unassigned)'
                            IncludedGroup  = ''
                            ExcludedGroup  = ''
                            IncludeFilter  = ''
                            ExcludeFilter  = ''
                        })
                        continue
                    }

                    foreach ($asn in $assignments) {
                        $rawIntent = [string]$asn['intent']
                        $intent = switch ($rawIntent) {
                            'available'                  { 'Available' }
                            'required'                   { 'Required' }
                            'uninstall'                  { 'Uninstall' }
                            'availableWithoutEnrollment' { 'Available (No Enrollment)' }
                            default                      { $rawIntent }
                        }

                        $target = $asn['target']
                        if (-not $target) { continue }
                        $otype     = [string]$target['@odata.type']
                        $isExclude = $otype -match 'exclusionGroup'

                        $groupId    = [string]$target['groupId']
                        $targetName = ''
                        if ($groupId) {
                            $targetName = Resolve-GroupName -GroupId $groupId
                        } elseif ($otype -match 'allDevices') {
                            $targetName = '[All Devices]'
                        } elseif ($otype -match 'allLicensedUsers|allUsers') {
                            $targetName = '[All Users]'
                        }

                        $filterId   = [string]$target['deviceAndAppManagementAssignmentFilterId']
                        $filterType = [string]$target['deviceAndAppManagementAssignmentFilterType']
                        $filterName = Resolve-FilterName -FilterId $filterId
                        $inclFilter = if ($filterType -eq 'include') { $filterName } else { '' }
                        $exclFilter = if ($filterType -eq 'exclude') { $filterName } else { '' }

                        $allRows.Add([PSCustomObject]@{
                            AppName        = $appName
                            AppType        = $appType
                            Platform       = $appPlat
                            Publisher      = $appPub
                            AppVersion     = $appVersion
                            Description    = $appDesc
                            DateCreated    = $appCreated
                            DateModified   = $appModified
                            Intent         = $intent
                            IncludedGroup  = if (-not $isExclude) { $targetName } else { '' }
                            ExcludedGroup  = if ($isExclude) { $targetName } else { '' }
                            IncludeFilter  = $inclFilter
                            ExcludeFilter  = $exclFilter
                        })
                    }
                } catch {
                    DLog "  Assignment fetch failed for ${appName}: $($_.Exception.Message)" 'Warning'
                }
            }

            DLog "Completed: $($allRows.Count) assignment record(s) across $($filteredApps.Count) app(s)." 'Success'
            $DaQueue.Enqueue(@{ Type='done'; Rows=$allRows; Error=$null })

        } catch {
            $DaQueue.Enqueue(@{ Type='done'; Rows=@(); Error=$_.Exception.Message })
        }
    })

    $script:DaHandle = $script:DaPs.BeginInvoke()
    # Timer to drain DaQueue on the UI thread
    $script:DaTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:DaTimer.Interval = [TimeSpan]::FromMilliseconds(300)
    $script:DaTimer.Add_Tick({
        $entry = $null
        while ($script:DaQueue.TryDequeue([ref]$entry)) {
            switch ($entry['Type']) {
                'log' {
                    $lvl = $entry['Level']
                    $col = switch ($lvl) {
                        'Success' { $DevConfig.LogColorSuccess }
                        'Warning' { $DevConfig.LogColorWarning }
                        'Error'   { $DevConfig.LogColorError   }
                        'Action'  { $DevConfig.LogColorAction  }
                        default   { $DevConfig.LogColorInfo    }
                    }
                    try {
                        $para = [System.Windows.Documents.Paragraph]::new()
                        $run  = [System.Windows.Documents.Run]::new($entry['Line'])
                        $run.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                            [System.Windows.Media.ColorConverter]::ConvertFromString($col))
                        $para.Margin = [System.Windows.Thickness]::new(0)
                        $para.Inlines.Add($run)
                        $script:RtbLog.Document.Blocks.Add($para)
                        $script:RtbLog.ScrollToEnd()
                    } catch {}
                }
                'progress' {
                    $script:Window.FindName('PnlDAProgress').Visibility       = 'Visible'
                    $script:Window.FindName('TxtDAProgressMsg').Text          = $entry['Msg']
                    $script:Window.FindName('TxtDAProgressDetail').Text       = $entry['Detail']
                    Update-StatusBar -Text "Managed App Assignments: $($entry['Msg'])"
                }
                'done' {
                    $script:DaTimer.Stop()
                    try { $script:DaPs.EndInvoke($script:DaHandle) } catch {}
                    try { $script:DaRs.Close() }                      catch {}
                    try { $script:DaPs.Dispose() }                    catch {}
                    $script:DaPs = $null; $script:DaRs = $null
                    $script:Window.FindName('BtnStop').Visibility = 'Collapsed'
                    $script:Window.FindName('BtnStop').IsEnabled  = $true
                    $script:Window.FindName('BtnStop').Content = (New-StopBtnContent 'Stop')
                    $script:Window.FindName('BtnDARun').IsEnabled    = $true
                    $script:Window.FindName('BtnDARunAll').IsEnabled = $true
                    $script:Window.FindName('BtnDAGetManagedAssignments').IsEnabled = $true
                    $script:Window.FindName('PnlDAProgress').Visibility = 'Collapsed'

                    $errMsg = $entry['Error']
                    if ($errMsg) {
                        Write-VerboseLog "Managed App Assignments error: $errMsg" -Level Error
                        Show-Notification "Error: $errMsg" -BgColor '#F8D7DA' -FgColor '#721C24'
                        Update-StatusBar -Text 'Connected'
                        return
                    }

                    $rows = @($entry['Rows'])
                    if ($rows.Count -eq 0) {
                        Write-VerboseLog 'Managed App Assignments: no results.' -Level Warning
                        $script:Window.FindName('TxtDANoResults').Visibility = 'Visible'
                        Show-Notification 'No managed app assignments found.' -BgColor '#FFF3CD' -FgColor '#7A4800'
                        Update-StatusBar -Text 'Connected'
                        return
                    }

                    $dg = $script:Window.FindName('DgDAResults')
                    $dg.Columns.Clear()
                    foreach ($prop in ($rows[0].PSObject.Properties)) {
                        $col = [System.Windows.Controls.DataGridTextColumn]::new()
                        $col.Header = $prop.Name
                        $col.Binding = [System.Windows.Data.Binding]::new($prop.Name)
                        if ($prop.Name -in @('AppDescription','Description')) {
                            $col.MaxWidth = 250
                            $col.Width = [System.Windows.Controls.DataGridLength]::new(200)
                            $style = [System.Windows.Style]::new([System.Windows.Controls.TextBlock])
                            $style.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.TextBlock]::TextWrappingProperty, [System.Windows.TextWrapping]::NoWrap))
                            $style.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.TextBlock]::TextTrimmingProperty, [System.Windows.TextTrimming]::CharacterEllipsis))
                            $col.ElementStyle = $style
                        }
                        $dg.Columns.Add($col)
                    }
                    $dg.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSCustomObject]]::new($rows)
                    $script:Window.FindName('TxtDACount').Text            = "$($rows.Count) managed app assignment(s) found"
                    $script:Window.FindName('PnlDAResults').Visibility     = 'Visible'
                    $script:Window.FindName('BtnDACopyAll').IsEnabled      = $true
                    $script:Window.FindName('TxtDAFilter').IsEnabled  = $true
                    $script:Window.FindName('BtnDAFilter').IsEnabled  = $true
                    $script:Window.FindName('BtnDAFilterClear').IsEnabled = $true
                    $script:Window.FindName('BtnDAExportXlsx').IsEnabled   = $true
                    Write-VerboseLog "Managed App Assignments: $($rows.Count) record(s) loaded." -Level Success
                    Show-Notification "$($rows.Count) managed app assignment(s) ready." -BgColor '#D4EDDA' -FgColor '#155724'
                    Update-StatusBar -Text 'Connected'
                }
            }
        }
    })
    $script:DaTimer.Start()
})
# ── DgDAResults: cell selection ──
$script:Window.FindName('DgDAResults').Add_SelectedCellsChanged({
    $dg       = $script:Window.FindName('DgDAResults')
    $selCells = $dg.SelectedCells.Count
    $script:Window.FindName('BtnDACopyValue').IsEnabled = ($selCells -ge 1)
    $script:Window.FindName('BtnDACopyRow').IsEnabled   = ($selCells -gt 0)
})

# ── BtnDACopyValue ──
# ── BtnDAFilter ──
$script:Window.FindName('BtnDAFilter').Add_Click({
    $dg      = $script:Window.FindName('DgDAResults')
    $keyword = $script:Window.FindName('TxtDAFilter').Text.Trim()
    if ($null -eq $script:DAAllData) {
        $script:DAAllData = @($dg.ItemsSource)
    }
    if ($keyword -eq '') {
        $dg.ItemsSource = $script:DAAllData
        Show-Notification "Filter cleared - $($script:DAAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    } else {
        $filtered = @($script:DAAllData | Where-Object {
            $row = $_
            $match = $false
            foreach ($prop in $row.PSObject.Properties) {
                if ([string]$prop.Value -like "*$keyword*") { $match = $true; break }
            }
            $match
        })

# ── BtnDAFilterClear ──
$script:Window.FindName('BtnDAFilterClear').Add_Click({
    $script:Window.FindName('TxtDAFilter').Text = ''
    $dg = $script:Window.FindName('DgDAResults')
    if ($null -ne $script:DAAllData) {
        $dg.ItemsSource = $script:DAAllData
        Show-Notification "Filter cleared - $($script:DAAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    }
})
        $dg.ItemsSource = $filtered
        Show-Notification "$($filtered.Count) of $($script:DAAllData.Count) result(s) shown after filter." -BgColor '#D4EDDA' -FgColor '#155724'
    }
    Write-VerboseLog "DA: Filter applied - keyword='$keyword'" -Level Info
})

$script:Window.FindName('BtnDACopyValue').Add_Click({
    $dg = $script:Window.FindName('DgDAResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $vals = [System.Collections.Generic.List[string]]::new()
    foreach ($cellInfo in @($dg.SelectedCells)) {
        $col = $cellInfo.Column -as [System.Windows.Controls.DataGridBoundColumn]
        if ($col) { $vals.Add(([string]$cellInfo.Item.($col.Binding.Path.Path)).Trim()) }
    }
    [System.Windows.Clipboard]::SetText($vals -join "`r`n")
    Show-Notification '$($vals.Count) value(s) copied to clipboard.' -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog 'DA: Copied cell value to clipboard.' -Level Success
})

# ── BtnDACopyRow ──
$script:Window.FindName('BtnDACopyRow').Add_Click({
    $dg = $script:Window.FindName('DgDAResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $seen = [System.Collections.Generic.HashSet[object]]::new()
    $rowObjs = [System.Collections.Generic.List[object]]::new()
    foreach ($cell in @($dg.SelectedCells)) { if ($seen.Add($cell.Item)) { $rowObjs.Add($cell.Item) } }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc     = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $txtLines = [System.Collections.Generic.List[string]]::new()
    $txtLines.Add(($colDefs | ForEach-Object { $_.H }) -join "`t")
    foreach ($row in $rowObjs) {
        $vals = $colDefs | ForEach-Object { [string]$row.($_.P) }
        $txtLines.Add($vals -join "`t")
    }
    [System.Windows.Clipboard]::SetText($txtLines -join "`n")
    Show-Notification "$($rowObjs.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "DA: Copied $($rowObjs.Count) row(s) to clipboard." -Level Success
    # Highlight full rows
    $dg.UnselectAllCells()
    foreach ($r in $rowObjs) { $null = $dg.SelectedItems.Add($r) }
})

# ── BtnDACopyAll ──
$script:Window.FindName('BtnDACopyAll').Add_Click({
    $dg       = $script:Window.FindName('DgDAResults')
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) { return }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc     = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $txtLines = [System.Collections.Generic.List[string]]::new()
    $txtLines.Add(($colDefs | ForEach-Object { $_.H }) -join "`t")
    foreach ($row in $allItems) {
        $vals = $colDefs | ForEach-Object { [string]$row.($_.P) }
        $txtLines.Add($vals -join "`t")
    }
    [System.Windows.Clipboard]::SetText($txtLines -join "`n")
    Show-Notification "$($allItems.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "DA: Copied all $($allItems.Count) row(s) to clipboard." -Level Success
})

# ── BtnDAExportXlsx ──
$script:Window.FindName('BtnDAExportXlsx').Add_Click({
    # Determine export context from DataGrid content
    $dg = $script:Window.FindName('DgDAResults')
    $firstItem = ($dg.ItemsSource | Select-Object -First 1)
    $hasManagedCols = ($null -ne $firstItem) -and ($firstItem.PSObject.Properties.Name -contains 'AssignmentIntent')
    $exportPrefix = if ($hasManagedCols) { 'ManagedAppAssignments' } else { 'AppAssignments' }
    $wsName = if ($hasManagedCols) { 'ManagedAssignments' } else { 'AppAssignments' }
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) {
        Show-Notification 'No results to export.' -BgColor '#FFF3CD' -FgColor '#856404'
        return
    }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc     = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title            = 'Export App Results'
    $dlg.Filter           = 'Excel Workbook (*.xlsx)|*.xlsx|CSV File (*.csv)|*.csv|All Files (*.*)|*.*'
    $dlg.FileName         = "${exportPrefix}_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($dlg.ShowDialog() -ne $true) { return }
    $outPath = $dlg.FileName
    try {
        $folder = Split-Path $outPath -Parent
        if ($folder -and -not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        $exportRows = $allItems | ForEach-Object {
            $rowObj = $_; $ht = [ordered]@{}
            foreach ($cd in $colDefs) { $ht[$cd['H']] = [string]$rowObj.($cd['P']) }
            [PSCustomObject]$ht
        }
        if (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue) {
            Import-Module ImportExcel -ErrorAction SilentlyContinue
            $exportRows | Export-Excel -Path $outPath -WorksheetName $wsName -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium6
            Show-Notification "Exported $($allItems.Count) rows: $outPath" -BgColor '#D4EDDA' -FgColor '#155724'
        } else {
            $exportRows | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8
            Show-Notification "ImportExcel not found - saved as CSV: $(Split-Path $outPath -Leaf)" -BgColor '#FFF3CD' -FgColor '#856404'
        }
        Write-VerboseLog "DA: Exported $($allItems.Count) row(s) to: $outPath" -Level Success
    } catch {
        Write-VerboseLog "DA Export failed: $($_.Exception.Message)" -Level Error
        Show-Notification "Export failed: $($_.Exception.Message)" -BgColor '#F8D7DA' -FgColor '#721C24'
    }
})

$script:Window.FindName('BtnOpCompareGroups').Add_Click({
    $script:Window.FindName('PnlCGResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtCGNoResults').Visibility = 'Collapsed'
    $script:Window.FindName('PnlCGProgress').Visibility  = 'Collapsed'
    Show-Panel 'PanelCompareGroups'
    Hide-Notification
    Write-VerboseLog 'Panel: Compare Groups' -Level Info
})

# ── GET DEVICE INFO  -  Run ────────────────────────────────────────────
$script:Window.FindName('BtnGDIRun').Add_Click({
    Hide-Notification

    $inputTxt   = $script:Window.FindName('TxtGDIInputList').Text
    $inputTypeSel = $script:Window.FindName('CmbGDIInputType').SelectedItem.Content

    if ([string]::IsNullOrWhiteSpace($inputTxt)) {
        Show-Notification 'Please enter at least one identifier.' -BgColor '#F8D7DA' -FgColor '#721C24'
        return
    }

    # Collect filters on UI thread
    $platFilter = [System.Collections.Generic.List[string]]::new()
    if ($script:Window.FindName('ChkGDIWindows').IsChecked) { $platFilter.Add('Windows') }
    if ($script:Window.FindName('ChkGDIAndroid').IsChecked) { $platFilter.Add('Android') }
    if ($script:Window.FindName('ChkGDIiOS').IsChecked)     { $platFilter.Add('iOS') }
    if ($script:Window.FindName('ChkGDIMacOS').IsChecked)   { $platFilter.Add('macOS') }

    $ownershipSel = $script:Window.FindName('CmbGDIOwnership').SelectedItem.Content
    $ownershipFilter = switch ($ownershipSel) {
        'Company only'  { 'Company' }
        'Personal only' { 'Personal' }
        default         { 'All' }
    }

    $script:GDIParams = @{
        InputTxt        = $inputTxt
        InputType       = $inputTypeSel
        PlatformFilter  = $platFilter
        OwnershipFilter = $ownershipFilter
    }

    # Reset table state
    $script:Window.FindName('PnlGDIResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtGDINoResults').Visibility = 'Collapsed'
    $script:Window.FindName('BtnGDICopyValue').IsEnabled  = $false
    $script:Window.FindName('BtnGDICopyRow').IsEnabled    = $false
    $script:Window.FindName('BtnGDICopyAll').IsEnabled    = $false
    $script:Window.FindName('TxtGDIFilter').IsEnabled  = $false
    $script:Window.FindName('BtnGDIFilter').IsEnabled  = $false
    $script:Window.FindName('BtnGDIFilterClear').IsEnabled = $false
    $script:Window.FindName('TxtGDIFilter').Text       = ''
    $script:GDIAllData = $null
    $script:Window.FindName('BtnGDIExportXlsx').IsEnabled = $false
    $dg = $script:Window.FindName('DgGDIResults')
    if ($dg) { $dg.ItemsSource = $null }
    $script:Window.FindName('PnlGDIProgress').Visibility      = 'Visible'
    $script:Window.FindName('TxtGDIProgressMsg').Text         = 'Resolving devices...'
    $script:Window.FindName('TxtGDIProgressDetail').Text      = ''

    $script:Window.FindName('BtnGDIRun').IsEnabled = $false
    $script:Window.FindName('BtnGDIRunAll').IsEnabled = $false
    $script:Window.FindName('BtnStop').Visibility = 'Visible'
    $script:Window.FindName('BtnStop').IsEnabled  = $true
    $script:Window.FindName('BtnStop').Content = (New-StopBtnContent 'Stop')
    $script:BgStopped = $false
    Write-VerboseLog '--- Get Device Info: starting ---' -Level Action

    # ── Dedicated runspace  -  all state in script scope so DispatcherTimer
    #    tick can access it reliably (local closure capture is unreliable in WPF handlers)
    $script:GdiQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $script:GdiStop  = [hashtable]::Synchronized(@{ Stop = $false })

    $script:GdiRs = [runspacefactory]::CreateRunspace()
    $script:GdiRs.ApartmentState = 'STA'
    $script:GdiRs.ThreadOptions  = 'ReuseThread'
    $script:GdiRs.Open()
    $script:GdiRs.SessionStateProxy.SetVariable('GdiQueue',  $script:GdiQueue)
    $script:GdiRs.SessionStateProxy.SetVariable('GdiStop',   $script:GdiStop)
    $script:GdiRs.SessionStateProxy.SetVariable('GdiParams', $script:GDIParams)
    $script:GdiRs.SessionStateProxy.SetVariable('LogFile',   $script:LogFile)

    $script:GdiPs = [powershell]::Create()
    $script:GdiPs.Runspace = $script:GdiRs

    $null = $script:GdiPs.AddScript({
        # ── Helper: enqueue a log message ──
        function GLog {
            param([string]$Msg, [string]$Level = 'Info')
            $prefix = switch ($Level) {
                "Success" { "$([char]0x2713) " } "Warning" { "[!] " } "Error" { "$([char]0x2717) " }
                'Action'  { '> ' } default   { '  ' }
            }
            $line = "[$(Get-Date -Format 'HH:mm:ss')] $prefix$Msg"
            $GdiQueue.Enqueue(@{ Type='log'; Line=$line; Level=$Level })
            if ($LogFile) {
                try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
            }
        }

        # ── OS normalisation (Entra raw → friendly) ──
        function Normalize-OS([string]$os) {
            switch -Wildcard ($os) {
                'Windows*' { 'Windows' } 'Android*' { 'Android' }
                'iPhone'   { 'iOS'     } 'iPadOS'   { 'iOS'     }
                'iOS'      { 'iOS'     } 'MacMDM'   { 'macOS'   }
                'Mac OS X' { 'macOS'   } 'macOS'    { 'macOS'   }
                default    { $os }
            }
        }

        # ── Windows release lookup ──
        $WinVerMap = @(
            @{ V='10.0.26200'; R='25H2'          }
            @{ V='10.0.28000'; R='26H1'           }
            @{ V='10.0.26100'; R='24H2'           }
            @{ V='10.0.22631'; R='23H2'           }
            @{ V='10.0.22621'; R='22H2 (Win11)'   }
            @{ V='10.0.19045'; R='22H2 (Win10)'   }
            @{ V='10.0.19044'; R='21H2 (LTSC)'    }
            @{ V='10.0.17763'; R='1809 (LTSC)'    }
        )
        function Get-WinRelease([string]$ver) {
            if ([string]::IsNullOrWhiteSpace($ver)) { return $null }
            $v = $ver.Trim()
            $m = $WinVerMap | Where-Object { $_.V -eq $v } | Select-Object -First 1
            if ($m) { return $m.R }
            # Try major.minor.build prefix only
            if ($v -match '^\d+\.\d+\.\d+') {
                $pfx = ($Matches[0])
                $m = $WinVerMap | Where-Object { $_.V -eq $pfx } | Select-Object -First 1
                if ($m) { return $m.R }
            }
            return $null
        }

        # ── Device type from name suffix ──
        $DevTypeMap = @(
            @{ S='-D'; T='Desktop'            }
            @{ S='-L'; T='Laptop'             }
            @{ S='-V'; T='Virtual'            }
            @{ S='-K'; T='Kiosk'              }
            @{ S='-F'; T='Teams Meeting Room' }
            @{ S='-A'; T='Shared Device'      }
        )
        function Get-DevType([string]$name) {
            if ([string]::IsNullOrWhiteSpace($name)) { return $null }
            $u = $name.Trim().ToUpperInvariant()
            foreach ($m in $DevTypeMap) {
                if ($u.EndsWith($m.S.ToUpperInvariant())) { return $m.T }
            }
            return $null
        }

        # ── Activity range label ──
        # Tries RoundtripKind first (ISO 8601/UTC), then InvariantCulture general parse
        # as fallback  -  same two-style approach as Format-Date to handle all Graph date formats.
        function Get-ActivityRange([string]$dateStr) {
            if ([string]::IsNullOrWhiteSpace($dateStr)) { return $null }
            $d = $null
            foreach ($style in @(
                [System.Globalization.DateTimeStyles]::RoundtripKind,
                [System.Globalization.DateTimeStyles]::None
            )) {
                try {
                    $d = [datetime]::Parse($dateStr,
                             [System.Globalization.CultureInfo]::InvariantCulture, $style)
                    break
                } catch {}
            }
            if ($null -eq $d) { return $null }
            $days = [int][Math]::Floor(([datetime]::UtcNow - $d.ToUniversalTime()).TotalDays)
            if ($days -le 7)       { return '0-7 days'        }
            elseif ($days -le 14)  { return '8-14 days'       }
            elseif ($days -le 30)  { return '15-30 days'      }
            elseif ($days -le 60)  { return '30-60 days'      }
            elseif ($days -le 90)  { return '60-90 days'      }
            elseif ($days -le 180) { return '90-180 days'     }
            elseif ($days -le 365) { return '180-365 days'    }
            else                   { return 'Older than 1 yr' }
        }

        # ── Page through a Graph URI ──
        function Get-GraphAll([string]$Uri, [hashtable]$Headers = @{}) {
            $items = [System.Collections.Generic.List[object]]::new()
            $next  = $Uri
            do {
                $r = Invoke-MgGraphRequest -Method GET -Uri $next -Headers $Headers
                if ($r['value']) { foreach ($i in @($r['value'])) { $items.Add($i) } }
                $next = if ($r.ContainsKey('@odata.nextLink')) { $r['@odata.nextLink'] } else { $null }
            } while ($next)
            return ,$items
        }

        # ── Fetch full Entra device record ──
        function Get-EntraDevice([string]$Uri) {
            try {
                return Invoke-MgGraphRequest -Method GET -Uri $Uri
            } catch { return $null }
        }

        # ── Fetch Intune device by azureADDeviceId ──
        function Get-IntuneByEntraId([string]$EntraId, [hashtable]$Cache) {
            if ([string]::IsNullOrWhiteSpace($EntraId)) { return $null }
            if ($Cache.ContainsKey($EntraId)) { return $Cache[$EntraId] }
            try {
                $enc  = [Uri]::EscapeDataString($EntraId)
                $resp = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=azureADDeviceId eq '$enc'`&`$select=id,deviceName,operatingSystem,osVersion,lastSyncDateTime,manufacturer,model,serialNumber,managedDeviceOwnerType,enrolledDateTime,azureADDeviceId,userPrincipalName,complianceState`&`$top=1"
                $dev = if ($resp['value'] -and $resp['value'].Count -gt 0) { $resp['value'][0] } else { $null }
                $Cache[$EntraId] = $dev
                return $dev
            } catch { $Cache[$EntraId] = $null; return $null }
        }

        # ── Fetch user info ──
        function Get-UserInfo([string]$UserId) {
            try {
                return Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/users/$([Uri]::EscapeDataString($UserId))?`$select=id,userPrincipalName,displayName,jobTitle,accountEnabled,city,country,officeLocation"
            } catch { return $null }
        }

        # ── Format any date string to MM-dd-yyyy, tolerating multiple input formats ──
        function Format-Date([string]$raw) {
            if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
            # Try ISO 8601 / RoundtripKind first, then general parse as fallback
            foreach ($style in @(
                [System.Globalization.DateTimeStyles]::RoundtripKind,
                [System.Globalization.DateTimeStyles]::None
            )) {
                try {
                    return ([datetime]::Parse($raw, [System.Globalization.CultureInfo]::InvariantCulture, $style)).ToString('MM-dd-yyyy')
                } catch {}
            }
            return $raw   # return raw if all parsing fails
        }

        # ── Build Autopilot lookup dictionary ──────────────────────────────────────
        # deploymentProfileAssignmentStatus is a NESTED OBJECT containing deploymentProfileId.
        # Profile name is resolved by: $ap['deploymentProfileAssignmentStatus']['deploymentProfileId']
        # → fetch profile display name from windowsAutopilotDeploymentProfiles/{id}.
        # Step 1: load all profiles into id->name map.
        # Step 2: load all identities, extract profileId from nested status object.
        function Build-AutopilotDictionary {
            $dict       = @{}
            $profileMap = @{}

            # ── Step 1: all deployment profiles ──
            try {
                $nextUri = 'https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles?$select=id,displayName&$top=100'
                do {
                    $r = Invoke-MgGraphRequest -Method GET -Uri $nextUri
                    foreach ($p in @($r['value'])) { $profileMap[[string]$p['id']] = [string]$p['displayName'] }
                    $nextUri = if ($r.ContainsKey('@odata.nextLink')) { $r['@odata.nextLink'] } else { $null }
                } while ($nextUri)
                GLog "Autopilot: $($profileMap.Count) profile(s) loaded" 'Info'
            } catch {
                GLog "Autopilot profiles failed: $($_.Exception.Message)" 'Warning'
            }

            # ── Step 2: all device identities ──
            try {
                GLog 'Autopilot: loading device identities...' 'Info'
                $nextUri = 'https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities?$top=500'
                $total   = 0
                do {
                    $resp  = Invoke-MgGraphRequest -Method GET -Uri $nextUri
                    $page  = @($resp['value'])
                    $total += $page.Count
                    foreach ($ap in $page) {
                        $aadId = [string]$ap['azureActiveDirectoryDeviceId']
                        if ([string]::IsNullOrWhiteSpace($aadId)) { continue }

                        # deploymentProfileAssignmentStatus is a nested object  - 
                        # the profileId lives inside it, not as a top-level field
                        $statusObj  = $ap['deploymentProfileAssignmentStatus']
                        $statusStr  = ''
                        $profileId  = $null
                        if ($statusObj -is [System.Collections.IDictionary]) {
                            $profileId = [string]$statusObj['deploymentProfileId']
                            $statusStr = [string]$statusObj['value']   # e.g. 'assigned'
                            if (-not $statusStr) { $statusStr = [string]$statusObj['status'] }
                        } else {
                            $statusStr = [string]$statusObj   # simple string fallback
                        }

                        # Resolve profile name from the profiles map
                        $profileName = if ($profileId -and $profileMap.ContainsKey($profileId)) {
                                           $profileMap[$profileId]
                                       } else { $null }

                        $dict[$aadId.ToLower()] = @{
                            GroupTag      = [string]$ap['groupTag']
                            ProfileStatus = $statusStr
                        }
                    }
                    $nextUri = if ($resp.ContainsKey('@odata.nextLink')) { $resp['@odata.nextLink'] } else { $null }
                } while ($nextUri)
                GLog "Autopilot: $total record(s), $($dict.Count) indexed by Entra Device ID" 'Success'
            } catch {
                GLog "Autopilot identities failed: $($_.Exception.Message)" 'Warning'
            }
            return ,$dict
        }

        # ── Build a result row from Entra + Intune data ──
        function Build-Row([hashtable]$EntraDev, [hashtable]$IntuneDev, [hashtable]$UserInfo, [string]$OwnerUPN = '', [hashtable]$AutopilotInfo = $null) {
            $NoRec = 'No Intune Record'
            $devName = [string]$EntraDev['displayName']

            # Entra fields
            $entraDevId  = [string]$EntraDev['deviceId']    # the GUID used as azureADDeviceId
            $entraObjId  = [string]$EntraDev['id']
            $entraOS     = [string]$EntraDev['operatingSystem']
            $entraOSVer  = [string]$EntraDev['operatingSystemVersion']
            $entraJoin   = switch ([string]$EntraDev['trustType']) {
                'AzureAd'    { 'Entra Joined'  }
                'ServerAd'   { 'Hybrid Joined' }
                'Workplace'  { 'Registered'    }
                default      { [string]$EntraDev['trustType'] }
            }
            $entraOwner  = [string]$EntraDev['deviceOwnership']
            $entraEnabled= [string]$EntraDev['accountEnabled']
            $entraSignIn = [string]$EntraDev['approximateLastSignInDateTime']
            $entraSignInFmt = Format-Date $entraSignIn
            $entraActivity = Get-ActivityRange $entraSignIn

            # Intune fields
            if ($IntuneDev) {
                $intuneId       = [string]$IntuneDev['id']
                $intuneOS       = [string]$IntuneDev['operatingSystem']
                $intuneOSVer    = [string]$IntuneDev['osVersion']
                $intuneSync     = [string]$IntuneDev['lastSyncDateTime']
                $intuneSyncFmt  = if ($intuneSync) { Format-Date $intuneSync } else { $NoRec }
                $intuneActivity = Get-ActivityRange $intuneSync
                $intuneOEM      = [string]$IntuneDev['manufacturer']
                $intuneModel    = [string]$IntuneDev['model']
                $intuneSerial   = [string]$IntuneDev['serialNumber']
                $intuneOwner    = [string]$IntuneDev['managedDeviceOwnerType']
                $intuneEnrolled = Format-Date ([string]$IntuneDev['enrolledDateTime'])
                $intuneUPN      = [string]$IntuneDev['userPrincipalName']
                $intuneCompliance = [string]$IntuneDev['complianceState']
            } else {
                $intuneId       = $NoRec
                $intuneOS       = $NoRec
                $intuneOSVer    = $NoRec
                $intuneSyncFmt  = $NoRec
                $intuneActivity = $NoRec
                $intuneOEM      = $NoRec
                $intuneModel    = $NoRec
                $intuneSerial   = $NoRec
                $intuneOwner    = $NoRec
                $intuneEnrolled = $NoRec
                $intuneUPN      = $NoRec
                $intuneCompliance = $NoRec
            }

            # Autopilot fields (from dedicated lookup  -  $null if not Autopilot registered)
            $apGroupTag  = if ($AutopilotInfo) { [string]$AutopilotInfo['GroupTag']     } else { $NoRec }
            $apStatus    = if ($AutopilotInfo) { [string]$AutopilotInfo['ProfileStatus'] } else { $NoRec }

            # Prefer Intune OS version for Windows release mapping
            $osVerForMap = if ($intuneOSVer -and $intuneOSVer -ne $NoRec) { $intuneOSVer } else { $entraOSVer }
            $winRelease  = Get-WinRelease $osVerForMap
            $devType     = Get-DevType $devName

            # User fields  -  UserInfo from Entra; fall back to OwnerUPN (Entra) or Intune UPN
            $resolvedUPN = if ($UserInfo) { [string]$UserInfo['userPrincipalName'] }
                           elseif ($OwnerUPN) { $OwnerUPN }
                           elseif ($intuneUPN -and $intuneUPN -ne $NoRec) { $intuneUPN }
                           else { $null }
            $uTitle   = if ($UserInfo) { [string]$UserInfo['jobTitle'] }       else { $null }
            $uEnabled = if ($UserInfo) { [string]$UserInfo['accountEnabled'] } else { $null }
            $uCity    = if ($UserInfo) { [string]$UserInfo['city'] }           else { $null }
            $uCountry = if ($UserInfo) { [string]$UserInfo['country'] }        else { $null }

            return [PSCustomObject]@{
                'Entra_DeviceName'            = $devName
                'DeviceType'                  = $devType
                'Entra_DeviceID'              = $entraDevId
                'Entra_ObjectID'              = $entraObjId
                'Intune_DeviceID'             = $intuneId
                'Entra_JoinType'              = $entraJoin
                'Entra_Ownership'             = $entraOwner
                'Intune_Ownership'            = $intuneOwner
                'Entra_OSPlatform'            = $entraOS
                'Intune_OSPlatform'           = $intuneOS
                'Entra_OSVersion'             = $entraOSVer
                'Intune_OSVersion'            = $intuneOSVer
                'Windows_Release'             = $winRelease
                'Entra_DeviceEnabled'         = $entraEnabled
                'Entra_LastSignIn'            = $entraSignInFmt
                'Entra_ActivityRange'         = $entraActivity
                'Intune_LastCheckIn'          = $intuneSyncFmt
                'Intune_ActivityRange'        = $intuneActivity
                'Intune_OEM'                  = $intuneOEM
                'Intune_Model'                = $intuneModel
                'Intune_SerialNumber'         = $intuneSerial
                'Intune_EnrolledDate'         = $intuneEnrolled
                'Intune_ComplianceState'      = $intuneCompliance
                'Entra_DeviceOwner'           = $OwnerUPN
                'Entra_PrimaryUserUPN'        = $resolvedUPN
                'Entra_UserAccountEnabled'    = $uEnabled
                'Entra_UserCountry'           = $uCountry
                'Entra_UserCity'              = $uCity
                'Entra_UserJobTitle'          = $uTitle
                'Autopilot_GroupTag'          = $apGroupTag
                'Autopilot_ProfileStatus'     = $apStatus
            }
        }

        # ══════════════════════════════════════════════════════════════════
        # MAIN WORK
        # ══════════════════════════════════════════════════════════════════
        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue

            $params         = $GdiParams
            $platFilter     = @($params['PlatformFilter'])
            $ownerFilter    = $params['OwnershipFilter']
            $allDevicesMode = $params['AllDevices'] -eq $true

            $intuneCache    = @{}  # azureADDeviceId -> Intune device hashtable
            $userCache      = @{}  # userId / UPN -> user hashtable
            # Build Autopilot lookup once upfront  -  keyed by azureActiveDirectoryDeviceId (lowercase)
            $autopilotDict  = Build-AutopilotDictionary
            $entraIdSet  = [System.Collections.Generic.HashSet[string]]::new()  # deduplicate Entra Object IDs

            $deviceObjIds = [System.Collections.Generic.List[string]]::new()
            $isGuidRx = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'

            if ($allDevicesMode) {
                # ── All Devices mode: fetch every Entra device object ──
                GLog 'All Devices mode  -  fetching all Entra device objects...' 'Action'
                GLog "Platform filter : $(if ($platFilter.Count -eq 0) { 'All' } else { $platFilter -join ', ' })" 'Info'
                GLog "Ownership filter: $ownerFilter" 'Info'

                $allEntra = Get-GraphAll 'https://graph.microsoft.com/v1.0/devices?$select=id&$top=999'
                GLog "Retrieved $($allEntra.Count) Entra device object(s)" 'Info'
                foreach ($d in $allEntra) {
                    $null = $deviceObjIds.Add([string]$d['id'])
                }
            } else {
                # ── Input-based mode: resolve entries to Entra device Object IDs ──
                $inputType      = $params['InputType']
                $entries        = $params['InputTxt'] -split "`n" |
                                  ForEach-Object { $_.Trim() } |
                                  Where-Object   { $_ -ne '' }

                GLog "Input type      : $inputType" 'Info'
                GLog "Entries         : $($entries.Count)" 'Info'
                GLog "Platform filter : $(if ($platFilter.Count -eq 0) { 'All' } else { $platFilter -join ', ' })" 'Info'
                GLog "Ownership filter: $ownerFilter" 'Info'

                foreach ($entry in $entries) {
                    if ($GdiStop['Stop']) { GLog 'Stop requested.' 'Warning'; break }

                    switch -Wildcard ($inputType) {

                    'User UPNs*' {
                        GLog "Resolving user devices: $entry" 'Info'
                        try {
                            $uid = [Uri]::EscapeDataString($entry)
                            $devs = Get-GraphAll "https://graph.microsoft.com/v1.0/users/$uid/ownedDevices?`$select=id,displayName,deviceId,operatingSystem,operatingSystemVersion,deviceOwnership,trustType,accountEnabled,approximateLastSignInDateTime`&`$top=999"
                            GLog "  $($devs.Count) owned device(s) found" 'Info'
                            foreach ($d in $devs) { $null = $deviceObjIds.Add([string]$d['id']) }
                        } catch {
                            GLog "  Could not resolve user $entry : $($_.Exception.Message)" 'Warning'
                        }
                    }

                    'Device names' {
                        GLog "Resolving device by name: $entry" 'Info'
                        try {
                            $safe = $entry -replace "'","''"
                            $r = Invoke-MgGraphRequest -Method GET `
                                -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=displayName eq '$safe'`&`$select=id`&`$top=5"
                            $found = @($r['value'])
                            if ($found.Count -eq 0) { GLog "  Not found: $entry" 'Warning' }
                            else { foreach ($d in $found) { $null = $deviceObjIds.Add([string]$d['id']) } }
                        } catch {
                            GLog "  Error resolving $entry : $($_.Exception.Message)" 'Warning'
                        }
                    }

                    'Device serial numbers' {
                        GLog "Resolving serial number: $entry" 'Info'
                        # Serial numbers only exist in Intune  -  search there first, get azureADDeviceId, then Entra
                        try {
                            $safe = $entry -replace "'","''"
                            $r = Invoke-MgGraphRequest -Method GET `
                                -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=serialNumber eq '$safe'`&`$select=id,azureADDeviceId,deviceName`&`$top=5"
                            $found = @($r['value'])
                            if ($found.Count -eq 0) { GLog "  Serial not found in Intune: $entry" 'Warning' }
                            else {
                                foreach ($md in $found) {
                                    $aadId = [string]$md['azureADDeviceId']
                                    if ($aadId) {
                                        # Find Entra Object ID from azureADDeviceId (deviceId field)
                                        $er = Invoke-MgGraphRequest -Method GET `
                                            -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$aadId'`&`$select=id`&`$top=1"
                                        if ($er['value'] -and $er['value'].Count -gt 0) {
                                            $null = $deviceObjIds.Add([string]$er['value'][0]['id'])
                                        }
                                    }
                                }
                            }
                        } catch {
                            GLog "  Error resolving serial $entry : $($_.Exception.Message)" 'Warning'
                        }
                    }

                    'Entra Device Object IDs' {
                        if ($entry -match $isGuidRx) {
                            $null = $deviceObjIds.Add($entry)
                        } else {
                            GLog "  Skipped (not a GUID): $entry" 'Warning'
                        }
                    }

                    'Intune Device IDs' {
                        GLog "Resolving Intune device ID: $entry" 'Info'
                        try {
                            $md = Invoke-MgGraphRequest -Method GET `
                                -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$entry`?`$select=id,azureADDeviceId,deviceName"
                            $aadId = [string]$md['azureADDeviceId']
                            if ($aadId) {
                                $er = Invoke-MgGraphRequest -Method GET `
                                    -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$aadId'`&`$select=id`&`$top=1"
                                if ($er['value'] -and $er['value'].Count -gt 0) {
                                    $null = $deviceObjIds.Add([string]$er['value'][0]['id'])
                                }
                            }
                        } catch {
                            GLog "  Could not resolve Intune ID $entry : $($_.Exception.Message)" 'Warning'
                        }
                    }

                    'Groups*' {
                        GLog "Resolving group: $entry" 'Info'
                        # Resolve group ID
                        $groupId = $null
                        if ($entry -match $isGuidRx) {
                            $groupId = $entry
                        } else {
                            try {
                                $safe = $entry -replace "'","''"
                                $gr = Invoke-MgGraphRequest -Method GET `
                                    -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe'`&`$select=id`&`$top=1"
                                if ($gr['value'] -and $gr['value'].Count -gt 0) {
                                    $groupId = [string]$gr['value'][0]['id']
                                }
                            } catch {}
                        }
                        if (-not $groupId) { GLog "  Group not found: $entry" 'Warning'; continue }

                        # Get all transitive members (users + devices)
                        GLog "  Scanning group members (transitive)..." 'Info'
                        # $select with @odata.type or user-subtype properties causes 400 on
                        # transitiveMembers (directoryObject base type). Use $select=id only;
                        # @odata.type is always returned automatically.
                        $members = Get-GraphAll "https://graph.microsoft.com/v1.0/groups/$groupId/transitiveMembers?`$select=id`&`$top=999"
                        GLog "  $($members.Count) member(s) found" 'Info'

                        foreach ($mem in $members) {
                            if ($GdiStop['Stop']) { break }
                            $odataType = [string]$mem['@odata.type']
                            if ($odataType -eq '#microsoft.graph.device') {
                                $null = $deviceObjIds.Add([string]$mem['id'])
                            } elseif ($odataType -eq '#microsoft.graph.user') {
                                # Resolve user's owned devices
                                $uid = [string]$mem['id']
                                try {
                                    $udevs = Get-GraphAll "https://graph.microsoft.com/v1.0/users/$uid/ownedDevices?`$select=id`&`$top=999"
                                    foreach ($ud in $udevs) { $null = $deviceObjIds.Add([string]$ud['id']) }
                                } catch {}
                            }
                        }
                    }
                }
            }

                        }  # end else (input-based mode)

            GLog "Total device IDs to process: $($deviceObjIds.Count)" 'Info'
            if ($deviceObjIds.Count -eq 0) {
                .Enqueue(@{ Type='done'; Rows=@(); Error='No devices resolved from the provided input.' })
                return
            }

            # ── Process each Entra device ──
            $results = [System.Collections.Generic.List[PSCustomObject]]::new()
            $processed = 0

            foreach ($objId in $deviceObjIds) {
                if ($GdiStop['Stop']) { GLog 'Stop requested  -  halting.' 'Warning'; break }
                if ($entraIdSet.Contains($objId)) { continue }   # deduplicate
                $null = $entraIdSet.Add($objId)

                $processed++
                $GdiQueue.Enqueue(@{ Type='progress'; Current=$processed; Total=$deviceObjIds.Count })

                # Fetch full Entra device record
                $entraRaw = $null
                try {
                    $entraRaw = Invoke-MgGraphRequest -Method GET `
                        -Uri "https://graph.microsoft.com/v1.0/devices/$objId`?`$select=id,deviceId,displayName,operatingSystem,operatingSystemVersion,deviceOwnership,trustType,accountEnabled,approximateLastSignInDateTime,registrationDateTime"
                } catch {
                    GLog "  Could not fetch Entra device $objId : $($_.Exception.Message)" 'Warning'
                    continue
                }

                $devName   = [string]$entraRaw['displayName']
                $entraDevId = [string]$entraRaw['deviceId']   # azureADDeviceId in Intune
                $entraOS   = Normalize-OS ([string]$entraRaw['operatingSystem'])

                GLog "[$processed] $devName  [$entraOS]" 'Info'

                # ── Platform filter ──
                if ($platFilter.Count -gt 0 -and $entraOS -notin $platFilter) {
                    GLog "  Skipped (platform $entraOS not in filter)" 'Info'
                    continue
                }

                # ── Ownership filter ──
                $rawOwner = [string]$entraRaw['deviceOwnership']
                if ($ownerFilter -ne 'All' -and $rawOwner -ne $ownerFilter) {
                    GLog "  Skipped (ownership $rawOwner not in filter)" 'Info'
                    continue
                }

                # ── Intune lookup ──
                $intuneDev = Get-IntuneByEntraId -EntraId $entraDevId -Cache $intuneCache
                if ($intuneDev) {
                    GLog "  Intune: $([string]$intuneDev['id'])" 'Success'
                } else {
                    GLog "  Intune: No Intune Record" 'Warning'
                }

                # ── Registered owner (user) ──
                # No $select on registeredOwners  -  $select is not supported on navigation
                # endpoints and causes userPrincipalName to be missing from the response.
                $userInfo    = $null
                $ownerUPN    = $null
                try {
                    $ownerResp = Invoke-MgGraphRequest -Method GET `
                        -Uri "https://graph.microsoft.com/v1.0/devices/$objId/registeredOwners?`$top=5"
                    $ownerList = @($ownerResp['value'])
                    $firstUser = $ownerList | Where-Object {
                        $t = [string]$_['@odata.type']
                        $t -eq '#microsoft.graph.user' -or $t -eq 'microsoft.graph.user'
                    } | Select-Object -First 1
                    if ($firstUser) {
                        # userPrincipalName lives in AdditionalProperties for some SDK versions,
                        # or directly on the hashtable when using Invoke-MgGraphRequest
                        $ownerUPN = if ($firstUser.ContainsKey('userPrincipalName')) {
                                        [string]$firstUser['userPrincipalName']
                                    } elseif ($firstUser.ContainsKey('AdditionalProperties')) {
                                        [string]$firstUser['AdditionalProperties']['userPrincipalName']
                                    } else { $null }
                        $uid = [string]$firstUser['id']
                        # If UPN still empty, fetch it from the user endpoint directly
                        if ([string]::IsNullOrWhiteSpace($ownerUPN) -and $uid) {
                            try {
                                $upnResp  = Invoke-MgGraphRequest -Method GET `
                                    -Uri "https://graph.microsoft.com/v1.0/users/$uid`?`$select=userPrincipalName"
                                $ownerUPN = [string]$upnResp['userPrincipalName']
                            } catch {}
                        }
                        GLog "  Owner: $ownerUPN" 'Info'
                        if ($userCache.ContainsKey($uid)) {
                            $userInfo = $userCache[$uid]
                        } else {
                            $userInfo = Get-UserInfo -UserId $uid
                            $userCache[$uid] = $userInfo
                        }
                    } else {
                        GLog "  No registered owner found" 'Info'
                    }
                } catch {
                    GLog "  Owner fetch error: $($_.Exception.Message)" 'Warning'
                }

                # ── Autopilot lookup  -  dictionary built upfront, keyed by entraDevId lowercase ──
                $autopilotInfo = $null
                $entraOSNorm = Normalize-OS ([string]$entraRaw['operatingSystem'])
                if ($entraOSNorm -eq 'Windows' -and $entraDevId) {
                    $apKey = $entraDevId.ToLower()
                    if ($autopilotDict.ContainsKey($apKey)) {
                        $autopilotInfo = $autopilotDict[$apKey]
                        GLog "  Autopilot: Tag=$($autopilotInfo['GroupTag']) | Status=$($autopilotInfo['ProfileStatus'])" 'Info'
                    } else {
                        GLog '  Autopilot: not registered' 'Info'
                    }
                }

                # ── Build row ──
                $row = Build-Row -EntraDev $entraRaw -IntuneDev $intuneDev -UserInfo $userInfo -OwnerUPN $ownerUPN -AutopilotInfo $autopilotInfo
                $results.Add($row)
            }

            GLog "Processed $($results.Count) device(s)." 'Success'
            $GdiQueue.Enqueue(@{ Type='done'; Rows=$results; Error=$null })

        } catch {
            $GdiQueue.Enqueue(@{ Type='done'; Rows=@(); Error=$_.Exception.Message })
        }
    })

    $script:GdiHandle = $script:GdiPs.BeginInvoke()

    # ── Timer to drain GdiQueue on the UI thread ──
    $script:GdiTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:GdiTimer.Interval = [TimeSpan]::FromMilliseconds(300)
    $script:GdiTimer.Add_Tick({
        $entry = $null
        while ($script:GdiQueue.TryDequeue([ref]$entry)) {
            switch ($entry['Type']) {
                'log' {
                    $lvl = $entry['Level']
                    $col = switch ($lvl) {
                        'Success' { $DevConfig.LogColorSuccess }
                        'Warning' { $DevConfig.LogColorWarning }
                        'Error'   { $DevConfig.LogColorError   }
                        'Action'  { $DevConfig.LogColorAction  }
                        default   { $DevConfig.LogColorInfo    }
                    }
                    try {
                        $para = [System.Windows.Documents.Paragraph]::new()
                        $run  = [System.Windows.Documents.Run]::new($entry['Line'])
                        $run.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                            [System.Windows.Media.ColorConverter]::ConvertFromString($col))
                        $para.Margin = [System.Windows.Thickness]::new(0)
                        $para.Inlines.Add($run)
                        $script:RtbLog.Document.Blocks.Add($para)
                        $script:RtbLog.ScrollToEnd()
                    } catch {}
                }
                'progress' {
                    $cur = $entry['Current']; $tot = $entry['Total']
                    $script:Window.FindName('PnlGDIProgress').Visibility       = 'Visible'
                    $script:Window.FindName('TxtGDIProgressMsg').Text          = 'Fetching device details...'
                    $script:Window.FindName('TxtGDIProgressDetail').Text       = "Processing device $cur of $tot..."
                    Update-StatusBar -Text "Get Device Info: $cur / $tot devices..."
                }
                'done' {
                    $script:GdiTimer.Stop()
                    try { $script:GdiPs.EndInvoke($script:GdiHandle) } catch {}
                    try { $script:GdiRs.Close() }                      catch {}
                    try { $script:GdiPs.Dispose() }                    catch {}
                    $script:GdiPs = $null; $script:GdiRs = $null
                    $script:Window.FindName('BtnStop').Visibility = 'Collapsed'
                    $script:Window.FindName('BtnStop').IsEnabled  = $true
                    $script:Window.FindName('BtnStop').Content = (New-StopBtnContent 'Stop')
                    $script:Window.FindName('BtnGDIRun').IsEnabled = $true
                    $script:Window.FindName('BtnGDIRunAll').IsEnabled = $true
                    $script:Window.FindName('PnlGDIProgress').Visibility = 'Collapsed'

                    $errMsg = $entry['Error']
                    if ($errMsg) {
                        Write-VerboseLog "Get Device Info error: $errMsg" -Level Error
                        Show-Notification "Error: $errMsg" -BgColor '#F8D7DA' -FgColor '#721C24'
                        Update-StatusBar -Text "Connected as: $($script:TxtStatusBar.Text -replace 'Get Device Info.*','')"
                        return
                    }

                    $rows     = @($entry['Rows'])

                    if ($rows.Count -eq 0) {
                        Write-VerboseLog 'Get Device Info: no rows  -  check filters or input.' -Level Warning
                        $script:Window.FindName('TxtGDINoResults').Visibility = 'Visible'
                        Show-Notification 'No devices matched the input and filters.' -BgColor '#FFF3CD' -FgColor '#7A4800'
                        Update-StatusBar -Text 'Connected'
                        return
                    }

                    # ── Populate DataGrid ──
                    $dg = $script:Window.FindName('DgGDIResults')
                    $dg.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSCustomObject]]::new($rows)
                    $script:Window.FindName('TxtGDICount').Text            = "$($rows.Count) device(s) found"
                    $script:Window.FindName('PnlGDIResults').Visibility     = 'Visible'
                    $script:Window.FindName('BtnGDICopyAll').IsEnabled      = $true
                    $script:Window.FindName('TxtGDIFilter').IsEnabled  = $true
                    $script:Window.FindName('BtnGDIFilter').IsEnabled  = $true
                    $script:Window.FindName('BtnGDIFilterClear').IsEnabled = $true
                    $script:Window.FindName('BtnGDIExportXlsx').IsEnabled   = $true
                    Write-VerboseLog "Get Device Info: $($rows.Count) device(s) loaded into table." -Level Success
                    Show-Notification "$($rows.Count) device(s) ready. Use Export XLSX to save." -BgColor '#D4EDDA' -FgColor '#155724'
                    Update-StatusBar -Text 'Connected'
                }
            }
        }
    })
    $script:GdiTimer.Start()
})


# -- BtnGDIRunAll: Get All Device Info --
$script:Window.FindName('BtnGDIRunAll').Add_Click({
    Hide-Notification

    # Collect filters on UI thread (no input list needed)
    $platFilter = [System.Collections.Generic.List[string]]::new()
    if ($script:Window.FindName('ChkGDIWindows').IsChecked) { $platFilter.Add('Windows') }
    if ($script:Window.FindName('ChkGDIAndroid').IsChecked) { $platFilter.Add('Android') }
    if ($script:Window.FindName('ChkGDIiOS').IsChecked)     { $platFilter.Add('iOS') }
    if ($script:Window.FindName('ChkGDIMacOS').IsChecked)   { $platFilter.Add('macOS') }

    $ownershipSel = $script:Window.FindName('CmbGDIOwnership').SelectedItem.Content
    $ownershipFilter = switch ($ownershipSel) {
        'Company only'  { 'Company' }
        'Personal only' { 'Personal' }
        default         { 'All' }
    }

    $script:GDIParams = @{
        AllDevices      = $true
        PlatformFilter  = $platFilter
        OwnershipFilter = $ownershipFilter
    }

    # Reset table state
    $script:Window.FindName('PnlGDIResults').Visibility   = 'Collapsed'
    $script:Window.FindName('TxtGDINoResults').Visibility = 'Collapsed'
    $script:Window.FindName('BtnGDICopyValue').IsEnabled  = $false
    $script:Window.FindName('BtnGDICopyRow').IsEnabled    = $false
    $script:Window.FindName('BtnGDICopyAll').IsEnabled    = $false
    $script:GDIAllData = $null
    $script:Window.FindName('TxtGDIFilter').IsEnabled  = $false
    $script:Window.FindName('BtnGDIFilter').IsEnabled  = $false
    $script:Window.FindName('BtnGDIFilterClear').IsEnabled = $false
    $script:Window.FindName('TxtGDIFilter').Text       = ''
    $script:Window.FindName('BtnGDIExportXlsx').IsEnabled = $false
    $dg = $script:Window.FindName('DgGDIResults')
    if ($dg) { $dg.ItemsSource = $null }
    $script:Window.FindName('PnlGDIProgress').Visibility      = 'Visible'
    $script:Window.FindName('TxtGDIProgressMsg').Text         = 'Fetching all Entra devices...'
    $script:Window.FindName('TxtGDIProgressDetail').Text      = ''

    $script:Window.FindName('BtnGDIRun').IsEnabled    = $false
    $script:Window.FindName('BtnGDIRunAll').IsEnabled = $false
    $script:Window.FindName('BtnStop').Visibility = 'Visible'
    $script:Window.FindName('BtnStop').IsEnabled  = $true
    $script:Window.FindName('BtnStop').Content = (New-StopBtnContent 'Stop')
    $script:BgStopped = $false
    Write-VerboseLog '--- Get All Device Info: starting ---' -Level Action

    $script:GdiQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $script:GdiStop  = [hashtable]::Synchronized(@{ Stop = $false })

    $script:GdiRs = [runspacefactory]::CreateRunspace()
    $script:GdiRs.ApartmentState = 'STA'
    $script:GdiRs.ThreadOptions  = 'ReuseThread'
    $script:GdiRs.Open()
    $script:GdiRs.SessionStateProxy.SetVariable('GdiQueue',  $script:GdiQueue)
    $script:GdiRs.SessionStateProxy.SetVariable('GdiStop',   $script:GdiStop)
    $script:GdiRs.SessionStateProxy.SetVariable('GdiParams', $script:GDIParams)
    $script:GdiRs.SessionStateProxy.SetVariable('LogFile',   $script:LogFile)

    $script:GdiPs = [powershell]::Create()
    $script:GdiPs.Runspace = $script:GdiRs

    $null = $script:GdiPs.AddScript({
        # ── Helper: enqueue a log message ──
        function GLog {
            param([string]$Msg, [string]$Level = 'Info')
            $prefix = switch ($Level) {
                "Success" { "$([char]0x2713) " } "Warning" { "[!] " } "Error" { "$([char]0x2717) " }
                'Action'  { '> ' } default   { '  ' }
            }
            $line = "[$(Get-Date -Format 'HH:mm:ss')] $prefix$Msg"
            $GdiQueue.Enqueue(@{ Type='log'; Line=$line; Level=$Level })
            if ($LogFile) {
                try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
            }
        }

        # ── OS normalisation (Entra raw → friendly) ──
        function Normalize-OS([string]$os) {
            switch -Wildcard ($os) {
                'Windows*' { 'Windows' } 'Android*' { 'Android' }
                'iPhone'   { 'iOS'     } 'iPadOS'   { 'iOS'     }
                'iOS'      { 'iOS'     } 'MacMDM'   { 'macOS'   }
                'Mac OS X' { 'macOS'   } 'macOS'    { 'macOS'   }
                default    { $os }
            }
        }

        # ── Windows release lookup ──
        $WinVerMap = @(
            @{ V='10.0.26200'; R='25H2'          }
            @{ V='10.0.28000'; R='26H1'           }
            @{ V='10.0.26100'; R='24H2'           }
            @{ V='10.0.22631'; R='23H2'           }
            @{ V='10.0.22621'; R='22H2 (Win11)'   }
            @{ V='10.0.19045'; R='22H2 (Win10)'   }
            @{ V='10.0.19044'; R='21H2 (LTSC)'    }
            @{ V='10.0.17763'; R='1809 (LTSC)'    }
        )
        function Get-WinRelease([string]$ver) {
            if ([string]::IsNullOrWhiteSpace($ver)) { return $null }
            $v = $ver.Trim()
            $m = $WinVerMap | Where-Object { $_.V -eq $v } | Select-Object -First 1
            if ($m) { return $m.R }
            # Try major.minor.build prefix only
            if ($v -match '^\d+\.\d+\.\d+') {
                $pfx = ($Matches[0])
                $m = $WinVerMap | Where-Object { $_.V -eq $pfx } | Select-Object -First 1
                if ($m) { return $m.R }
            }
            return $null
        }

        # ── Device type from name suffix ──
        $DevTypeMap = @(
            @{ S='-D'; T='Desktop'            }
            @{ S='-L'; T='Laptop'             }
            @{ S='-V'; T='Virtual'            }
            @{ S='-K'; T='Kiosk'              }
            @{ S='-F'; T='Teams Meeting Room' }
            @{ S='-A'; T='Shared Device'      }
        )
        function Get-DevType([string]$name) {
            if ([string]::IsNullOrWhiteSpace($name)) { return $null }
            $u = $name.Trim().ToUpperInvariant()
            foreach ($m in $DevTypeMap) {
                if ($u.EndsWith($m.S.ToUpperInvariant())) { return $m.T }
            }
            return $null
        }

        # ── Activity range label ──
        # Tries RoundtripKind first (ISO 8601/UTC), then InvariantCulture general parse
        # as fallback  -  same two-style approach as Format-Date to handle all Graph date formats.
        function Get-ActivityRange([string]$dateStr) {
            if ([string]::IsNullOrWhiteSpace($dateStr)) { return $null }
            $d = $null
            foreach ($style in @(
                [System.Globalization.DateTimeStyles]::RoundtripKind,
                [System.Globalization.DateTimeStyles]::None
            )) {
                try {
                    $d = [datetime]::Parse($dateStr,
                             [System.Globalization.CultureInfo]::InvariantCulture, $style)
                    break
                } catch {}
            }
            if ($null -eq $d) { return $null }
            $days = [int][Math]::Floor(([datetime]::UtcNow - $d.ToUniversalTime()).TotalDays)
            if ($days -le 7)       { return '0-7 days'        }
            elseif ($days -le 14)  { return '8-14 days'       }
            elseif ($days -le 30)  { return '15-30 days'      }
            elseif ($days -le 60)  { return '30-60 days'      }
            elseif ($days -le 90)  { return '60-90 days'      }
            elseif ($days -le 180) { return '90-180 days'     }
            elseif ($days -le 365) { return '180-365 days'    }
            else                   { return 'Older than 1 yr' }
        }

        # ── Page through a Graph URI ──
        function Get-GraphAll([string]$Uri, [hashtable]$Headers = @{}) {
            $items = [System.Collections.Generic.List[object]]::new()
            $next  = $Uri
            do {
                $r = Invoke-MgGraphRequest -Method GET -Uri $next -Headers $Headers
                if ($r['value']) { foreach ($i in @($r['value'])) { $items.Add($i) } }
                $next = if ($r.ContainsKey('@odata.nextLink')) { $r['@odata.nextLink'] } else { $null }
            } while ($next)
            return ,$items
        }

        # ── Fetch full Entra device record ──
        function Get-EntraDevice([string]$Uri) {
            try {
                return Invoke-MgGraphRequest -Method GET -Uri $Uri
            } catch { return $null }
        }

        # ── Fetch Intune device by azureADDeviceId ──
        function Get-IntuneByEntraId([string]$EntraId, [hashtable]$Cache) {
            if ([string]::IsNullOrWhiteSpace($EntraId)) { return $null }
            if ($Cache.ContainsKey($EntraId)) { return $Cache[$EntraId] }
            try {
                $enc  = [Uri]::EscapeDataString($EntraId)
                $resp = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=azureADDeviceId eq '$enc'`&`$select=id,deviceName,operatingSystem,osVersion,lastSyncDateTime,manufacturer,model,serialNumber,managedDeviceOwnerType,enrolledDateTime,azureADDeviceId,userPrincipalName,complianceState`&`$top=1"
                $dev = if ($resp['value'] -and $resp['value'].Count -gt 0) { $resp['value'][0] } else { $null }
                $Cache[$EntraId] = $dev
                return $dev
            } catch { $Cache[$EntraId] = $null; return $null }
        }

        # ── Fetch user info ──
        function Get-UserInfo([string]$UserId) {
            try {
                return Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/users/$([Uri]::EscapeDataString($UserId))?`$select=id,userPrincipalName,displayName,jobTitle,accountEnabled,city,country,officeLocation"
            } catch { return $null }
        }

        # ── Format any date string to MM-dd-yyyy, tolerating multiple input formats ──
        function Format-Date([string]$raw) {
            if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
            # Try ISO 8601 / RoundtripKind first, then general parse as fallback
            foreach ($style in @(
                [System.Globalization.DateTimeStyles]::RoundtripKind,
                [System.Globalization.DateTimeStyles]::None
            )) {
                try {
                    return ([datetime]::Parse($raw, [System.Globalization.CultureInfo]::InvariantCulture, $style)).ToString('MM-dd-yyyy')
                } catch {}
            }
            return $raw   # return raw if all parsing fails
        }

        # ── Build Autopilot lookup dictionary ──────────────────────────────────────
        # deploymentProfileAssignmentStatus is a NESTED OBJECT containing deploymentProfileId.
        # Profile name is resolved by: $ap['deploymentProfileAssignmentStatus']['deploymentProfileId']
        # → fetch profile display name from windowsAutopilotDeploymentProfiles/{id}.
        # Step 1: load all profiles into id->name map.
        # Step 2: load all identities, extract profileId from nested status object.
        function Build-AutopilotDictionary {
            $dict       = @{}
            $profileMap = @{}

            # ── Step 1: all deployment profiles ──
            try {
                $nextUri = 'https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles?$select=id,displayName&$top=100'
                do {
                    $r = Invoke-MgGraphRequest -Method GET -Uri $nextUri
                    foreach ($p in @($r['value'])) { $profileMap[[string]$p['id']] = [string]$p['displayName'] }
                    $nextUri = if ($r.ContainsKey('@odata.nextLink')) { $r['@odata.nextLink'] } else { $null }
                } while ($nextUri)
                GLog "Autopilot: $($profileMap.Count) profile(s) loaded" 'Info'
            } catch {
                GLog "Autopilot profiles failed: $($_.Exception.Message)" 'Warning'
            }

            # ── Step 2: all device identities ──
            try {
                GLog 'Autopilot: loading device identities...' 'Info'
                $nextUri = 'https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities?$top=500'
                $total   = 0
                do {
                    $resp  = Invoke-MgGraphRequest -Method GET -Uri $nextUri
                    $page  = @($resp['value'])
                    $total += $page.Count
                    foreach ($ap in $page) {
                        $aadId = [string]$ap['azureActiveDirectoryDeviceId']
                        if ([string]::IsNullOrWhiteSpace($aadId)) { continue }

                        # deploymentProfileAssignmentStatus is a nested object  - 
                        # the profileId lives inside it, not as a top-level field
                        $statusObj  = $ap['deploymentProfileAssignmentStatus']
                        $statusStr  = ''
                        $profileId  = $null
                        if ($statusObj -is [System.Collections.IDictionary]) {
                            $profileId = [string]$statusObj['deploymentProfileId']
                            $statusStr = [string]$statusObj['value']   # e.g. 'assigned'
                            if (-not $statusStr) { $statusStr = [string]$statusObj['status'] }
                        } else {
                            $statusStr = [string]$statusObj   # simple string fallback
                        }

                        # Resolve profile name from the profiles map
                        $profileName = if ($profileId -and $profileMap.ContainsKey($profileId)) {
                                           $profileMap[$profileId]
                                       } else { $null }

                        $dict[$aadId.ToLower()] = @{
                            GroupTag      = [string]$ap['groupTag']
                            ProfileStatus = $statusStr
                        }
                    }
                    $nextUri = if ($resp.ContainsKey('@odata.nextLink')) { $resp['@odata.nextLink'] } else { $null }
                } while ($nextUri)
                GLog "Autopilot: $total record(s), $($dict.Count) indexed by Entra Device ID" 'Success'
            } catch {
                GLog "Autopilot identities failed: $($_.Exception.Message)" 'Warning'
            }
            return ,$dict
        }

        # ── Build a result row from Entra + Intune data ──
        function Build-Row([hashtable]$EntraDev, [hashtable]$IntuneDev, [hashtable]$UserInfo, [string]$OwnerUPN = '', [hashtable]$AutopilotInfo = $null) {
            $NoRec = 'No Intune Record'
            $devName = [string]$EntraDev['displayName']

            # Entra fields
            $entraDevId  = [string]$EntraDev['deviceId']    # the GUID used as azureADDeviceId
            $entraObjId  = [string]$EntraDev['id']
            $entraOS     = [string]$EntraDev['operatingSystem']
            $entraOSVer  = [string]$EntraDev['operatingSystemVersion']
            $entraJoin   = switch ([string]$EntraDev['trustType']) {
                'AzureAd'    { 'Entra Joined'  }
                'ServerAd'   { 'Hybrid Joined' }
                'Workplace'  { 'Registered'    }
                default      { [string]$EntraDev['trustType'] }
            }
            $entraOwner  = [string]$EntraDev['deviceOwnership']
            $entraEnabled= [string]$EntraDev['accountEnabled']
            $entraSignIn = [string]$EntraDev['approximateLastSignInDateTime']
            $entraSignInFmt = Format-Date $entraSignIn
            $entraActivity = Get-ActivityRange $entraSignIn

            # Intune fields
            if ($IntuneDev) {
                $intuneId       = [string]$IntuneDev['id']
                $intuneOS       = [string]$IntuneDev['operatingSystem']
                $intuneOSVer    = [string]$IntuneDev['osVersion']
                $intuneSync     = [string]$IntuneDev['lastSyncDateTime']
                $intuneSyncFmt  = if ($intuneSync) { Format-Date $intuneSync } else { $NoRec }
                $intuneActivity = Get-ActivityRange $intuneSync
                $intuneOEM      = [string]$IntuneDev['manufacturer']
                $intuneModel    = [string]$IntuneDev['model']
                $intuneSerial   = [string]$IntuneDev['serialNumber']
                $intuneOwner    = [string]$IntuneDev['managedDeviceOwnerType']
                $intuneEnrolled = Format-Date ([string]$IntuneDev['enrolledDateTime'])
                $intuneUPN      = [string]$IntuneDev['userPrincipalName']
                $intuneCompliance = [string]$IntuneDev['complianceState']
            } else {
                $intuneId       = $NoRec
                $intuneOS       = $NoRec
                $intuneOSVer    = $NoRec
                $intuneSyncFmt  = $NoRec
                $intuneActivity = $NoRec
                $intuneOEM      = $NoRec
                $intuneModel    = $NoRec
                $intuneSerial   = $NoRec
                $intuneOwner    = $NoRec
                $intuneEnrolled = $NoRec
                $intuneUPN      = $NoRec
                $intuneCompliance = $NoRec
            }

            # Autopilot fields (from dedicated lookup  -  $null if not Autopilot registered)
            $apGroupTag  = if ($AutopilotInfo) { [string]$AutopilotInfo['GroupTag']     } else { $NoRec }
            $apStatus    = if ($AutopilotInfo) { [string]$AutopilotInfo['ProfileStatus'] } else { $NoRec }

            # Prefer Intune OS version for Windows release mapping
            $osVerForMap = if ($intuneOSVer -and $intuneOSVer -ne $NoRec) { $intuneOSVer } else { $entraOSVer }
            $winRelease  = Get-WinRelease $osVerForMap
            $devType     = Get-DevType $devName

            # User fields  -  UserInfo from Entra; fall back to OwnerUPN (Entra) or Intune UPN
            $resolvedUPN = if ($UserInfo) { [string]$UserInfo['userPrincipalName'] }
                           elseif ($OwnerUPN) { $OwnerUPN }
                           elseif ($intuneUPN -and $intuneUPN -ne $NoRec) { $intuneUPN }
                           else { $null }
            $uTitle   = if ($UserInfo) { [string]$UserInfo['jobTitle'] }       else { $null }
            $uEnabled = if ($UserInfo) { [string]$UserInfo['accountEnabled'] } else { $null }
            $uCity    = if ($UserInfo) { [string]$UserInfo['city'] }           else { $null }
            $uCountry = if ($UserInfo) { [string]$UserInfo['country'] }        else { $null }

            return [PSCustomObject]@{
                'Entra_DeviceName'            = $devName
                'DeviceType'                  = $devType
                'Entra_DeviceID'              = $entraDevId
                'Entra_ObjectID'              = $entraObjId
                'Intune_DeviceID'             = $intuneId
                'Entra_JoinType'              = $entraJoin
                'Entra_Ownership'             = $entraOwner
                'Intune_Ownership'            = $intuneOwner
                'Entra_OSPlatform'            = $entraOS
                'Intune_OSPlatform'           = $intuneOS
                'Entra_OSVersion'             = $entraOSVer
                'Intune_OSVersion'            = $intuneOSVer
                'Windows_Release'             = $winRelease
                'Entra_DeviceEnabled'         = $entraEnabled
                'Entra_LastSignIn'            = $entraSignInFmt
                'Entra_ActivityRange'         = $entraActivity
                'Intune_LastCheckIn'          = $intuneSyncFmt
                'Intune_ActivityRange'        = $intuneActivity
                'Intune_OEM'                  = $intuneOEM
                'Intune_Model'                = $intuneModel
                'Intune_SerialNumber'         = $intuneSerial
                'Intune_EnrolledDate'         = $intuneEnrolled
                'Intune_ComplianceState'      = $intuneCompliance
                'Entra_DeviceOwner'           = $OwnerUPN
                'Entra_PrimaryUserUPN'        = $resolvedUPN
                'Entra_UserAccountEnabled'    = $uEnabled
                'Entra_UserCountry'           = $uCountry
                'Entra_UserCity'              = $uCity
                'Entra_UserJobTitle'          = $uTitle
                'Autopilot_GroupTag'          = $apGroupTag
                'Autopilot_ProfileStatus'     = $apStatus
            }
        }

        # ══════════════════════════════════════════════════════════════════
        # MAIN WORK
        # ══════════════════════════════════════════════════════════════════
        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue

            $params         = $GdiParams
            $platFilter     = @($params['PlatformFilter'])
            $ownerFilter    = $params['OwnershipFilter']
            $allDevicesMode = $params['AllDevices'] -eq $true

            $intuneCache    = @{}  # azureADDeviceId -> Intune device hashtable
            $userCache      = @{}  # userId / UPN -> user hashtable
            # Build Autopilot lookup once upfront  -  keyed by azureActiveDirectoryDeviceId (lowercase)
            $autopilotDict  = Build-AutopilotDictionary
            $entraIdSet  = [System.Collections.Generic.HashSet[string]]::new()  # deduplicate Entra Object IDs

            $deviceObjIds = [System.Collections.Generic.List[string]]::new()
            $isGuidRx = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'

            if ($allDevicesMode) {
                # ── All Devices mode: fetch every Entra device object ──
                GLog 'All Devices mode  -  fetching all Entra device objects...' 'Action'
                GLog "Platform filter : $(if ($platFilter.Count -eq 0) { 'All' } else { $platFilter -join ', ' })" 'Info'
                GLog "Ownership filter: $ownerFilter" 'Info'

                $allEntra = Get-GraphAll 'https://graph.microsoft.com/v1.0/devices?$select=id&$top=999'
                GLog "Retrieved $($allEntra.Count) Entra device object(s)" 'Info'
                foreach ($d in $allEntra) {
                    $null = $deviceObjIds.Add([string]$d['id'])
                }
            } else {
                # ── Input-based mode: resolve entries to Entra device Object IDs ──
                $inputType      = $params['InputType']
                $entries        = $params['InputTxt'] -split "`n" |
                                  ForEach-Object { $_.Trim() } |
                                  Where-Object   { $_ -ne '' }

                GLog "Input type      : $inputType" 'Info'
                GLog "Entries         : $($entries.Count)" 'Info'
                GLog "Platform filter : $(if ($platFilter.Count -eq 0) { 'All' } else { $platFilter -join ', ' })" 'Info'
                GLog "Ownership filter: $ownerFilter" 'Info'

                foreach ($entry in $entries) {
                    if ($GdiStop['Stop']) { GLog 'Stop requested.' 'Warning'; break }

                    switch -Wildcard ($inputType) {

                    'User UPNs*' {
                        GLog "Resolving user devices: $entry" 'Info'
                        try {
                            $uid = [Uri]::EscapeDataString($entry)
                            $devs = Get-GraphAll "https://graph.microsoft.com/v1.0/users/$uid/ownedDevices?`$select=id,displayName,deviceId,operatingSystem,operatingSystemVersion,deviceOwnership,trustType,accountEnabled,approximateLastSignInDateTime`&`$top=999"
                            GLog "  $($devs.Count) owned device(s) found" 'Info'
                            foreach ($d in $devs) { $null = $deviceObjIds.Add([string]$d['id']) }
                        } catch {
                            GLog "  Could not resolve user $entry : $($_.Exception.Message)" 'Warning'
                        }
                    }

                    'Device names' {
                        GLog "Resolving device by name: $entry" 'Info'
                        try {
                            $safe = $entry -replace "'","''"
                            $r = Invoke-MgGraphRequest -Method GET `
                                -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=displayName eq '$safe'`&`$select=id`&`$top=5"
                            $found = @($r['value'])
                            if ($found.Count -eq 0) { GLog "  Not found: $entry" 'Warning' }
                            else { foreach ($d in $found) { $null = $deviceObjIds.Add([string]$d['id']) } }
                        } catch {
                            GLog "  Error resolving $entry : $($_.Exception.Message)" 'Warning'
                        }
                    }

                    'Device serial numbers' {
                        GLog "Resolving serial number: $entry" 'Info'
                        # Serial numbers only exist in Intune  -  search there first, get azureADDeviceId, then Entra
                        try {
                            $safe = $entry -replace "'","''"
                            $r = Invoke-MgGraphRequest -Method GET `
                                -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=serialNumber eq '$safe'`&`$select=id,azureADDeviceId,deviceName`&`$top=5"
                            $found = @($r['value'])
                            if ($found.Count -eq 0) { GLog "  Serial not found in Intune: $entry" 'Warning' }
                            else {
                                foreach ($md in $found) {
                                    $aadId = [string]$md['azureADDeviceId']
                                    if ($aadId) {
                                        # Find Entra Object ID from azureADDeviceId (deviceId field)
                                        $er = Invoke-MgGraphRequest -Method GET `
                                            -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$aadId'`&`$select=id`&`$top=1"
                                        if ($er['value'] -and $er['value'].Count -gt 0) {
                                            $null = $deviceObjIds.Add([string]$er['value'][0]['id'])
                                        }
                                    }
                                }
                            }
                        } catch {
                            GLog "  Error resolving serial $entry : $($_.Exception.Message)" 'Warning'
                        }
                    }

                    'Entra Device Object IDs' {
                        if ($entry -match $isGuidRx) {
                            $null = $deviceObjIds.Add($entry)
                        } else {
                            GLog "  Skipped (not a GUID): $entry" 'Warning'
                        }
                    }

                    'Intune Device IDs' {
                        GLog "Resolving Intune device ID: $entry" 'Info'
                        try {
                            $md = Invoke-MgGraphRequest -Method GET `
                                -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$entry`?`$select=id,azureADDeviceId,deviceName"
                            $aadId = [string]$md['azureADDeviceId']
                            if ($aadId) {
                                $er = Invoke-MgGraphRequest -Method GET `
                                    -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$aadId'`&`$select=id`&`$top=1"
                                if ($er['value'] -and $er['value'].Count -gt 0) {
                                    $null = $deviceObjIds.Add([string]$er['value'][0]['id'])
                                }
                            }
                        } catch {
                            GLog "  Could not resolve Intune ID $entry : $($_.Exception.Message)" 'Warning'
                        }
                    }

                    'Groups*' {
                        GLog "Resolving group: $entry" 'Info'
                        # Resolve group ID
                        $groupId = $null
                        if ($entry -match $isGuidRx) {
                            $groupId = $entry
                        } else {
                            try {
                                $safe = $entry -replace "'","''"
                                $gr = Invoke-MgGraphRequest -Method GET `
                                    -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe'`&`$select=id`&`$top=1"
                                if ($gr['value'] -and $gr['value'].Count -gt 0) {
                                    $groupId = [string]$gr['value'][0]['id']
                                }
                            } catch {}
                        }
                        if (-not $groupId) { GLog "  Group not found: $entry" 'Warning'; continue }

                        # Get all transitive members (users + devices)
                        GLog "  Scanning group members (transitive)..." 'Info'
                        # $select with @odata.type or user-subtype properties causes 400 on
                        # transitiveMembers (directoryObject base type). Use $select=id only;
                        # @odata.type is always returned automatically.
                        $members = Get-GraphAll "https://graph.microsoft.com/v1.0/groups/$groupId/transitiveMembers?`$select=id`&`$top=999"
                        GLog "  $($members.Count) member(s) found" 'Info'

                        foreach ($mem in $members) {
                            if ($GdiStop['Stop']) { break }
                            $odataType = [string]$mem['@odata.type']
                            if ($odataType -eq '#microsoft.graph.device') {
                                $null = $deviceObjIds.Add([string]$mem['id'])
                            } elseif ($odataType -eq '#microsoft.graph.user') {
                                # Resolve user's owned devices
                                $uid = [string]$mem['id']
                                try {
                                    $udevs = Get-GraphAll "https://graph.microsoft.com/v1.0/users/$uid/ownedDevices?`$select=id`&`$top=999"
                                    foreach ($ud in $udevs) { $null = $deviceObjIds.Add([string]$ud['id']) }
                                } catch {}
                            }
                        }
                    }
                }
            }

                        }  # end else (input-based mode)

            GLog "Total device IDs to process: $($deviceObjIds.Count)" 'Info'
            if ($deviceObjIds.Count -eq 0) {
                .Enqueue(@{ Type='done'; Rows=@(); Error='No devices resolved from the provided input.' })
                return
            }

            # ── Process each Entra device ──
            $results = [System.Collections.Generic.List[PSCustomObject]]::new()
            $processed = 0

            foreach ($objId in $deviceObjIds) {
                if ($GdiStop['Stop']) { GLog 'Stop requested  -  halting.' 'Warning'; break }
                if ($entraIdSet.Contains($objId)) { continue }   # deduplicate
                $null = $entraIdSet.Add($objId)

                $processed++
                $GdiQueue.Enqueue(@{ Type='progress'; Current=$processed; Total=$deviceObjIds.Count })

                # Fetch full Entra device record
                $entraRaw = $null
                try {
                    $entraRaw = Invoke-MgGraphRequest -Method GET `
                        -Uri "https://graph.microsoft.com/v1.0/devices/$objId`?`$select=id,deviceId,displayName,operatingSystem,operatingSystemVersion,deviceOwnership,trustType,accountEnabled,approximateLastSignInDateTime,registrationDateTime"
                } catch {
                    GLog "  Could not fetch Entra device $objId : $($_.Exception.Message)" 'Warning'
                    continue
                }

                $devName   = [string]$entraRaw['displayName']
                $entraDevId = [string]$entraRaw['deviceId']   # azureADDeviceId in Intune
                $entraOS   = Normalize-OS ([string]$entraRaw['operatingSystem'])

                GLog "[$processed] $devName  [$entraOS]" 'Info'

                # ── Platform filter ──
                if ($platFilter.Count -gt 0 -and $entraOS -notin $platFilter) {
                    GLog "  Skipped (platform $entraOS not in filter)" 'Info'
                    continue
                }

                # ── Ownership filter ──
                $rawOwner = [string]$entraRaw['deviceOwnership']
                if ($ownerFilter -ne 'All' -and $rawOwner -ne $ownerFilter) {
                    GLog "  Skipped (ownership $rawOwner not in filter)" 'Info'
                    continue
                }

                # ── Intune lookup ──
                $intuneDev = Get-IntuneByEntraId -EntraId $entraDevId -Cache $intuneCache
                if ($intuneDev) {
                    GLog "  Intune: $([string]$intuneDev['id'])" 'Success'
                } else {
                    GLog "  Intune: No Intune Record" 'Warning'
                }

                # ── Registered owner (user) ──
                # No $select on registeredOwners  -  $select is not supported on navigation
                # endpoints and causes userPrincipalName to be missing from the response.
                $userInfo    = $null
                $ownerUPN    = $null
                try {
                    $ownerResp = Invoke-MgGraphRequest -Method GET `
                        -Uri "https://graph.microsoft.com/v1.0/devices/$objId/registeredOwners?`$top=5"
                    $ownerList = @($ownerResp['value'])
                    $firstUser = $ownerList | Where-Object {
                        $t = [string]$_['@odata.type']
                        $t -eq '#microsoft.graph.user' -or $t -eq 'microsoft.graph.user'
                    } | Select-Object -First 1
                    if ($firstUser) {
                        # userPrincipalName lives in AdditionalProperties for some SDK versions,
                        # or directly on the hashtable when using Invoke-MgGraphRequest
                        $ownerUPN = if ($firstUser.ContainsKey('userPrincipalName')) {
                                        [string]$firstUser['userPrincipalName']
                                    } elseif ($firstUser.ContainsKey('AdditionalProperties')) {
                                        [string]$firstUser['AdditionalProperties']['userPrincipalName']
                                    } else { $null }
                        $uid = [string]$firstUser['id']
                        # If UPN still empty, fetch it from the user endpoint directly
                        if ([string]::IsNullOrWhiteSpace($ownerUPN) -and $uid) {
                            try {
                                $upnResp  = Invoke-MgGraphRequest -Method GET `
                                    -Uri "https://graph.microsoft.com/v1.0/users/$uid`?`$select=userPrincipalName"
                                $ownerUPN = [string]$upnResp['userPrincipalName']
                            } catch {}
                        }
                        GLog "  Owner: $ownerUPN" 'Info'
                        if ($userCache.ContainsKey($uid)) {
                            $userInfo = $userCache[$uid]
                        } else {
                            $userInfo = Get-UserInfo -UserId $uid
                            $userCache[$uid] = $userInfo
                        }
                    } else {
                        GLog "  No registered owner found" 'Info'
                    }
                } catch {
                    GLog "  Owner fetch error: $($_.Exception.Message)" 'Warning'
                }

                # ── Autopilot lookup  -  dictionary built upfront, keyed by entraDevId lowercase ──
                $autopilotInfo = $null
                $entraOSNorm = Normalize-OS ([string]$entraRaw['operatingSystem'])
                if ($entraOSNorm -eq 'Windows' -and $entraDevId) {
                    $apKey = $entraDevId.ToLower()
                    if ($autopilotDict.ContainsKey($apKey)) {
                        $autopilotInfo = $autopilotDict[$apKey]
                        GLog "  Autopilot: Tag=$($autopilotInfo['GroupTag']) | Status=$($autopilotInfo['ProfileStatus'])" 'Info'
                    } else {
                        GLog '  Autopilot: not registered' 'Info'
                    }
                }

                # ── Build row ──
                $row = Build-Row -EntraDev $entraRaw -IntuneDev $intuneDev -UserInfo $userInfo -OwnerUPN $ownerUPN -AutopilotInfo $autopilotInfo
                $results.Add($row)
            }

            GLog "Processed $($results.Count) device(s)." 'Success'
            $GdiQueue.Enqueue(@{ Type='done'; Rows=$results; Error=$null })

        } catch {
            $GdiQueue.Enqueue(@{ Type='done'; Rows=@(); Error=$_.Exception.Message })
        }
    })

    $script:GdiHandle = $script:GdiPs.BeginInvoke()

    # Timer to drain GdiQueue on UI thread
    $script:GdiTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:GdiTimer.Interval = [TimeSpan]::FromMilliseconds(300)
    $script:GdiTimer.Add_Tick({
        $entry = $null
        while ($script:GdiQueue.TryDequeue([ref]$entry)) {
            switch ($entry['Type']) {
                'log' {
                    $lvl = $entry['Level']
                    $col = switch ($lvl) {
                        'Success' { $DevConfig.LogColorSuccess }
                        'Warning' { $DevConfig.LogColorWarning }
                        'Error'   { $DevConfig.LogColorError   }
                        'Action'  { $DevConfig.LogColorAction  }
                        default   { $DevConfig.LogColorInfo    }
                    }
                    try {
                        $para = [System.Windows.Documents.Paragraph]::new()
                        $run  = [System.Windows.Documents.Run]::new($entry['Line'])
                        $run.Foreground = [System.Windows.Media.SolidColorBrush]::new(
                            [System.Windows.Media.ColorConverter]::ConvertFromString($col))
                        $para.Margin = [System.Windows.Thickness]::new(0)
                        $para.Inlines.Add($run)
                        $script:RtbLog.Document.Blocks.Add($para)
                        $script:RtbLog.ScrollToEnd()
                    } catch {}
                }
                'progress' {
                    $cur = $entry['Current']; $tot = $entry['Total']
                    $script:Window.FindName('PnlGDIProgress').Visibility       = 'Visible'
                    $script:Window.FindName('TxtGDIProgressMsg').Text          = 'Fetching device details...'
                    $script:Window.FindName('TxtGDIProgressDetail').Text       = "Processing device $cur of $tot..."
                    Update-StatusBar -Text "Get Device Info: $cur / $tot devices..."
                }
                'done' {
                    $script:GdiTimer.Stop()
                    try { $script:GdiPs.EndInvoke($script:GdiHandle) } catch {}
                    try { $script:GdiRs.Close() }                      catch {}
                    try { $script:GdiPs.Dispose() }                    catch {}
                    $script:GdiPs = $null; $script:GdiRs = $null
                    $script:Window.FindName('BtnStop').Visibility = 'Collapsed'
                    $script:Window.FindName('BtnStop').IsEnabled  = $true
                    $script:Window.FindName('BtnStop').Content = (New-StopBtnContent 'Stop')
                    $script:Window.FindName('BtnGDIRun').IsEnabled = $true
                    $script:Window.FindName('BtnGDIRunAll').IsEnabled = $true
                    $script:Window.FindName('PnlGDIProgress').Visibility = 'Collapsed'

                    $errMsg = $entry['Error']
                    if ($errMsg) {
                        Write-VerboseLog "Get Device Info error: $errMsg" -Level Error
                        Show-Notification "Error: $errMsg" -BgColor '#F8D7DA' -FgColor '#721C24'
                        Update-StatusBar -Text "Connected as: $($script:TxtStatusBar.Text -replace 'Get Device Info.*','')"
                        return
                    }

                    $rows     = @($entry['Rows'])

                    if ($rows.Count -eq 0) {
                        Write-VerboseLog 'Get Device Info: no rows  -  check filters or input.' -Level Warning
                        $script:Window.FindName('TxtGDINoResults').Visibility = 'Visible'
                        Show-Notification 'No devices matched the input and filters.' -BgColor '#FFF3CD' -FgColor '#7A4800'
                        Update-StatusBar -Text 'Connected'
                        return
                    }

                    # ── Populate DataGrid ──
                    $dg = $script:Window.FindName('DgGDIResults')
                    $dg.ItemsSource = [System.Collections.ObjectModel.ObservableCollection[PSCustomObject]]::new($rows)
                    $script:Window.FindName('TxtGDICount').Text            = "$($rows.Count) device(s) found"
                    $script:Window.FindName('PnlGDIResults').Visibility     = 'Visible'
                    $script:Window.FindName('BtnGDICopyAll').IsEnabled      = $true
                    $script:Window.FindName('TxtGDIFilter').IsEnabled  = $true
                    $script:Window.FindName('BtnGDIFilter').IsEnabled  = $true
                    $script:Window.FindName('BtnGDIFilterClear').IsEnabled = $true
                    $script:Window.FindName('BtnGDIExportXlsx').IsEnabled   = $true
                    Write-VerboseLog "Get Device Info: $($rows.Count) device(s) loaded into table." -Level Success
                    Show-Notification "$($rows.Count) device(s) ready. Use Export XLSX to save." -BgColor '#D4EDDA' -FgColor '#155724'
                    Update-StatusBar -Text 'Connected'
                }
            }
        }
    })
    $script:GdiTimer.Start()
})

# ── DgGDIResults: cell selection ─────────────────────────────────────────────
$script:Window.FindName('DgGDIResults').Add_SelectedCellsChanged({
    $dg       = $script:Window.FindName('DgGDIResults')
    $selCells = $dg.SelectedCells.Count
    $script:Window.FindName('BtnGDICopyValue').IsEnabled = ($selCells -ge 1)
    $script:Window.FindName('BtnGDICopyRow').IsEnabled   = ($selCells -gt 0)
})

# ── BtnGDICopyValue ──────────────────────────────────────────────────────────
# ── BtnGDIFilter ──
$script:Window.FindName('BtnGDIFilter').Add_Click({
    $dg      = $script:Window.FindName('DgGDIResults')
    $keyword = $script:Window.FindName('TxtGDIFilter').Text.Trim()
    if ($null -eq $script:GDIAllData) {
        $script:GDIAllData = @($dg.ItemsSource)
    }
    if ($keyword -eq '') {
        $dg.ItemsSource = $script:GDIAllData
        Show-Notification "Filter cleared - $($script:GDIAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    } else {
        $filtered = @($script:GDIAllData | Where-Object {
            $row = $_
            $match = $false
            foreach ($prop in $row.PSObject.Properties) {
                if ([string]$prop.Value -like "*$keyword*") { $match = $true; break }
            }
            $match
        })

# ── BtnGDIFilterClear ──
$script:Window.FindName('BtnGDIFilterClear').Add_Click({
    $script:Window.FindName('TxtGDIFilter').Text = ''
    $dg = $script:Window.FindName('DgGDIResults')
    if ($null -ne $script:GDIAllData) {
        $dg.ItemsSource = $script:GDIAllData
        Show-Notification "Filter cleared - $($script:GDIAllData.Count) result(s) shown." -BgColor '#D4EDDA' -FgColor '#155724'
    }
})
        $dg.ItemsSource = $filtered
        Show-Notification "$($filtered.Count) of $($script:GDIAllData.Count) result(s) shown after filter." -BgColor '#D4EDDA' -FgColor '#155724'
    }
    Write-VerboseLog "GDI: Filter applied - keyword='$keyword'" -Level Info
})

$script:Window.FindName('BtnGDICopyValue').Add_Click({
    $dg = $script:Window.FindName('DgGDIResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $vals = [System.Collections.Generic.List[string]]::new()
    foreach ($cellInfo in @($dg.SelectedCells)) {
        $col = $cellInfo.Column -as [System.Windows.Controls.DataGridBoundColumn]
        if ($col) { $vals.Add(([string]$cellInfo.Item.($col.Binding.Path.Path)).Trim()) }
    }
    [System.Windows.Clipboard]::SetText($vals -join "`r`n")
    Show-Notification '$($vals.Count) value(s) copied to clipboard.' -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog 'GDI: Copied cell value to clipboard.' -Level Success
})

# ── BtnGDICopyRow ────────────────────────────────────────────────────────────
$script:Window.FindName('BtnGDICopyRow').Add_Click({
    $dg = $script:Window.FindName('DgGDIResults')
    if ($dg.SelectedCells.Count -eq 0) { return }
    $seen = [System.Collections.Generic.HashSet[object]]::new()
    $rowObjs = [System.Collections.Generic.List[object]]::new()
    foreach ($cell in @($dg.SelectedCells)) { if ($seen.Add($cell.Item)) { $rowObjs.Add($cell.Item) } }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc     = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $txtLines = [System.Collections.Generic.List[string]]::new()
    $txtLines.Add(($colDefs | ForEach-Object { $_.H }) -join "	")
    foreach ($row in $rowObjs) {
        $vals = $colDefs | ForEach-Object { [string]$row.($_.P) }
        $txtLines.Add($vals -join "	")
    }
    [System.Windows.Clipboard]::SetText($txtLines -join "
")
    Show-Notification "$($rowObjs.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "GDI: Copied $($rowObjs.Count) row(s) to clipboard." -Level Success
    # Highlight full rows
    $dg.UnselectAllCells()
    foreach ($r in $rowObjs) { $null = $dg.SelectedItems.Add($r) }
})

# ── BtnGDICopyAll ────────────────────────────────────────────────────────────
$script:Window.FindName('BtnGDICopyAll').Add_Click({
    $dg       = $script:Window.FindName('DgGDIResults')
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) { return }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc     = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $txtLines = [System.Collections.Generic.List[string]]::new()
    $txtLines.Add(($colDefs | ForEach-Object { $_.H }) -join "	")
    foreach ($row in $allItems) {
        $vals = $colDefs | ForEach-Object { [string]$row.($_.P) }
        $txtLines.Add($vals -join "	")
    }
    [System.Windows.Clipboard]::SetText($txtLines -join "
")
    Show-Notification "$($allItems.Count) row(s) copied to clipboard." -BgColor '#D4EDDA' -FgColor '#155724'
    Write-VerboseLog "GDI: Copied all $($allItems.Count) row(s) to clipboard." -Level Success
})

# ── BtnGDIExportXlsx ─────────────────────────────────────────────────────────
$script:Window.FindName('BtnGDIExportXlsx').Add_Click({
    $dg       = $script:Window.FindName('DgGDIResults')
    $allItems = @($dg.ItemsSource)
    if ($allItems.Count -eq 0) {
        Show-Notification 'No results to export.' -BgColor '#FFF3CD' -FgColor '#856404'
        return
    }
    $colDefs = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($col in @($dg.Columns)) {
        $bc     = $col -as [System.Windows.Controls.DataGridBoundColumn]
        $bcPath = ''
        if ($bc) { $bcPath = $bc.Binding.Path.Path }
        $colDefs.Add(@{ H = [string]$col.Header; P = $bcPath })
    }
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title            = 'Export Get Device Info Results'
    $dlg.Filter           = 'Excel Workbook (*.xlsx)|*.xlsx|CSV File (*.csv)|*.csv|All Files (*.*)|*.*'
    $dlg.FileName         = "DeviceInfo_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($dlg.ShowDialog() -ne $true) { return }
    $outPath = $dlg.FileName
    try {
        $folder = Split-Path $outPath -Parent
        if ($folder -and -not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        $exportRows = $allItems | ForEach-Object {
            $rowObj = $_; $ht = [ordered]@{}
            foreach ($cd in $colDefs) { $ht[$cd['H']] = [string]$rowObj.($cd['P']) }
            [PSCustomObject]$ht
        }
        if (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue) {
            Import-Module ImportExcel -ErrorAction SilentlyContinue
            $exportRows | Export-Excel -Path $outPath -WorksheetName 'DeviceInfo' -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium6
            Show-Notification "Exported $($allItems.Count) rows: $outPath" -BgColor '#D4EDDA' -FgColor '#155724'
        } else {
            $exportRows | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8
            Show-Notification "ImportExcel not found — saved as CSV: $(Split-Path $outPath -Leaf)" -BgColor '#FFF3CD' -FgColor '#856404'
        }
        Write-VerboseLog "GDI: Exported $($allItems.Count) row(s) to: $outPath" -Level Success
    } catch {
        Write-VerboseLog "GDI Export failed: $($_.Exception.Message)" -Level Error
        Show-Notification "Export failed: $($_.Exception.Message)" -BgColor '#F8D7DA' -FgColor '#721C24'
    }
})

#endregion


# ============================================================
#region LAUNCH
# ============================================================

Write-VerboseLog "========================================" -Level Info
Write-VerboseLog "$($DevConfig.WindowTitle)" -Level Info
Write-VerboseLog "Author: Satish Singhi" -Level Info
Write-VerboseLog "Session: $script:SessionStamp" -Level Info
if ($script:LogFile) { Write-VerboseLog "Log file: $script:LogFile" -Level Info }
Write-VerboseLog "========================================" -Level Info
Write-VerboseLog "Click 'Connect' to authenticate to Microsoft Graph." -Level Info

$script:Window.Add_Closed({
    Stop-PimAutoRefresh
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    Write-VerboseLog "Session ended. Disconnected." -Level Info
})

$null = $script:Window.ShowDialog()

#endregion
#========== END_GUI_TOOL_CONTENT ===========
__PS7_PAYLOAD_END__#>

