<#
.SYNOPSIS
    Windows health check and repair: DISM, SFC, chkdsk and WMI, with optional prompts.

.DESCRIPTION
    Sequence: pre-checks -> (WU reset) -> DISM ScanHealth -> RestoreHealth (if needed)
    -> (ComponentCleanup) -> SFC -> chkdsk /scan -> WMI verify (-> salvage).

    INTERACTIVE MODE
    When the script runs in a user session (session ID <> 0), every repair step
    asks Yes/No in a dialog. Each dialog auto-answers with its default after
    $PromptTimeoutSec seconds, so an unattended device never hangs.
      - Safe steps (scan, SFC, chkdsk scan) default to Yes.
      - Disruptive steps default to the value of their switch below.
    In session 0 (SYSTEM, no UI) or with -NoPrompt, defaults are used silently.

    Switches can also be set through WS1 script variables named WS1R_<SwitchName>
    with value true/1/yes (e.g. WS1R_ResetWindowsUpdate = true).

    Exit codes: 0 = healthy, repaired or cancelled by user
                3010 (or -RebootExitCode) = repaired, reboot required
                1 = one or more repairs failed or were declined while needed
                2 = precondition not met (not elevated, low disk space, pending reboot)

    Output: log + last-result.json in C:\ProgramData\WS1Repair (readable by a WS1 Sensor).

.EXAMPLE
    Run directly from the web in an elevated Windows PowerShell 5.1 console:
    irm https://raw.githubusercontent.com/<user>/<repo>/main/Invoke-WindowsRepair.ps1 | iex

.EXAMPLE
    Run as a file with parameters:
    .\Invoke-WindowsRepair.ps1 -ResetWindowsUpdate -NoPrompt

.NOTES
    Requires admin rights and Windows PowerShell 5.1 (powershell.exe).
    When deployed through an MDM such as Workspace ONE UEM, set the script timeout
    to at least 60 minutes plus the time users may take to answer prompts.

    Provided as-is, without warranty. Review the script before running it.
#>
[CmdletBinding()]
param(
    [switch]$ForceRestoreHealth,
    [switch]$ComponentCleanup,        # Default answer for the cleanup prompt
    [switch]$ResetWindowsUpdate,      # Default answer for the WU reset prompt
    [switch]$SalvageWmi,              # Default answer for the WMI salvage prompt
    [switch]$ScheduleChkdskFix,       # Default answer for the offline chkdsk prompt
    [switch]$IgnorePendingReboot,
    [switch]$NoPrompt,                # Never show dialogs, even in a user session
    [int]$PromptTimeoutSec = 120,
    [int]$MinFreeSpaceGB = 5,
    [int]$RebootExitCode = 0,
    [int]$KeepLogs = 10
)

# --- Relaunch in native 64-bit Windows PowerShell 5.1 if needed -------------
# Covers 32-bit hosts on x64 and ARM64 (SysNative points to the native System32) and PowerShell 7.
if ($PSVersionTable.PSEdition -eq 'Core' -or
    ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess)) {
    if (-not $PSCommandPath) {
        throw 'Run this in 64-bit Windows PowerShell 5.1 (powershell.exe). PowerShell 7 and 32-bit PowerShell are not supported with irm | iex.'
    }
    $ps64 = if ([Environment]::Is64BitProcess) { "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" }
            else { "$env:WINDIR\SysNative\WindowsPowerShell\v1.0\powershell.exe" }
    $argList = foreach ($k in $PSBoundParameters.Keys) {
        $v = $PSBoundParameters[$k]
        if ($v -is [System.Management.Automation.SwitchParameter]) { if ($v.IsPresent) { "-$k" } }
        else { "-$k"; "$v" }
    }
    & $ps64 -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath @argList
    exit $LASTEXITCODE
}

# Everything below runs inside try{}. Complete-Run ends the run by throwing a marker
# instead of calling exit, so 'irm | iex' does not close the user's console window.
$ExitCode = 0
try {

# --- WS1 script variables (environment) override ----------------------------
foreach ($n in 'ForceRestoreHealth','ComponentCleanup','ResetWindowsUpdate',
               'SalvageWmi','ScheduleChkdskFix','IgnorePendingReboot','NoPrompt') {
    if ([Environment]::GetEnvironmentVariable("WS1R_$n") -match '^(1|true|yes)$') {
        Set-Variable -Name $n -Value ([switch]$true)
    }
}

# --- Logging and result -----------------------------------------------------
$LogDir  = Join-Path $env:ProgramData 'WS1Repair'
$Stamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
$LogFile = Join-Path $LogDir "Repair-$Stamp.log"
New-Item -Path $LogDir -ItemType Directory -Force | Out-Null

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 's'), $Level, $Message
    Add-Content -Path $LogFile -Value $line
    Write-Host $line
}

$Result = [ordered]@{
    Timestamp      = (Get-Date).ToString('s')
    Computer       = $env:COMPUTERNAME
    OSBuild        = [Environment]::OSVersion.Version.ToString()
    Interactive    = $false
    Status         = 'Running'
    RebootRequired = $false
    Failed         = $false
    Steps          = [ordered]@{}
}

function Set-Step([string]$Name, [string]$Value, [switch]$Fail, [switch]$Reboot) {
    $Result.Steps[$Name] = $Value
    if ($Fail)   { $Result.Failed = $true }
    if ($Reboot) { $Result.RebootRequired = $true }
}

# --- Interactive dialogs ----------------------------------------------------
$Interactive = (-not $NoPrompt) -and ((Get-Process -Id $PID).SessionId -ne 0)
$Result.Interactive = $Interactive
if ($Interactive) {
    try { Add-Type -AssemblyName System.Windows.Forms, System.Drawing -ErrorAction Stop }
    catch { $Interactive = $false; $Result.Interactive = $false }
}


function Show-Dialog {
    param([string]$Text, [ValidateSet('YesNo','OK')][string]$Buttons = 'YesNo',
          [bool]$Default = $true, [int]$Timeout = 120)

    # Layout panels size themselves to the rendered text, so nothing clips at any DPI/scaling
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Windows Repair'
    $form.TopMost = $true
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ControlBox = $false
    $form.AutoSize = $true
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Padding = New-Object System.Windows.Forms.Padding(0, 0, 15, 15)
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $table = New-Object System.Windows.Forms.TableLayoutPanel
    $table.AutoSize = $true
    $table.AutoSizeMode = 'GrowAndShrink'
    $table.ColumnCount = 1
    $table.RowCount = 2
    $table.Location = New-Object System.Drawing.Point(15, 15)


    $label = New-Object System.Windows.Forms.Label
    $label.AutoSize = $true
    $label.MinimumSize = New-Object System.Drawing.Size(450, 0)
    $label.MaximumSize = New-Object System.Drawing.Size(450, 0)
    $label.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 15)
    $label.Text = $Text
    $table.Controls.Add($label, 0, 0)

    $bar = New-Object System.Windows.Forms.FlowLayoutPanel
    $bar.AutoSize = $true
    $bar.AutoSizeMode = 'GrowAndShrink'
    $bar.WrapContents = $false
    $bar.Margin = New-Object System.Windows.Forms.Padding(0)

    $countdown = New-Object System.Windows.Forms.Label
    $countdown.AutoSize = $true
    $countdown.MinimumSize = New-Object System.Drawing.Size(260, 0)
    $countdown.Margin = New-Object System.Windows.Forms.Padding(0, 7, 10, 0)
    $countdown.ForeColor = [System.Drawing.Color]::Gray
    $bar.Controls.Add($countdown)

    function New-DialogButton([string]$Caption) {
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $Caption
        $b.DialogResult = $Caption
        $b.AutoSize = $true
        $b.MinimumSize = New-Object System.Drawing.Size(85, 28)
        $b.Margin = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
        $b
    }

    if ($Buttons -eq 'OK') {
        $countdown.MinimumSize = New-Object System.Drawing.Size(350, 0)
        $ok = New-DialogButton 'OK'
        $bar.Controls.Add($ok)
        $form.AcceptButton = $ok
        $defaultResult = 'OK'
    } else {
        $yes = New-DialogButton 'Yes'
        $no  = New-DialogButton 'No'
        $bar.Controls.Add($yes)
        $bar.Controls.Add($no)
        $form.AcceptButton = if ($Default) { $yes } else { $no }
        $defaultResult = if ($Default) { 'Yes' } else { 'No' }
    }
    $table.Controls.Add($bar, 0, 1)
    $form.Controls.Add($table)

    $state = @{ Left = $Timeout }
    $countdown.Text = "Automatic answer '$defaultResult' in $Timeout s"
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000
    $timer.Add_Tick({
        $state.Left--
        $countdown.Text = "Automatic answer '$defaultResult' in $($state.Left) s"
        if ($state.Left -le 0) { $timer.Stop(); $form.DialogResult = $defaultResult }
    }.GetNewClosure())

    $timer.Start()
    $res = $form.ShowDialog()
    $timer.Stop(); $timer.Dispose(); $form.Dispose()

    @{ Answer = ("$res" -in 'Yes','OK'); TimedOut = ($state.Left -le 0) }
}

function Confirm-Step([string]$Text, [bool]$Default) {
    if (-not $Interactive) { return $Default }
    $r = Show-Dialog -Text $Text -Default $Default -Timeout $PromptTimeoutSec
    $how = if ($r.TimedOut) { 'timeout' } else { 'user' }
    Write-Log ('Prompt [{0}] -> {1} ({2})' -f ($Text -split "`n")[0], $(if ($r.Answer) { 'Yes' } else { 'No' }), $how)
    return $r.Answer
}

# --- Work-in-progress window ------------------------------------------------
# Runs on its own UI thread, so it stays responsive while DISM/SFC block the main thread.
$WorkStatus = [hashtable]::Synchronized(@{ Text = ''; Started = Get-Date; Visible = $false; Close = $false })
$StatusPS = $null

function Start-StatusWindow {
    if (-not $Interactive -or $script:StatusPS) { return }
    try {
        $rs = [runspacefactory]::CreateRunspace()
        $rs.ApartmentState = 'STA'
        $rs.ThreadOptions = 'ReuseThread'
        $rs.Open()
        $rs.SessionStateProxy.SetVariable('WorkStatus', $WorkStatus)
        $ps = [powershell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript({
            Add-Type -AssemblyName System.Windows.Forms, System.Drawing
            [System.Windows.Forms.Application]::EnableVisualStyles()   # needed for the marquee animation

            $form = New-Object System.Windows.Forms.Form
            $form.Text = 'Windows Repair'
            $form.TopMost = $true
            $form.StartPosition = 'CenterScreen'
            $form.FormBorderStyle = 'FixedDialog'
            $form.ControlBox = $false
            $form.AutoSize = $true
            $form.AutoSizeMode = 'GrowAndShrink'
            $form.Padding = New-Object System.Windows.Forms.Padding(0, 0, 15, 15)
            $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

            $scale = 1.0
            try { if ($form.DeviceDpi -gt 0) { $scale = $form.DeviceDpi / 96 } } catch { }

            $table = New-Object System.Windows.Forms.TableLayoutPanel
            $table.AutoSize = $true
            $table.AutoSizeMode = 'GrowAndShrink'
            $table.ColumnCount = 1
            $table.RowCount = 3
            $table.Location = New-Object System.Drawing.Point(15, 15)

            $label = New-Object System.Windows.Forms.Label
            $label.AutoSize = $true
            $label.MinimumSize = New-Object System.Drawing.Size(420, 0)
            $label.MaximumSize = New-Object System.Drawing.Size(420, 0)
            $label.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
            $table.Controls.Add($label, 0, 0)

            $bar = New-Object System.Windows.Forms.ProgressBar
            $bar.Style = 'Marquee'
            $bar.MarqueeAnimationSpeed = 30
            $bar.Size = New-Object System.Drawing.Size(420, [int](16 * $scale))
            $bar.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
            $table.Controls.Add($bar, 0, 1)

            $elapsed = New-Object System.Windows.Forms.Label
            $elapsed.AutoSize = $true
            $elapsed.MinimumSize = New-Object System.Drawing.Size(420, 0)
            $elapsed.MaximumSize = New-Object System.Drawing.Size(420, 0)
            $elapsed.ForeColor = [System.Drawing.Color]::Gray
            $table.Controls.Add($elapsed, 0, 2)
            $form.Controls.Add($table)

            $ctx = New-Object System.Windows.Forms.ApplicationContext
            $form.Add_FormClosing({ param($s, $e) if (-not $WorkStatus.Close) { $e.Cancel = $true } }.GetNewClosure())

            $timer = New-Object System.Windows.Forms.Timer
            $timer.Interval = 500
            $timer.Add_Tick({
                if ($WorkStatus.Close) { $timer.Stop(); $form.Close(); $ctx.ExitThread(); return }
                if ($WorkStatus.Visible) {
                    if ($label.Text -ne $WorkStatus.Text) { $label.Text = $WorkStatus.Text }
                    $elapsed.Text = ('Running for {0:hh\:mm\:ss}. Please do not restart or shut down your computer.' -f ((Get-Date) - $WorkStatus.Started))
                    if (-not $form.Visible) { $form.Show(); $form.Activate() }
                } elseif ($form.Visible) {
                    $form.Hide()
                }
            }.GetNewClosure())
            $timer.Start()
            [System.Windows.Forms.Application]::Run($ctx)
            $form.Dispose()
        })
        $script:StatusAsync = $ps.BeginInvoke()
        $script:StatusPS = $ps
    } catch {
        Write-Log "Status window unavailable: $($_.Exception.Message)" 'WARN'
    }
}

function Show-Status([string]$Text) {
    if (-not $script:StatusPS) { return }
    $WorkStatus.Text = $Text
    $WorkStatus.Started = Get-Date
    $WorkStatus.Visible = $true
}

function Hide-Status {
    $WorkStatus.Visible = $false
    if ($script:StatusPS) { Start-Sleep -Milliseconds 600 }   # let the window hide before a prompt opens
}

function Stop-StatusWindow {
    if (-not $script:StatusPS) { return }
    $WorkStatus.Close = $true
    [void]$script:StatusAsync.AsyncWaitHandle.WaitOne(3000)
    try { $script:StatusPS.Runspace.Close(); $script:StatusPS.Dispose() } catch { }
    $script:StatusPS = $null
}

function Complete-Run([int]$Code, [string]$Status) {
    Stop-StatusWindow
    if (-not $Status) {
        $Status = switch ($Code) {
            0       { 'Healthy' }
            1       { 'Failed' }
            2       { 'PreconditionFailed' }
            default { 'RebootRequired' }
        }
    }
    $Result.Status = $Status
    $Result | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $LogDir 'last-result.json') -Encoding UTF8
    Get-ChildItem $LogDir -Filter 'Repair-*.log' | Sort-Object LastWriteTime -Descending |
        Select-Object -Skip $KeepLogs | Remove-Item -Force -ErrorAction SilentlyContinue
    Write-Log "Finished: $Status (exit $Code)"

    if ($Interactive -and $Status -ne 'Cancelled') {
        $lines = $Result.Steps.GetEnumerator() | ForEach-Object { '{0}: {1}' -f $_.Key, $_.Value }
        $msg = "Result: $Status`n`n" + ($lines -join "`n")
        if ($Status -eq 'RebootRequired') { $msg += "`n`nPlease restart your computer to finish the repair." }
        $msg += "`n`nLog: $LogFile"
        Show-Dialog -Text $msg -Buttons OK -Timeout 300 | Out-Null
    }
    $script:ExitCode = $Code
    throw '__WindowsRepairExit__'
}

Write-Log "Start on $env:COMPUTERNAME, build $($Result.OSBuild), arch $env:PROCESSOR_ARCHITECTURE, user $([Security.Principal.WindowsIdentity]::GetCurrent().Name), interactive $Interactive"

# --- Pre-checks -------------------------------------------------------------
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Set-Step 'PreCheck' 'Not elevated: run as SYSTEM or administrator' -Fail
    Write-Log 'Not running elevated/SYSTEM.' 'ERROR'
    Complete-Run 2
}

$freeGB = [math]::Round((Get-PSDrive -Name $env:SystemDrive.TrimEnd(':')).Free / 1GB, 1)
Write-Log "Free space on ${env:SystemDrive}: $freeGB GB"
if ($freeGB -lt $MinFreeSpaceGB) {
    Set-Step 'PreCheck' "Free space $freeGB GB < $MinFreeSpaceGB GB" -Fail
    Write-Log 'Insufficient free space; DISM/SFC are unreliable below this threshold.' 'ERROR'
    Complete-Run 2
}

$cbsPending = Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
$wuPending  = Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
if ($wuPending) { Write-Log 'Windows Update reports a pending reboot.' 'WARN' }
if ($cbsPending -and -not $IgnorePendingReboot) {
    Set-Step 'PreCheck' 'Restart pending: restart the computer first, then run again' -Reboot
    Write-Log 'CBS reboot pending; servicing repairs will fail or be overwritten. Reboot first.' 'WARN'
    Complete-Run 2
}
Set-Step 'PreCheck' 'OK'

if (-not (Confirm-Step "This tool checks and repairs Windows system files.`nIt can take 30-60 minutes and may slow down your computer while it runs.`n`nStart now?" $true)) {
    Set-Step 'Start' 'Cancelled by user'
    Complete-Run 0 'Cancelled'
}
Start-StatusWindow


# --- Optional: reset Windows Update components ------------------------------
if (Confirm-Step "Reset Windows Update components?`nUse this when updates keep failing. The local update history will be cleared." ([bool]$ResetWindowsUpdate)) {
    Write-Log 'Resetting Windows Update components'
    Show-Status 'Resetting Windows Update components...'
    $svcs = 'wuauserv','bits','cryptsvc','usosvc'
    try {
        Stop-Service -Name $svcs -Force -ErrorAction SilentlyContinue
        foreach ($p in "$env:WINDIR\SoftwareDistribution", "$env:WINDIR\System32\catroot2") {
            if (Test-Path $p) { Rename-Item -Path $p -NewName ('{0}.bak-{1}' -f (Split-Path $p -Leaf), $Stamp) -ErrorAction Stop }
        }
        Set-Step 'WUReset' 'OK'
    } catch {
        Set-Step 'WUReset' "Failed: $($_.Exception.Message)" -Fail
        Write-Log "WU reset failed: $($_.Exception.Message)" 'ERROR'
    } finally {
        Start-Service -Name 'cryptsvc','bits','wuauserv' -ErrorAction SilentlyContinue
        Hide-Status
    }
}

# --- DISM -------------------------------------------------------------------
function Invoke-Dism([string]$Mode) {
    # RestoreHealth uses Windows Update as its repair source
    $p = @{ Online = $true; NoRestart = $true; ErrorAction = 'Stop'; $Mode = $true }
    $statusText = @{
        ScanHealth            = "Checking the Windows component store (DISM ScanHealth).`nThis can take 10-20 minutes."
        RestoreHealth         = "Repairing the Windows component store (DISM RestoreHealth).`nThis can take 15-30 minutes."
        StartComponentCleanup = "Cleaning up old Windows components (DISM).`nThis can take 5-15 minutes."
    }
    Write-Log "DISM $Mode started"
    Show-Status $statusText[$Mode]
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        $r = Repair-WindowsImage @p
        Write-Log ('DISM {0}: state={1}, restartNeeded={2} ({3:N0}s)' -f $Mode, $r.ImageHealthState, $r.RestartNeeded, $sw.Elapsed.TotalSeconds)
        if ($r.RestartNeeded) { $Result.RebootRequired = $true }
        return $r
    } catch {
        Write-Log ('DISM {0} failed: 0x{1:X8} {2}' -f $Mode, $_.Exception.HResult, $_.Exception.Message) 'ERROR'
        return $null
    } finally {
        Hide-Status
    }
}

if (Confirm-Step "Check the Windows component store for corruption?`n(DISM ScanHealth, about 10-20 minutes)" $true) {
    $scan  = Invoke-Dism 'ScanHealth'
    $state = if ($scan) { "$($scan.ImageHealthState)" } else { 'ScanFailed' }
} else {
    $state = 'Skipped'
}
Set-Step 'DismScan' $state

if ($state -eq 'NonRepairable') {
    Set-Step 'DismRestore' 'Skipped: store NonRepairable (in-place upgrade or reimage needed)' -Fail
} elseif ($state -eq 'Healthy' -and -not $ForceRestoreHealth) {
    # Nothing to repair
} elseif (Confirm-Step "Repair the Windows component store?`n(DISM RestoreHealth, about 15-30 minutes)`nScan result: $state" ($state -ne 'Skipped')) {
    $restore = Invoke-Dism 'RestoreHealth'
    if ($restore -and "$($restore.ImageHealthState)" -eq 'Healthy') {
        Set-Step 'DismRestore' 'Healthy'
    } else {
        # 0x800F081F / 0x800F0954 = repair content not found / Windows Update blocked by policy
        Set-Step 'DismRestore' 'Failed (see log; check Windows Update access)' -Fail
    }
} elseif ($state -in 'Repairable','ScanFailed') {
    Set-Step 'DismRestore' 'Declined while repair was needed' -Fail
} else {
    Set-Step 'DismRestore' 'Skipped'
}

if (-not $Result.Failed -and
    (Confirm-Step "Clean up old Windows components to free disk space?`n(DISM StartComponentCleanup, about 5-15 minutes)" ([bool]$ComponentCleanup))) {
    $c = Invoke-Dism 'StartComponentCleanup'
    if ($c) { Set-Step 'ComponentCleanup' 'OK' } else { Set-Step 'ComponentCleanup' 'Failed' -Fail }
}

# --- SFC --------------------------------------------------------------------
if (Confirm-Step "Check and repair protected Windows system files?`n(SFC /scannow, about 10-20 minutes)" $true) {
    Write-Log 'SFC /scannow started'
    $sfcStart = Get-Date
    Show-Status "Checking and repairing Windows system files (SFC).`nThis can take 10-20 minutes."
    $sfc = Start-Process -FilePath "$env:WINDIR\System32\sfc.exe" -ArgumentList '/scannow' -Wait -PassThru -WindowStyle Hidden
    Hide-Status
    Write-Log "SFC exit code $($sfc.ExitCode)"

    $cbsLog = "$env:WINDIR\Logs\CBS\CBS.log"
    $srLines = @()
    if (Test-Path $cbsLog) {
        $srLines = Get-Content $cbsLog -ErrorAction SilentlyContinue | Where-Object {
            if ($_ -notmatch '\[SR\]' -or $_.Length -lt 19) { return $false }
            $t = [datetime]::MinValue
            [datetime]::TryParseExact($_.Substring(0,19), 'yyyy-MM-dd HH:mm:ss',
                [Globalization.CultureInfo]::InvariantCulture, 'None', [ref]$t) -and $t -ge $sfcStart.AddSeconds(-5)
        }
        $srLines | Set-Content -Path (Join-Path $LogDir "SFC-SR-$Stamp.log")
    }
    $repaired   = @($srLines | Where-Object { $_ -match 'Repairing corrupted file' }).Count
    $unrepaired = @($srLines | Where-Object { $_ -match 'Cannot repair member file' }).Count
    Write-Log "SFC: $repaired repaired, $unrepaired not repairable ($(@($srLines).Count) [SR] lines)"

    if ($unrepaired -gt 0)   { Set-Step 'SFC' "$unrepaired file(s) not repairable" -Fail }
    elseif ($repaired -gt 0) { Set-Step 'SFC' "$repaired file(s) repaired" -Reboot }
    elseif (-not $srLines)   { Set-Step 'SFC' "No [SR] entries found (exit $($sfc.ExitCode)); verify manually" }
    else                     { Set-Step 'SFC' 'No integrity violations' }
} else {
    Set-Step 'SFC' 'Skipped'
}

# --- chkdsk online scan -----------------------------------------------------
if (Confirm-Step "Scan the system drive for file system errors?`n(chkdsk /scan, runs online, no restart)" $true) {
    Write-Log "chkdsk $env:SystemDrive /scan started"
    Show-Status "Scanning the system drive for file system errors (chkdsk).`nThis can take a few minutes."
    & "$env:WINDIR\System32\chkdsk.exe" $env:SystemDrive /scan | Out-Null
    $ck = $LASTEXITCODE
    Hide-Status
    Write-Log "chkdsk exit code $ck"
    switch ($ck) {
        0 { Set-Step 'Chkdsk' 'No errors' }
        1 { Set-Step 'Chkdsk' 'Errors found and fixed online' }
        2 { Set-Step 'Chkdsk' 'Cleanup performed' }
        default {
            if (Confirm-Step "The disk scan found errors that can only be fixed during a restart.`nSchedule the repair for the next restart? That restart will take longer than usual." ([bool]$ScheduleChkdskFix)) {
                & fsutil.exe dirty set $env:SystemDrive | Out-Null
                Set-Step 'Chkdsk' "Exit $ck; offline repair scheduled at next restart" -Reboot
            } else {
                Set-Step 'Chkdsk' "Exit $ck; offline repair needed but not scheduled" -Fail
            }
        }
    }
} else {
    Set-Step 'Chkdsk' 'Skipped'
}

# --- WMI (verification is read-only and quick, so it is not prompted) -------
& "$env:WINDIR\System32\wbem\winmgmt.exe" /verifyrepository | Out-Null
$wmiRc = $LASTEXITCODE
$wmiOk = $true
try { Get-CimInstance Win32_OperatingSystem -ErrorAction Stop | Out-Null } catch { $wmiOk = $false }
Write-Log "WMI verify exit $wmiRc, CIM query ok: $wmiOk"

if ($wmiRc -eq 0 -and $wmiOk) {
    Set-Step 'WMI' 'Consistent'
} elseif (Confirm-Step "The WMI repository is damaged.`nRepair it? (winmgmt /salvagerepository) A restart is recommended afterwards." ([bool]$SalvageWmi)) {
    Show-Status 'Repairing the WMI repository...'
    & "$env:WINDIR\System32\wbem\winmgmt.exe" /salvagerepository | Out-Null
    $salvageRc = $LASTEXITCODE
    Hide-Status
    Write-Log "WMI salvage exit $salvageRc"
    if ($salvageRc -eq 0) { Set-Step 'WMI' 'Salvaged' -Reboot } else { Set-Step 'WMI' 'Salvage failed' -Fail }
} else {
    Set-Step 'WMI' "Inconsistent (exit $wmiRc); salvage not run" -Fail
}

# --- Result -----------------------------------------------------------------
if     ($Result.Failed)         { Complete-Run 1 }
elseif ($Result.RebootRequired) { Complete-Run $RebootExitCode }
else                            { Complete-Run 0 }

} catch {
    if ($_.Exception.Message -ne '__WindowsRepairExit__') { throw }
}

# Exit code only when run as a file (MDM/Task Scheduler); with irm | iex the console stays open
if ($PSCommandPath) { exit $ExitCode }
