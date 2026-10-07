<#
.SYNOPSIS
    Run from a jumpbox against a list of computers. On each PC it enables (Disabled -> Manual)
    and starts the Windows Update services, installs all applicable updates, re-checks, then
    sets the services it enabled back to Disabled. Does NOT reboot anything.

    Every computer you pass in gets exactly one row with Result = Complete or NotComplete.
    Complete means: no applicable updates left, no reboot pending, nothing failed, and the
    services were put back the way they were. Anything else is NotComplete with a Reason,
    and those names are written to NotComplete_<time>.txt so you can just re-run on them.

.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\Install-RemoteWindowsUpdates_v5.ps1
    # prompts you to paste computer names
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\Install-RemoteWindowsUpdates_v5.ps1 -ListPath .\pcs.txt
    # ...reboot the PCs that report "reboot pending", then re-run on the leftovers:
    powershell.exe -ExecutionPolicy Bypass -File .\Install-RemoteWindowsUpdates_v5.ps1 -ListPath .\NotComplete_20261007_140000.txt
    # repeat until it reports all complete
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\Install-RemoteWindowsUpdates_v5.ps1 -ListPath .\pcs.txt -Verify   # check only, installs nothing
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\Install-RemoteWindowsUpdates_v5.ps1 -ComputerName PC01 -UpdateSource WindowsUpdate   # bypass WSUS

.NOTES
    Version: 5 FINAL - 2026-10-07
    Requirements: WinRM/PowerShell remoting enabled on targets, local admin rights there.
    Why a scheduled task? The Windows Update Agent COM API refuses to download/install
    when called directly over a remote session (Access Denied). The script drops a
    temporary worker + one-shot SYSTEM scheduled task on each PC, polls it, then removes
    only the files/task it created.
    The original service startup types are saved on each PC in
    C:\Windows\Temp\WUServiceBackup.txt until they have been restored, so a run that times
    out still gets the services re-disabled by the next run.
#>
[CmdletBinding()]
param(
    [string[]]$ComputerName,
    [string]$ListPath,
    [pscredential]$Credential,
    [int]$ThrottleLimit = 32,
    [int]$TimeoutMinutes = 240,
    [int]$PollSeconds = 60,
    [Alias('ScanOnly')]
    [switch]$Verify,                    # check only: search + reboot check, install nothing
    [ValidateSet('Default','WSUS','WindowsUpdate')]
    [string]$UpdateSource = 'Default',  # Default = whatever the PC is configured for (GPO/WSUS)
    [string]$SearchCriteria = "IsInstalled=0 and IsHidden=0 and Type='Software'",
    [int]$MaxPasses = 3,                # re-scan + install again for updates that only appear after others
    [switch]$SkipInteractive,           # don't attempt updates flagged as possibly needing user input
    [int]$InteractiveTimeoutMinutes = 30,  # per-update limit when attempting those silently
    [switch]$LeaveServicesEnabled,      # don't put Disabled services back to Disabled afterward
    [switch]$Force                      # skip the "proceed?" confirmation
)

$ErrorActionPreference = 'Stop'
$scriptVersion = '5 FINAL - 2026-10-07'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
# wuauserv is required; the rest are best-effort (a failure on them is reported, not fatal)
$serviceNames = 'wuauserv','bits','cryptsvc','UsoSvc','DoSvc'
$serverSelection = @{ Default = 0; WSUS = 1; WindowsUpdate = 2 }[$UpdateSource]

# ---------------------------------------------------------------- gather targets
$targets = @()
if ($ComputerName) { $targets += $ComputerName }
if ($ListPath)     { $targets += Get-Content -Path $ListPath }
if (-not $targets) {
    Write-Host "Paste computer names (one per line, or comma/space separated). Blank line to finish:" -ForegroundColor Cyan
    while ($true) {
        $line = Read-Host
        if ([string]::IsNullOrWhiteSpace($line)) { break }
        $targets += $line
    }
}
$targets = @($targets | ForEach-Object { $_ -split '[,;\s]+' } |
           Where-Object { $_ } | ForEach-Object { $_.Trim() } | Sort-Object -Unique)
if (-not $targets) { throw "No computer names supplied." }

$mode = if ($Verify) { 'VERIFY ONLY (search + reboot check, no download/install)' } else { 'DOWNLOAD + INSTALL updates (no reboot)' }
$preview = ($targets | Select-Object -First 10) -join ', '
if ($targets.Count -gt 10) { $preview += ", ... (+$($targets.Count - 10) more)" }
Write-Host "`nInstall-RemoteWindowsUpdates version $scriptVersion" -ForegroundColor Green
Write-Host "Scope of this run:" -ForegroundColor Yellow
Write-Host "  Action : $mode"
Write-Host "  Source : $UpdateSource   (up to $MaxPasses install passes, stops early if a reboot is needed)"
if (-not $Verify) {
    $ia = if ($SkipInteractive) { 'skipped (PC reported NotComplete)' } else { "attempted silently, one at a time, $InteractiveTimeoutMinutes min limit each" }
    Write-Host "  Updates that may need user input: $ia"
}
Write-Host "  Targets: $($targets.Count) -> $preview"
Write-Host "  Changes: set $($serviceNames -join '/') to Manual if Disabled, and start them;"
Write-Host "           create a temp worker script + one-shot SYSTEM scheduled task in C:\Windows\Temp"
Write-Host "           Set them back to Disabled afterward: $(-not $LeaveServicesEnabled)"
if (-not $Force) {
    if ((Read-Host "Proceed? (y/N)") -notmatch '^(y|yes)$') { Write-Host "Aborted."; return }
}

$cred = @{}
if ($Credential) { $cred.Credential = $Credential }

$params = @{
    Verify                    = [bool]$Verify
    ServerSelection           = $serverSelection
    MaxPasses                 = $MaxPasses
    Criteria                  = $SearchCriteria
    SkipInteractive           = [bool]$SkipInteractive
    InteractiveTimeoutMinutes = $InteractiveTimeoutMinutes
    TaskLimitMinutes          = $TimeoutMinutes + 120
}

# ---------------------------------------------------------------- shared: restore services (runs on target)
$restoreText = @'
function Restore-WUServices([string]$BackupFile) {
    $detail = @(); $ok = $true
    if (Test-Path $BackupFile) {
        foreach ($line in Get-Content -Path $BackupFile) {
            $n, $m = $line -split '\|'
            if ($m -ne 'Disabled') { continue }
            Stop-Service -Name $n -Force -ErrorAction SilentlyContinue
            & sc.exe config $n start= disabled | Out-Null
            $w = Get-CimInstance -ClassName Win32_Service -Filter "Name='$n'" -ErrorAction SilentlyContinue
            if (-not $w)                     { continue }
            if ($w.StartMode -eq 'Disabled') { $detail += "$n -> Disabled/$($w.State)" }
            else                             { $ok = $false; $detail += "RESTORE FAILED: $n is $($w.StartMode)/$($w.State)" }
        }
        # keep the backup until everything is really back, so the next run retries
        if ($ok) { Remove-Item -Path $BackupFile -Force -ErrorAction SilentlyContinue }
    }
    [pscustomobject]@{ Ok = $ok; Detail = $detail }
}
'@

# ---------------------------------------------------------------- worker (runs on target as SYSTEM)
$workerText = @'
param([string]$Stamp)
$base    = "$env:windir\Temp"
$OutFile = "$base\WUResult_$Stamp.json"
$cfg     = Get-Content -Path "$base\WUParams_$Stamp.json" -Raw | ConvertFrom-Json
$out = [ordered]@{
    Status = 'Started'; Mode = $(if ($cfg.Verify) { 'Verify' } else { 'Install' }); WuauservAtRun = $null; Source = $null
    Passes = 0; Found = 0; Available = @(); Installed = @(); Failed = @(); Skipped = @()
    Remaining = $null; RemainingTitles = @(); RebootRequired = $false; RebootReasons = @(); RebootPendingBefore = $false; Error = $null
}
function HexCode($h) { '0x{0:X8}' -f ([long]$h -band 4294967295) }
function Get-RebootReasons {
    $r = @()
    if ((New-Object -ComObject Microsoft.Update.SystemInfo).RebootRequired) { $r += 'Windows Update' }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $r += 'Component Based Servicing' }
    $r
}
# Installs ONE update in a separate process so it can be killed if it sits waiting for input
$childText = @"
`$r = @{ Code = -1; HResult = 0; Reboot = `$false; Error = `$null }
try {
    `$s = New-Object -ComObject Microsoft.Update.Session
    `$s.ClientApplicationID = 'Install-RemoteWindowsUpdates'
    `$q = `$s.CreateUpdateSearcher()
    `$q.ServerSelection = __SEL__
    `$found = `$q.Search("UpdateID='__ID__'").Updates
    if (`$found.Count -eq 0) { throw 'update is no longer offered' }
    `$c = New-Object -ComObject Microsoft.Update.UpdateColl
    `$u = `$found.Item(0)
    if (-not `$u.EulaAccepted) { `$u.AcceptEula() }
    [void]`$c.Add(`$u)
    `$d = `$s.CreateUpdateDownloader(); `$d.Updates = `$c; [void]`$d.Download()
    if (-not `$u.IsDownloaded) { throw 'download failed' }
    `$i = `$s.CreateUpdateInstaller(); `$i.Updates = `$c
    `$i.AllowSourcePrompts = `$false; `$i.ForceQuiet = `$true
    `$ir = `$i.Install()
    `$r.Code = `$ir.GetUpdateResult(0).ResultCode
    `$r.HResult = `$ir.GetUpdateResult(0).HResult
    `$r.Reboot = [bool]`$ir.RebootRequired
} catch { `$r.Error = `$_.Exception.Message; `$r.HResult = `$_.Exception.HResult }
`$r | ConvertTo-Json | Set-Content -Path '__OUT__' -Encoding UTF8
"@

try {
    # proves the service was really enabled when the update ran (and catches GPO re-disabling it)
    $w = Get-CimInstance -ClassName Win32_Service -Filter "Name='wuauserv'"
    $out.WuauservAtRun = "$($w.StartMode)/$($w.State)"
    if ($w.StartMode -eq 'Disabled') { throw "wuauserv was Disabled again before the update task ran (GPO or another tool re-disabling it?)" }
    $out.RebootPendingBefore = @(Get-RebootReasons).Count -gt 0

    $session  = New-Object -ComObject Microsoft.Update.Session
    $session.ClientApplicationID = 'Install-RemoteWindowsUpdates'
    $searcher = $session.CreateUpdateSearcher()
    $searcher.ServerSelection = [int]$cfg.ServerSelection
    $out.Source = @('Default (machine config)', 'WSUS', 'Windows Update')[[int]$cfg.ServerSelection]
    $tried       = @{}   # UpdateID -> $true, so a failing update isn't retried every pass
    $interactive = New-Object System.Collections.ArrayList
    $first       = $null

    for ($pass = 1; $pass -le [int]$cfg.MaxPasses; $pass++) {
        $sr = $searcher.Search($cfg.Criteria)
        if ($pass -eq 1) {
            $first = $sr
            $out.Found = $sr.Updates.Count
            foreach ($u in $sr.Updates) { $out.Available += $u.Title }
        }
        if ($cfg.Verify) { break }

        $coll = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($u in $sr.Updates) {
            $id = $u.Identity.UpdateID
            if ($tried.ContainsKey($id)) { continue }
            $tried[$id] = $true
            # done separately below, after everything else
            if ($u.InstallationBehavior.CanRequestUserInput) { [void]$interactive.Add($u); continue }
            if (-not $u.EulaAccepted) { $u.AcceptEula() }
            [void]$coll.Add($u)
        }
        if ($coll.Count -eq 0) { break }
        $out.Passes = $pass

        $dl = $session.CreateUpdateDownloader()
        $dl.Updates = $coll
        $dr = $dl.Download()
        $ready = New-Object -ComObject Microsoft.Update.UpdateColl
        for ($i = 0; $i -lt $coll.Count; $i++) {
            $u = $coll.Item($i)
            if ($u.IsDownloaded) { [void]$ready.Add($u) }
            else { $out.Failed += "$($u.Title) [download failed $(HexCode $dr.GetUpdateResult($i).HResult)]" }
        }
        if ($ready.Count -eq 0) { break }

        $inst = $session.CreateUpdateInstaller()
        $inst.Updates = $ready
        $inst.AllowSourcePrompts = $false
        $inst.ForceQuiet = $true
        $ir = $inst.Install()
        for ($i = 0; $i -lt $ready.Count; $i++) {
            $ur = $ir.GetUpdateResult($i)
            $t  = $ready.Item($i).Title
            if ($ur.ResultCode -eq 2 -or $ur.ResultCode -eq 3) { $out.Installed += $t }
            else { $out.Failed += "$t [install result $($ur.ResultCode), $(HexCode $ur.HResult)]" }
        }
        # anything still waiting (e.g. the cumulative update after a servicing stack update) needs the reboot first
        if ($ir.RebootRequired) { $out.RebootRequired = $true; break }
    }

    # updates flagged "may request user input": nobody is logged on to click, so try each one
    # silently in its own process with a time limit; a hang is killed and reported, never left silent
    foreach ($u in $interactive) {
        if ($cfg.SkipInteractive) { $out.Skipped += "$($u.Title) (may need user input; -SkipInteractive)"; continue }
        if ($out.RebootRequired)  { $out.Skipped += "$($u.Title) (deferred: reboot needed first)"; continue }
        $res  = "$base\WUChild_$([guid]::NewGuid().ToString('N')).json"
        $code = $childText.Replace('__SEL__', "$([int]$cfg.ServerSelection)").Replace('__ID__', $u.Identity.UpdateID).Replace('__OUT__', $res)
        $enc  = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($code))
        $p = Start-Process -FilePath powershell.exe -ArgumentList "-NoProfile -NonInteractive -EncodedCommand $enc" -WindowStyle Hidden -PassThru
        if (-not $p.WaitForExit([int]$cfg.InteractiveTimeoutMinutes * 60000)) {
            try { $p.Kill() } catch { }
            $out.Failed += "$($u.Title) [no finish after $($cfg.InteractiveTimeoutMinutes) min - likely waiting for user input; install it by hand]"
        } elseif (Test-Path $res) {
            $cr = Get-Content -Path $res -Raw | ConvertFrom-Json
            if ($cr.Code -eq 2 -or $cr.Code -eq 3) { $out.Installed += $u.Title; if ($cr.Reboot) { $out.RebootRequired = $true } }
            else { $out.Failed += "$($u.Title) [install result $($cr.Code), $(HexCode $cr.HResult) $($cr.Error)]" }
        } else {
            $out.Failed += "$($u.Title) [install process ended without a result]"
        }
        Remove-Item -Path $res -ErrorAction SilentlyContinue
    }

    # independent re-check (in Verify mode the first search already is that check)
    $rs = if ($cfg.Verify) { $first } else { $searcher.Search($cfg.Criteria) }
    $out.Remaining = $rs.Updates.Count
    for ($i = 0; $i -lt [Math]::Min($rs.Updates.Count, 15); $i++) { $out.RemainingTitles += $rs.Updates.Item($i).Title }
    $out.RebootReasons  = @(Get-RebootReasons)
    $out.RebootRequired = $out.RebootRequired -or $out.RebootReasons.Count -gt 0
    $out.Status = 'Completed'
} catch {
    $out.Status = 'Error'
    $out.Error  = "$($_.Exception.Message) ($(HexCode $_.Exception.HResult))"
}
# write then rename, so the poller never reads a half-written file
$out | ConvertTo-Json -Depth 4 | Set-Content -Path "$OutFile.tmp" -Encoding UTF8
Move-Item -Path "$OutFile.tmp" -Destination $OutFile -Force
'@

# ---------------------------------------------------------------- phase 1: prep + launch
Write-Host "`n[1/3] Enabling services and launching update task on $($targets.Count) machine(s)..." -ForegroundColor Cyan

$p1Err = $null
$p1 = Invoke-Command @cred -ComputerName $targets -ThrottleLimit $ThrottleLimit `
    -ArgumentList $workerText, $restoreText, $stamp, $serviceNames, $params `
    -ErrorAction SilentlyContinue -ErrorVariable p1Err -ScriptBlock {
    param($WorkerText, $RestoreText, $Stamp, $ServiceNames, $Params)
    # $ErrorActionPreference from the jumpbox does NOT carry into the remote session
    $ErrorActionPreference = 'Stop'
    . ([scriptblock]::Create($RestoreText))

    $base       = "$env:windir\Temp"
    $backupFile = "$base\WUServiceBackup.txt"
    $r = [ordered]@{ Ok = $false; Busy = $false; Error = $null; Services = @(); Notes = @(); Restore = $null }
    # Ask the Service Control Manager directly (Get-Service StartType is missing on older PS)
    function Get-SvcInfo([string]$Name) {
        Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue
    }
    try {
        $running = @(Get-ScheduledTask -TaskName 'WUInstall_*' -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Running' })
        if ($running) { $r.Busy = $true; throw "update task $($running[0].TaskName) from an earlier run is still running" }

        # leftovers from earlier runs of this script (timed out, etc.) - nothing of them is still running
        Get-ScheduledTask -TaskName 'WUInstall_*' -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
        Remove-Item -Path "$base\WUWorker_*.ps1", "$base\WUResult_*.json*", "$base\WUParams_*.json", "$base\WUChild_*.json" -ErrorAction SilentlyContinue

        # Save each service's ORIGINAL startup type on the PC itself (merged with any earlier
        # unrestored backup), so whichever run finishes last still knows to put Disabled back
        $backup = [ordered]@{}
        if (Test-Path $backupFile) {
            foreach ($line in Get-Content -Path $backupFile) { $n, $m = $line -split '\|'; if ($n) { $backup[$n] = $m } }
        }
        foreach ($n in $ServiceNames) {
            $s = Get-SvcInfo $n
            if ($s -and (-not $backup.Contains($n) -or $s.StartMode -eq 'Disabled')) { $backup[$n] = $s.StartMode }
        }
        $backup.GetEnumerator() | ForEach-Object { "$($_.Key)|$($_.Value)" } | Set-Content -Path $backupFile -Encoding ASCII

        foreach ($n in $ServiceNames) {
            $before = Get-SvcInfo $n
            if (-not $before) { continue }
            try {
                if ($before.StartMode -eq 'Disabled') {
                    $scOut = & sc.exe config $n start= demand   # sc.exe reports errors on stdout
                    if ($LASTEXITCODE -ne 0) { throw "sc.exe config failed (exit $LASTEXITCODE): $($scOut -join ' ')" }
                    if ((Get-SvcInfo $n).StartMode -eq 'Disabled') { throw "startup type is still Disabled after sc.exe config" }
                }
                if ((Get-Service -Name $n).Status -ne 'Running') {
                    Start-Service -Name $n
                    (Get-Service -Name $n).WaitForStatus('Running', [TimeSpan]::FromSeconds(60))
                }
                $after = Get-SvcInfo $n
                $r.Services += "$n $($before.StartMode)/$($before.State) -> $($after.StartMode)/$($after.State)"
            } catch {
                if ($n -eq 'wuauserv') { throw "wuauserv could not be enabled/started: $($_.Exception.Message)" }
                $r.Notes += "$n not enabled/started: $($_.Exception.Message)"
            }
        }
        if ((Get-Service -Name wuauserv).Status -ne 'Running') { throw "wuauserv is not running" }

        $wuBusy = $false
        try { $wuBusy = [bool](New-Object -ComObject Microsoft.Update.Installer).IsBusy } catch { }
        if ($wuBusy) { $r.Busy = $true; throw "another Windows Update installation is already in progress" }

        # Not changed by this script - reported because they commonly cause 'found 0' or search errors
        $pol = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -ErrorAction SilentlyContinue
        $au  = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' -ErrorAction SilentlyContinue
        if ($au.UseWUServer -eq 1)                                   { $r.Notes += "Policy: WSUS server $($pol.WUServer)" }
        if ($pol.DisableWindowsUpdateAccess -eq 1)                   { $r.Notes += "Policy: DisableWindowsUpdateAccess=1" }
        if ($pol.DoNotConnectToWindowsUpdateInternetLocations -eq 1) { $r.Notes += "Policy: DoNotConnectToWindowsUpdateInternetLocations=1" }

        $worker = "$base\WUWorker_$Stamp.ps1"
        Set-Content -Path $worker -Value $WorkerText -Encoding UTF8
        $Params | ConvertTo-Json | Set-Content -Path "$base\WUParams_$Stamp.json" -Encoding UTF8

        # -File is blocked when GPO enforces an execution policy (the worker silently never runs);
        # a scriptblock passed via -Command is not subject to execution policy
        $cmd       = "& ([scriptblock]::Create([IO.File]::ReadAllText('$worker'))) -Stamp '$Stamp'"
        $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -NonInteractive -Command `"$cmd`""
        $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
        $settings  = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes $Params.TaskLimitMinutes) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
        [void](Register-ScheduledTask -TaskName "WUInstall_$Stamp" -Action $action -Principal $principal -Settings $settings)
        Start-ScheduledTask -TaskName "WUInstall_$Stamp"
        $r.Ok = $true
    } catch {
        $r.Error = $_.Exception.Message
        # launch failed - put back anything we enabled so the machine isn't left half-changed
        # (not when busy: something of ours or Windows Update itself is still working)
        if (-not $r.Busy) { $r.Restore = Restore-WUServices $backupFile }
    }
    [pscustomobject]$r
}

$state   = @{}   # computer -> launch info (hashtables are case-insensitive, like computer names)
$results = @{}   # computer -> final result row
$errText = @{}   # computer -> connection error text
function New-Row {
    param($Computer, $Result = 'NotComplete', $Reason = '', $Status, $Found = '', $Installed = '', $Failed = '',
          $Remaining = '', $RebootRequired = '', $Services = '', $Detail = '')
    [pscustomobject]@{ Computer=$Computer; Result=$Result; Reason=$Reason; Status=$Status; Found=$Found; Installed=$Installed
                       Failed=$Failed; Remaining=$Remaining; RebootRequired=$RebootRequired; Services=$Services; Detail=$Detail }
}
function Join-Text { (@($args | ForEach-Object { $_ }) | Where-Object { $_ }) -join '; ' }

foreach ($item in $p1) {
    $c = $item.PSComputerName
    if ($item.Ok) { $state[$c] = $item; continue }
    $svc = Join-Text $item.Services $item.Restore.Detail
    if ($item.Busy) {
        $results[$c] = New-Row -Computer $c -Status 'Busy' -Services $svc -Detail (Join-Text $item.Notes) `
            -Reason "busy: $($item.Error) - re-run later"
    } else {
        $why = "could not start: $($item.Error)"
        if ($item.Restore -and -not $item.Restore.Ok) { $why += '; services not re-disabled' }
        $results[$c] = New-Row -Computer $c -Status 'LaunchFailed' -Reason $why -Services $svc -Detail (Join-Text $item.Notes)
    }
}
foreach ($e in $p1Err) {
    $name = @($e.OriginInfo.PSComputerName, $e.CategoryInfo.TargetName, $(if ($e.TargetObject -is [string]) { $e.TargetObject })) |
            Where-Object { $_ } | Select-Object -First 1
    if ($name -and -not $errText.ContainsKey("$name")) { $errText["$name"] = $e.Exception.Message }
}
Write-Host "      Launched on $($state.Count); busy/failed: $($results.Count); no connection: $($targets.Count - $state.Count - $results.Count)"

# ---------------------------------------------------------------- phase 2: poll
$finished = @{}  # computer -> parsed json (or $null)
$taskExit = @{}  # computer -> scheduled task LastTaskResult when no result file was written
$pending  = New-Object 'System.Collections.Generic.List[string]'
foreach ($k in $state.Keys) { $pending.Add([string]$k) }
if ($pending.Count -gt 0) {
    Write-Host "`n[2/3] Waiting for updates to finish (timeout $TimeoutMinutes min)..." -ForegroundColor Cyan
}
$deadline = (Get-Date).AddMinutes($TimeoutMinutes)

while ($pending.Count -gt 0 -and (Get-Date) -lt $deadline) {
    Start-Sleep -Seconds $PollSeconds
    $poll = Invoke-Command @cred -ComputerName $pending.ToArray() -ThrottleLimit $ThrottleLimit `
        -ArgumentList $stamp -ErrorAction SilentlyContinue -ScriptBlock {
        param($Stamp)
        $out = "$env:windir\Temp\WUResult_$Stamp.json"
        if (Test-Path $out) {
            [pscustomobject]@{ Done = $true; Json = (Get-Content $out -Raw); LastTaskResult = $null }
        } else {
            $t = Get-ScheduledTask -TaskName "WUInstall_$Stamp" -ErrorAction SilentlyContinue
            $running = $t -and $t.State -in 'Running', 'Queued'
            $last = (Get-ScheduledTaskInfo -TaskName "WUInstall_$Stamp" -ErrorAction SilentlyContinue).LastTaskResult
            [pscustomobject]@{ Done = (-not $running); Json = $null; LastTaskResult = $last }
        }
    }
    foreach ($p in $poll) {
        if (-not $p.Done) { continue }
        $c = $p.PSComputerName
        $j = $null
        if ($p.Json) { try { $j = $p.Json | ConvertFrom-Json } catch { } }
        if (-not $j) { $taskExit[$c] = $p.LastTaskResult }
        $finished[$c] = $j
        [void]$pending.Remove($c)
        $st = if ($j) { $j.Status } else { 'NoResult' }
        Write-Host "      done: $c -> $st  ($($pending.Count) remaining)"
    }
}

foreach ($c in $pending) {
    $results[$c] = New-Row -Computer $c -Status 'TimedOut' -Services (Join-Text $state[$c].Services) -Detail (Join-Text $state[$c].Notes) `
        -Reason "no result within $TimeoutMinutes min (may still be installing, or lost contact); services left enabled - re-run later"
}

# ---------------------------------------------------------------- phase 3: cleanup + restore
$done    = @($finished.Keys)
$restore = @{}  # computer -> restore result
if ($done.Count -gt 0) {
    Write-Host "`n[3/3] Cleaning up temp task/files and restoring service state..." -ForegroundColor Cyan
    $p3 = Invoke-Command @cred -ComputerName $done -ThrottleLimit $ThrottleLimit -ErrorAction SilentlyContinue `
        -ArgumentList $stamp, $restoreText, (-not $LeaveServicesEnabled) -ScriptBlock {
        param($Stamp, $RestoreText, $DoRestore)
        . ([scriptblock]::Create($RestoreText))
        $base = "$env:windir\Temp"
        # remove only what this script created
        Unregister-ScheduledTask -TaskName "WUInstall_$Stamp" -Confirm:$false -ErrorAction SilentlyContinue
        Remove-Item -Path "$base\WUWorker_$Stamp.ps1", "$base\WUResult_$Stamp.json", "$base\WUParams_$Stamp.json" -ErrorAction SilentlyContinue
        if ($DoRestore) { Restore-WUServices "$base\WUServiceBackup.txt" }
        else            { [pscustomobject]@{ Ok = $true; Detail = @('services left enabled (-LeaveServicesEnabled)') } }
    }
    foreach ($x in $p3) { $restore[$x.PSComputerName] = $x }
}

foreach ($c in $done) {
    $j  = $finished[$c]
    $rs = $restore[$c]
    $why = @()
    if (-not $rs)       { $why += 'could not reach PC afterward to re-disable services' }
    elseif (-not $rs.Ok) { $why += 'services not re-disabled' }
    $svc = Join-Text $state[$c].Services $rs.Detail

    if ($null -eq $j) {
        $why = @("update task ended without a result (task exit code 0x{0:X8})" -f ([long]$taskExit[$c] -band 4294967295)) + $why
        $results[$c] = New-Row -Computer $c -Status 'NoResult' -Reason ($why -join '; ') -Services $svc -Detail (Join-Text $state[$c].Notes)
        continue
    }

    $nFailed = @($j.Failed).Count
    if ($j.Status -ne 'Completed') {
        $why = @("error: $($j.Error)") + $why
    } else {
        $pre = @()
        if ($nFailed)               { $pre += "$nFailed update(s) failed" }
        if ($j.Remaining -gt 0)     { $pre += "$($j.Remaining) update(s) still available" }
        if ($j.RebootRequired)      { $pre += "reboot pending ($(@($j.RebootReasons) -join ', '))" }
        $why = $pre + $why
    }

    $info = @("wuauserv during run: $($j.WuauservAtRun)", "source: $($j.Source)", "passes: $($j.Passes)")
    if ($j.RebootPendingBefore) { $info += 'a reboot was already pending before this run' }
    $results[$c] = New-Row -Computer $c `
        -Result $(if ($why) { 'NotComplete' } else { 'Complete' }) -Reason ($why -join '; ') -Status $j.Mode `
        -Found $j.Found -Installed @($j.Installed).Count -Failed $nFailed -Remaining $j.Remaining `
        -RebootRequired $j.RebootRequired -Services $svc `
        -Detail (Join-Text $info `
                    @($j.Installed       | ForEach-Object { "INSTALLED: $_" }) `
                    @($j.Failed          | ForEach-Object { "FAILED: $_" }) `
                    @($j.Skipped         | ForEach-Object { "SKIPPED: $_" }) `
                    @($j.RemainingTitles | ForEach-Object { "STILL AVAILABLE: $_" }) `
                    $j.Error $state[$c].Notes)
}

# every name that went in gets exactly one row, even if nothing came back from it
foreach ($t in $targets) {
    if (-not $results.ContainsKey($t)) {
        $msg = if ($errText.ContainsKey($t)) { $errText[$t] } else { 'no response' }
        $results[$t] = New-Row -Computer $t -Status 'Unreachable' -Reason "unreachable: $msg"
    }
}

# ---------------------------------------------------------------- report
$report = @($targets | ForEach-Object { $results[$_] })
$csv = Join-Path (Get-Location) "WUReport_$stamp.csv"
$report | Export-Csv -Path $csv -NoTypeInformation

$notDone = @($report | Where-Object { $_.Result -ne 'Complete' })
Write-Host "`n===== Summary =====" -ForegroundColor Green
$color = if ($notDone) { 'Yellow' } else { 'Green' }
Write-Host "Complete: $($report.Count - $notDone.Count) of $($report.Count)    NotComplete: $($notDone.Count)" -ForegroundColor $color
Write-Host "Full report: $csv"

if ($notDone) {
    $notDone | Format-Table Computer, Status, Reason -AutoSize -Wrap | Out-Host
    $retry = Join-Path (Get-Location) "NotComplete_$stamp.txt"
    $notDone | ForEach-Object { $_.Computer } | Set-Content -Path $retry
    Write-Host "NotComplete list: $retry" -ForegroundColor Yellow
    if ($notDone | Where-Object { $_.RebootRequired -eq $true }) {
        Write-Host "Some need a reboot (NOT rebooted - schedule it yourself). After rebooting, re-run:" -ForegroundColor Yellow
    } else {
        Write-Host "Fix or wait out the reasons above, then re-run:" -ForegroundColor Yellow
    }
    Write-Host "  .\$(Split-Path -Leaf $PSCommandPath) -ListPath `"$retry`"" -ForegroundColor Yellow
} else {
    Write-Host "All $($report.Count) computer(s) are fully updated, with no reboot pending." -ForegroundColor Green
}
