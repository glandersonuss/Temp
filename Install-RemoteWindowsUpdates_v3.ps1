<#
.SYNOPSIS
    Run from a jumpbox. Takes many computer names, enables (Disabled -> Manual) and starts
    the Windows Update services on each, downloads + installs updates, then sets the ones
    it enabled back to Disabled. Does NOT reboot anything - it only reports which machines need one.

.EXAMPLE
    .\Install-RemoteWindowsUpdates.ps1                       # prompts you to paste names
.EXAMPLE
    .\Install-RemoteWindowsUpdates.ps1 -ListPath .\pcs.txt -Credential (Get-Credential)
.EXAMPLE
    .\Install-RemoteWindowsUpdates.ps1 -ComputerName PC01,PC02 -ScanOnly

.NOTES
    Requirements: WinRM/PowerShell remoting enabled on targets, local admin rights there.
    Why a scheduled task? The Windows Update Agent COM API refuses to download/install
    when called directly over a remote session (Access Denied). The script drops a
    temporary worker + one-shot SYSTEM scheduled task on each PC, polls it, then removes
    only the files/task it created.
#>
[CmdletBinding()]
param(
    [string[]]$ComputerName,
    [string]$ListPath,
    [pscredential]$Credential,
    [int]$ThrottleLimit = 16,
    [int]$TimeoutMinutes = 180,
    [int]$PollSeconds = 60,
    [switch]$ScanOnly,              # search only, don't download/install
    [switch]$LeaveServicesEnabled,  # don't put Disabled services back to Disabled afterward
    [switch]$Force                  # skip the "proceed?" confirmation
)

$ErrorActionPreference = 'Stop'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
# wuauserv is required; the rest are best-effort (a failure on them is reported, not fatal)
$serviceNames = 'wuauserv','bits','cryptsvc','UsoSvc','DoSvc'

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
$targets = $targets | ForEach-Object { $_ -split '[,;\s]+' } |
           Where-Object { $_ } | ForEach-Object { $_.Trim() } | Sort-Object -Unique
if (-not $targets) { throw "No computer names supplied." }

$mode = if ($ScanOnly) { 'SCAN ONLY (no download/install)' } else { 'DOWNLOAD + INSTALL updates (no reboot)' }
Write-Host "`nScope of this run:" -ForegroundColor Yellow
Write-Host "  Action : $mode"
Write-Host "  Targets: $($targets.Count) -> $($targets -join ', ')"
Write-Host "  Changes: set $($serviceNames -join '/') to Manual if Disabled, and start them;"
Write-Host "           create a temp worker script + one-shot SYSTEM scheduled task in C:\Windows\Temp"
Write-Host "           Set them back to Disabled afterward: $(-not $LeaveServicesEnabled)"
if (-not $Force) {
    if ((Read-Host "Proceed? (y/N)") -notmatch '^(y|yes)$') { Write-Host "Aborted."; return }
}

$cred = @{}
if ($Credential) { $cred.Credential = $Credential }

# ---------------------------------------------------------------- worker (runs on target as SYSTEM)
$workerText = @'
param([string]$OutFile, [switch]$ScanOnly)
$out = [ordered]@{ Status='Started'; WuauservAtRun=$null; Found=0; Installed=@(); Failed=@(); RebootRequired=$false; Error=$null }
try {
    # proves the service was really enabled when the update ran (and catches GPO re-disabling it)
    $w = Get-CimInstance -ClassName Win32_Service -Filter "Name='wuauserv'"
    $out.WuauservAtRun = "$($w.StartMode)/$($w.State)"
    if ($w.StartMode -eq 'Disabled') { throw "wuauserv was Disabled again before the update task ran (GPO or another tool re-disabling it?)" }

    $session  = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $sr = $searcher.Search("IsInstalled=0 and IsHidden=0 and Type='Software'")
    $out.Found = $sr.Updates.Count
    if ($sr.Updates.Count -gt 0 -and -not $ScanOnly) {
        $coll = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($u in $sr.Updates) {
            if (-not $u.EulaAccepted) { $u.AcceptEula() }
            [void]$coll.Add($u)
        }
        $dl = $session.CreateUpdateDownloader()
        $dl.Updates = $coll
        [void]$dl.Download()

        $ready = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($u in $coll) { if ($u.IsDownloaded) { [void]$ready.Add($u) } }

        if ($ready.Count -gt 0) {
            $inst = $session.CreateUpdateInstaller()
            $inst.Updates = $ready
            $inst.AllowSourcePrompts = $false
            $ir = $inst.Install()
            for ($i = 0; $i -lt $ready.Count; $i++) {
                $rc = $ir.GetUpdateResult($i).ResultCode
                $t  = "$($ready.Item($i).Title) [code $rc]"
                if ($rc -eq 2 -or $rc -eq 3) { $out.Installed += $t } else { $out.Failed += $t }
            }
            $out.RebootRequired = [bool]$ir.RebootRequired
        }
        $notDl = $coll.Count - $ready.Count
        if ($notDl -gt 0) { $out.Failed += "$notDl update(s) failed to download" }
    }
    $out.Status = 'Completed'
} catch {
    $out.Status = 'Error'
    $out.Error  = $_.Exception.Message
}
$out | ConvertTo-Json -Depth 4 | Set-Content -Path $OutFile -Encoding UTF8
'@

# ---------------------------------------------------------------- phase 1: prep + launch
Write-Host "`n[1/3] Enabling services and launching update task on $($targets.Count) machine(s)..." -ForegroundColor Cyan

$p1Err = $null
$p1 = Invoke-Command @cred -ComputerName $targets -ThrottleLimit $ThrottleLimit `
    -ArgumentList $workerText, $ScanOnly.IsPresent, $stamp, $serviceNames `
    -ErrorAction SilentlyContinue -ErrorVariable p1Err -ScriptBlock {
    param($WorkerText, $ScanOnly, $Stamp, $ServiceNames)
    # $ErrorActionPreference from the jumpbox does NOT carry into the remote session;
    # without this, Set-Service/Start-Service failures were silently ignored.
    $ErrorActionPreference = 'Stop'

    $r = [ordered]@{
        Ok       = $false
        Error    = $null
        Changed  = @()   # services we switched from Disabled -> Manual (restored later)
        Services = @()   # human-readable before -> after per service
        Notes    = @()
        TaskName = "WUInstall_$Stamp"
        OutFile  = "$env:windir\Temp\WUResult_$Stamp.json"
        Worker   = "$env:windir\Temp\WUWorker_$Stamp.ps1"
    }
    # Ask the Service Control Manager directly (Get-Service StartType is missing on older PS)
    function Get-SvcInfo([string]$Name) {
        Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue
    }
    try {
        foreach ($n in $ServiceNames) {
            $before = Get-SvcInfo $n
            if (-not $before) { continue }
            try {
                if ($before.StartMode -eq 'Disabled') {
                    $scOut = & sc.exe config $n start= demand   # sc.exe reports errors on stdout
                    if ($LASTEXITCODE -ne 0) { throw "sc.exe config failed (exit $LASTEXITCODE): $($scOut -join ' ')" }
                    $r.Changed += $n
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

        # Not changed by this script - reported because they commonly cause 'found 0' or search errors
        $pol = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -ErrorAction SilentlyContinue
        $au  = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' -ErrorAction SilentlyContinue
        if ($au.UseWUServer -eq 1)                               { $r.Notes += "Policy: WSUS server $($pol.WUServer)" }
        if ($pol.DisableWindowsUpdateAccess -eq 1)               { $r.Notes += "Policy: DisableWindowsUpdateAccess=1" }
        if ($pol.DoNotConnectToWindowsUpdateInternetLocations -eq 1) { $r.Notes += "Policy: DoNotConnectToWindowsUpdateInternetLocations=1" }

        Set-Content -Path $r.Worker -Value $WorkerText -Encoding UTF8
        $taskArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$($r.Worker)`" -OutFile `"$($r.OutFile)`""
        if ($ScanOnly) { $taskArgs += ' -ScanOnly' }

        $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $taskArgs
        $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
        $settings  = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 6) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
        [void](Register-ScheduledTask -TaskName $r.TaskName -Action $action -Principal $principal -Settings $settings)
        Start-ScheduledTask -TaskName $r.TaskName
        $r.Ok = $true
    } catch {
        $r.Error = $_.Exception.Message
        # launch failed - put back anything we enabled so the machine isn't left half-changed
        foreach ($n in $r.Changed) {
            Stop-Service -Name $n -Force -ErrorAction SilentlyContinue
            & sc.exe config $n start= disabled | Out-Null
        }
    }
    [pscustomobject]$r
}

$state   = @{}   # computer -> launch info
$results = @{}   # computer -> final result row
function New-Row {
    param($Computer, $Status, $Found = '', $Installed = 0, $Failed = 0, $RebootRequired = '', $Services = '', $Detail = '')
    [pscustomobject]@{ Computer=$Computer; Status=$Status; Found=$Found; Installed=$Installed; Failed=$Failed
                       RebootRequired=$RebootRequired; Services=$Services; Detail=$Detail }
}

foreach ($item in $p1) {
    if ($item.Ok) { $state[$item.PSComputerName] = $item }
    else {
        $results[$item.PSComputerName] = New-Row -Computer $item.PSComputerName -Status 'LaunchFailed' `
            -Services (@($item.Services) -join '; ') -Detail ((@($item.Error) + @($item.Notes) | Where-Object { $_ }) -join '; ')
    }
}
foreach ($e in $p1Err) {
    $name = "$($e.TargetObject)"
    if (-not $name) { $name = '(unknown)' }
    if (-not $results.ContainsKey($name) -and -not $state.ContainsKey($name)) {
        $results[$name] = New-Row -Computer $name -Status 'Unreachable' -Detail $e.Exception.Message
    }
}
foreach ($t in $targets) {
    if (-not $state.ContainsKey($t) -and -not $results.ContainsKey($t)) {
        $results[$t] = New-Row -Computer $t -Status 'NoResponse' -Detail 'No result returned'
    }
}
Write-Host "      Launched on $($state.Count); failed/unreachable: $($results.Count)"
foreach ($c in $state.Keys) {
    Write-Host "      $c : $(@($state[$c].Services) -join ', ')"
}

# ---------------------------------------------------------------- phase 2: poll
Write-Host "`n[2/3] Waiting for updates to finish (timeout $TimeoutMinutes min)..." -ForegroundColor Cyan
$pending  = New-Object 'System.Collections.Generic.List[string]'
foreach ($k in $state.Keys) { $pending.Add([string]$k) }
$deadline = (Get-Date).AddMinutes($TimeoutMinutes)
$finished = @{}  # computer -> parsed json (or $null)

while ($pending.Count -gt 0 -and (Get-Date) -lt $deadline) {
    Start-Sleep -Seconds $PollSeconds
    $any = $state[$pending[0]]
    $poll = Invoke-Command @cred -ComputerName $pending.ToArray() -ThrottleLimit $ThrottleLimit `
        -ArgumentList $any.OutFile, $any.TaskName -ErrorAction SilentlyContinue -ScriptBlock {
        param($Out, $Task)
        if (Test-Path $Out) {
            [pscustomobject]@{ Done = $true; Json = (Get-Content $Out -Raw) }
        } else {
            $t = Get-ScheduledTask -TaskName $Task -ErrorAction SilentlyContinue
            $running = $t -and $t.State -eq 'Running'
            [pscustomobject]@{ Done = (-not $running); Json = $null }
        }
    }
    foreach ($p in $poll) {
        if ($p.Done) {
            $finished[$p.PSComputerName] = if ($p.Json) { $p.Json | ConvertFrom-Json } else { $null }
            [void]$pending.Remove($p.PSComputerName)
            Write-Host "      done: $($p.PSComputerName)  ($($pending.Count) remaining)"
        }
    }
}

foreach ($c in $pending) {
    $results[$c] = New-Row -Computer $c -Status 'TimedOut' -Services (@($state[$c].Services) -join '; ') `
        -Detail 'Task may still be running; task/files/services left as-is (services still enabled)'
}

# ---------------------------------------------------------------- phase 3: cleanup + restore
Write-Host "`n[3/3] Cleaning up temp task/files and restoring service state..." -ForegroundColor Cyan
$done     = @($finished.Keys)
$restored = @{}  # computer -> restore summary
if ($done.Count -gt 0) {
    $any = $state[$done[0]]
    Invoke-Command @cred -ComputerName $done -ThrottleLimit $ThrottleLimit -ErrorAction SilentlyContinue `
        -ArgumentList $any.TaskName, $any.OutFile, $any.Worker -ScriptBlock {
        param($Task, $Out, $Worker)
        # remove only what this script created
        Unregister-ScheduledTask -TaskName $Task -Confirm:$false -ErrorAction SilentlyContinue
        Remove-Item -Path $Out, $Worker -ErrorAction SilentlyContinue
    } | Out-Null

    foreach ($c in $done) {
        $changed = @($state[$c].Changed)
        if (-not $changed) { continue }
        if ($LeaveServicesEnabled) { $restored[$c] = "left enabled: $($changed -join ', ')"; continue }
        $out = Invoke-Command @cred -ComputerName $c -ArgumentList (,$changed) -ErrorAction SilentlyContinue -ScriptBlock {
            param($Names)
            foreach ($n in $Names) {
                Stop-Service -Name $n -Force -ErrorAction SilentlyContinue
                & sc.exe config $n start= disabled | Out-Null
                $w = Get-CimInstance -ClassName Win32_Service -Filter "Name='$n'"
                if ($w.StartMode -eq 'Disabled') { "$n -> Disabled/$($w.State)" }
                else                             { "RESTORE FAILED: $n is $($w.StartMode)/$($w.State)" }
            }
        }
        $restored[$c] = if ($out) { @($out) -join ', ' } else { "RESTORE FAILED: could not reach $c to re-disable $($changed -join ', ')" }
        if ($restored[$c] -match 'RESTORE FAILED') { Write-Host "      $c : $($restored[$c])" -ForegroundColor Red }
    }
}

foreach ($c in $done) {
    $j = $finished[$c]
    $svc = (@($state[$c].Services) + @($restored[$c]) | Where-Object { $_ }) -join '; '
    if ($null -eq $j) {
        $results[$c] = New-Row -Computer $c -Status 'NoResult' -Services $svc -Detail 'Task ended without writing a result'
        continue
    }
    $results[$c] = New-Row -Computer $c -Status $j.Status -Found $j.Found -Installed @($j.Installed).Count `
        -Failed @($j.Failed).Count -RebootRequired $j.RebootRequired -Services $svc `
        -Detail ((@("wuauserv during run: $($j.WuauservAtRun)") + @($j.Installed) + @($j.Failed | ForEach-Object { "FAILED: $_" }) +
                  @($j.Error) + @($state[$c].Notes) | Where-Object { $_ }) -join '; ')
}

# ---------------------------------------------------------------- report
$report = $results.Values | Sort-Object Computer
Write-Host "`n===== Summary =====" -ForegroundColor Green
$report | Format-Table Computer, Status, Found, Installed, Failed, RebootRequired, Services -AutoSize -Wrap

$csv = Join-Path (Get-Location) "WUReport_$stamp.csv"
$report | Export-Csv -Path $csv -NoTypeInformation
Write-Host "Detailed report: $csv" -ForegroundColor Green

$needReboot = $report | Where-Object { $_.RebootRequired -eq $true }
if ($needReboot) {
    Write-Host "`nPending reboot (NOT rebooted - schedule these yourself):" -ForegroundColor Yellow
    $needReboot | ForEach-Object { Write-Host "  $($_.Computer)" }
}
