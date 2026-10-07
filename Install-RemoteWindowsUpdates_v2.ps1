<#
.SYNOPSIS
    Run from a jumpbox. Takes many computer names, enables/starts the Windows Update
    services on each, downloads + installs updates, then restores the original service
    state. Does NOT reboot anything - it only reports which machines need one.

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
Write-Host "  Changes: temporarily set cryptsvc/bits/wuauserv/UsoSvc to Manual+start if disabled/stopped;"
Write-Host "           create a temp worker script + one-shot SYSTEM scheduled task in C:\Windows\Temp"
Write-Host "           Restore disabled services afterward: $(-not $LeaveServicesEnabled)"
if (-not $Force) {
    if ((Read-Host "Proceed? (y/N)") -notmatch '^(y|yes)$') { Write-Host "Aborted."; return }
}

$cred = @{}
if ($Credential) { $cred.Credential = $Credential }

# ---------------------------------------------------------------- worker (runs on target as SYSTEM)
$workerText = @'
param([string]$OutFile, [switch]$ScanOnly)
$out = [ordered]@{ Status='Started'; Found=0; Installed=@(); Failed=@(); RebootRequired=$false; Error=$null }
try {
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
    -ArgumentList $workerText, $ScanOnly.IsPresent, $stamp `
    -ErrorAction SilentlyContinue -ErrorVariable p1Err -ScriptBlock {
    param($WorkerText, $ScanOnly, $Stamp)

    $r = [ordered]@{
        Ok       = $false
        Error    = $null
        Original = @()
        TaskName = "WUInstall_$Stamp"
        OutFile  = "$env:windir\Temp\WUResult_$Stamp.json"
        Worker   = "$env:windir\Temp\WUWorker_$Stamp.ps1"
    }
    try {
        foreach ($n in 'cryptsvc','bits','wuauserv','UsoSvc') {
            $s = Get-Service -Name $n -ErrorAction SilentlyContinue
            if (-not $s) { continue }
            $r.Original += "$n|$($s.StartType)|$($s.Status)"
            if ($s.StartType -eq 'Disabled') { Set-Service -Name $n -StartupType Manual }
            if ($s.Status -ne 'Running')     { Start-Service -Name $n -ErrorAction SilentlyContinue }
        }
        if ((Get-Service wuauserv).Status -ne 'Running') { throw "wuauserv would not start" }

        Set-Content -Path $r.Worker -Value $WorkerText -Encoding UTF8
        $args = "-NoProfile -ExecutionPolicy Bypass -File `"$($r.Worker)`" -OutFile `"$($r.OutFile)`""
        if ($ScanOnly) { $args += ' -ScanOnly' }

        $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $args
        $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
        $settings  = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 6) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
        [void](Register-ScheduledTask -TaskName $r.TaskName -Action $action -Principal $principal -Settings $settings)
        Start-ScheduledTask -TaskName $r.TaskName
        $r.Ok = $true
    } catch {
        $r.Error = $_.Exception.Message
    }
    [pscustomobject]$r
}

$state   = @{}   # computer -> launch info
$results = @{}   # computer -> final result row

foreach ($item in $p1) {
    if ($item.Ok) { $state[$item.PSComputerName] = $item }
    else          { $results[$item.PSComputerName] = [pscustomobject]@{ Computer=$item.PSComputerName; Status='LaunchFailed'; Found=''; Installed=0; Failed=0; RebootRequired=''; Detail=$item.Error } }
}
foreach ($e in $p1Err) {
    $name = "$($e.TargetObject)"
    if (-not $name) { $name = '(unknown)' }
    if (-not $results.ContainsKey($name) -and -not $state.ContainsKey($name)) {
        $results[$name] = [pscustomobject]@{ Computer=$name; Status='Unreachable'; Found=''; Installed=0; Failed=0; RebootRequired=''; Detail=$e.Exception.Message }
    }
}
foreach ($t in $targets) {
    if (-not $state.ContainsKey($t) -and -not $results.ContainsKey($t)) {
        $results[$t] = [pscustomobject]@{ Computer=$t; Status='NoResponse'; Found=''; Installed=0; Failed=0; RebootRequired=''; Detail='No result returned' }
    }
}
Write-Host "      Launched on $($state.Count); failed/unreachable: $($results.Count)"

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
    $results[$c] = [pscustomobject]@{ Computer=$c; Status='TimedOut'; Found=''; Installed=0; Failed=0; RebootRequired=''; Detail='Task may still be running; task/files/services left as-is' }
}

# ---------------------------------------------------------------- phase 3: cleanup + restore
Write-Host "`n[3/3] Cleaning up temp task/files and restoring service state..." -ForegroundColor Cyan
$done = @($finished.Keys)
if ($done.Count -gt 0) {
    $any = $state[$done[0]]
    Invoke-Command @cred -ComputerName $done -ThrottleLimit $ThrottleLimit -ErrorAction SilentlyContinue `
        -ArgumentList $any.TaskName, $any.OutFile, $any.Worker, (-not $LeaveServicesEnabled) -ScriptBlock {
        param($Task, $Out, $Worker, $Restore)
        # remove only what this script created
        Unregister-ScheduledTask -TaskName $Task -Confirm:$false -ErrorAction SilentlyContinue
        Remove-Item -Path $Out, $Worker -ErrorAction SilentlyContinue
    } | Out-Null

    if (-not $LeaveServicesEnabled) {
        foreach ($c in $done) {
            $disabled = $state[$c].Original | Where-Object { ($_ -split '\|')[1] -eq 'Disabled' } |
                        ForEach-Object { ($_ -split '\|')[0] }
            if (-not $disabled) { continue }
            Invoke-Command @cred -ComputerName $c -ArgumentList (,$disabled) -ErrorAction SilentlyContinue -ScriptBlock {
                param($Names)
                foreach ($n in $Names) {
                    Stop-Service -Name $n -Force -ErrorAction SilentlyContinue
                    Set-Service  -Name $n -StartupType Disabled
                }
            } | Out-Null
        }
    }
}

foreach ($c in $done) {
    $j = $finished[$c]
    if ($null -eq $j) {
        $results[$c] = [pscustomobject]@{ Computer=$c; Status='NoResult'; Found=''; Installed=0; Failed=0; RebootRequired=''; Detail='Task ended without writing a result' }
        continue
    }
    $results[$c] = [pscustomobject]@{
        Computer       = $c
        Status         = $j.Status
        Found          = $j.Found
        Installed      = @($j.Installed).Count
        Failed         = @($j.Failed).Count
        RebootRequired = $j.RebootRequired
        Detail         = (@($j.Installed) + @($j.Failed | ForEach-Object { "FAILED: $_" }) + @($j.Error) | Where-Object { $_ }) -join '; '
    }
}

# ---------------------------------------------------------------- report
$report = $results.Values | Sort-Object Computer
Write-Host "`n===== Summary =====" -ForegroundColor Green
$report | Format-Table Computer, Status, Found, Installed, Failed, RebootRequired -AutoSize

$csv = Join-Path (Get-Location) "WUReport_$stamp.csv"
$report | Export-Csv -Path $csv -NoTypeInformation
Write-Host "Detailed report: $csv" -ForegroundColor Green

$needReboot = $report | Where-Object { $_.RebootRequired -eq $true }
if ($needReboot) {
    Write-Host "`nPending reboot (NOT rebooted - schedule these yourself):" -ForegroundColor Yellow
    $needReboot | ForEach-Object { Write-Host "  $($_.Computer)" }
}
