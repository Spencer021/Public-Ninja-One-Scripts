<#
.SYNOPSIS
    Service Watchdog Deployment Script for NinjaOne

.DESCRIPTION
    Deploys a watchdog script that monitors core NinjaOne agent services.
    The watchdog:
      • Validates its own integrity using a SHA256 hash
      • Monitors and restarts NinjaOne services when needed
      • Reports activity back to NinjaOne if available
      • Runs every 5 minutes as a SYSTEM-level scheduled task

.NOTES
    Original inspiration and portions of the logic from Mikey O'Toole (homotechsual).
    Reference:
    https://github.com/homotechsual/Blog-Scripts/blob/main/NinjaOne%20Scripts/WatchDog.ps1

.REVISION
    Version: 1.0.0
    Last Updated: 2025-02-21
    Author: Spencer A. Heath
    Change Log:
      • Initial release of revised deployment script
      • Hardened file ACL handling
      • Added automatic SHA256 integrity validation
      • Parametrized the scheduled task configuration
      • Cleaned and documented code for public sharing

.DISCLAIMER
    This script is provided “as-is” with no guarantees or warranties of any kind.
    You are responsible for reviewing and testing it before deploying in any environment.
    The author (Spencer A. Heath) assumes no liability for damage, service interruption,
    or data loss resulting from use or misuse of this script.
#>

## Scheduled Task Variables
$InitialDelayMinutes = 1
$RepeatEveryMinutes = 5
$RepeatForDays = 3650
$RunAsUser = 'NT AUTHORITY\SYSTEM'
$RunElevation = 'Highest'
$Executable = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"


## Script Variables
$watchdogDir = 'C:\RMM\Scripts'
$watchdogPath = Join-Path $watchdogDir 'ServiceWatchdog.ps1'
$watchdogHashPath = Join-Path $watchdogDir 'ServiceWatchdog.hash'
$taskName = 'ServiceWatchdog'
$eventSource = 'NinjaOne-Watchdog'
$eventLogName = 'Application'

#####################################################################
## Do not modify below this line unless you know what you're doing ##
#####################################################################
if (-not (Test-Path $watchdogDir)) {
    New-Item -Path $watchdogDir -ItemType Directory -Force | Out-Null
}

$watchdogScript = @'

$ErrorActionPreference = 'Stop'

$services       = 'NinjaRMMAgent','ncstreamer'
$maxRetries     = 3
$hashFile       = 'C:\RMM\Scripts\ServiceWatchdog.hash'
$eventSource    = 'PTS-Watchdog'
$eventLogName   = 'Application'
$watchdogField  = 'watchDogActivity'

function Bark {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    try {
        ## Ensure Ninja module is loaded
        $ninjaModule = Get-Module 'NJCliPsh'
        if (-not $ninjaModule) {
            $ninjaAvailable = Get-Module -ListAvailable -Name 'NJCliPsh'
            if ($ninjaAvailable) {
                Import-Module 'NJCliPsh' -ErrorAction Stop
            }
            else {
                ## No Ninja module available; bail quietly
                return
            }
        }

        $payload = [PSCustomObject]@{
            Time     = (Get-Date).ToString('o')
            Hostname = $env:COMPUTERNAME
            Message  = $Message
        } | ConvertTo-Json -Compress

        Ninja-Property-Set -Name $watchdogField -Value $payload | Out-Null
    }
    catch {
        ## Swallow all, watchdog must not crash due to logging
    }
}

try {
    if (-not (Test-Path $hashFile)) {
        Bark -Message 'WATCHDOG_HASH_MISSING'
        try {
            if ([System.Diagnostics.EventLog]::SourceExists($eventSource)) {
                Write-EventLog -LogName $eventLogName -Source $eventSource -EventId 50001 -EntryType Warning -Message 'Watchdog hash file is missing.'
            }
        } catch {}
        exit 1
    }

    $expected = (Get-Content $hashFile -Raw).Trim()
    $current  = (Get-FileHash -Algorithm SHA256 -Path $MyInvocation.MyCommand.Path).Hash.Trim()

    if ($expected -ne $current) {
        Bark -Message 'WATCHDOG_SCRIPT_TAMPERED'
        try {
            if ([System.Diagnostics.EventLog]::SourceExists($eventSource)) {
                Write-EventLog -LogName $eventLogName -Source $eventSource -EventId 50002 -EntryType Error -Message 'Watchdog script integrity check FAILED. Script has been modified.'
            }
        } catch {}
        exit 1
    }
}
catch {
    ## If integrity check itself fails, fail safe
    Bark -Message "WATCHDOG_INTEGRITY_ERROR: $($_.Exception.Message)"
    exit 1
}

foreach ($service in $services) {
    try {
        $serviceObject = Get-CimInstance -ClassName 'Win32_Service' -Filter "Name = '$service'"
    }
    catch {
        Bark -Message "SERVICE_QUERY_ERROR: Service=$service Error=$($_.Exception.Message)"
        continue
    }

    if (-not $serviceObject) {
        Bark -Message "SERVICE_NOT_FOUND: Service=$service"
        continue
    }

    ## Ensure start mode is Automatic
    try {
        if ($serviceObject.StartMode -ne 'Auto') {
            $null = $serviceObject | Invoke-CimMethod -MethodName 'ChangeStartMode' -Arguments @{ StartMode = 'Automatic' }
            Bark -Message "SERVICE_STARTMODE_CORRECTED: Service=$service Previous=$($serviceObject.StartMode) New=Automatic"
        }
    }
    catch {
        Bark -Message "SERVICE_STARTMODE_ERROR: Service=$service Error=$($_.Exception.Message)"
    }

    $retries = 0
    while ($retries -lt $maxRetries) {
        $serviceObject = Get-CimInstance -ClassName 'Win32_Service' -Filter "Name = '$service'"

        if ($serviceObject.State -ne 'Running') {
            try {
                $null = $serviceObject | Invoke-CimMethod -MethodName 'StartService'
                Start-Sleep -Seconds 10
                $refreshed = Get-CimInstance -ClassName 'Win32_Service' -Filter "Name = '$service'"

                if ($refreshed.State -eq 'Running') {
                    Bark -Message "SERVICE_RESTARTED: Service=$service Attempts=$($retries + 1)"
                    break
                }
                else {
                    $retries++
                    if ($retries -ge $maxRetries) {
                        Bark -Message "SERVICE_FAILED_TO_START: Service=$service Attempts=$retries"
                        try {
                            if ([System.Diagnostics.EventLog]::SourceExists($eventSource)) {
                                $msg = "Failed to start service $service after $retries retries."
                                Write-EventLog -LogName $eventLogName -Source $eventSource -EventId 50003 -EntryType Error -Message $msg
                            }
                        } catch {}
                    }
                }
            }
            catch {
                $retries++
                Bark -Message "SERVICE_START_ERROR: Service=$service Attempt=$retries Error=$($_.Exception.Message)"
            }
        }
        else {
            ## Already running
            break
        }
    }
}

Bark -Message 'WATCHDOG_RUN_COMPLETED'
'@
$watchdogScript | Out-File -FilePath $watchdogPath -Encoding UTF8 -Force

$hash = (Get-FileHash -Algorithm SHA256 -Path $watchdogPath).Hash
$hash | Out-File -FilePath $watchdogHashPath -Encoding ASCII -Force

function Set-SecureAcl {
    param([string]$Path)

    if (-not (Test-Path $Path)) { return }

    $acl = Get-Acl $Path
    $acl.SetAccessRuleProtection($true, $false)

    foreach ($ace in $acl.Access) {
        [void]$acl.RemoveAccessRule($ace)
    }

    $rules = @(
        [System.Security.AccessControl.FileSystemAccessRule]::new("SYSTEM", "FullControl", "Allow"),
        [System.Security.AccessControl.FileSystemAccessRule]::new("Administrators", "FullControl", "Allow"),
        [System.Security.AccessControl.FileSystemAccessRule]::new("Users", "ReadAndExecute", "Allow")
    )

    foreach ($rule in $rules) {
        [void]$acl.AddAccessRule($rule)
    }

    Set-Acl -Path $Path -AclObject $acl
}


Set-SecureAcl -Path $watchdogPath
Set-SecureAcl -Path $watchdogHashPath

## Ensure event source exists
try {
    if (-not [System.Diagnostics.EventLog]::SourceExists($eventSource)) {
        New-EventLog -LogName $eventLogName -Source $eventSource
    }
}
catch {
    ## Not critical if this fails
}

$ScriptPath = $watchdogPath
$TaskName = $taskName
$taskAction = New-ScheduledTaskAction -Execute $Executable -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""
$taskTrigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes($InitialDelayMinutes) -RepetitionInterval ([TimeSpan]::FromMinutes($RepeatEveryMinutes)) -RepetitionDuration ([TimeSpan]::FromDays($RepeatForDays))
$taskPrincipal = New-ScheduledTaskPrincipal -UserId $RunAsUser -RunLevel $RunElevation
$task = New-ScheduledTask -Action $taskAction -Trigger $taskTrigger -Principal $taskPrincipal
try {
    Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force -ErrorAction Stop | Out-Null
}
catch {
    Write-Host "ERROR: Failed to register scheduled task $TaskName"
    Write-Host $_.Exception.Message
    exit 1
}

Write-Host "ServiceWatchdog deployed:"
Write-Host " - Script: $ScriptPath"
Write-Host " - Task:   $TaskName"
Write-Host " - Runs:   Every $RepeatEveryMinutes minutes"
Write-Host " - Starts: $(Get-Date).AddMinutes($InitialDelayMinutes)"
Write-Host " - Duration: $RepeatForDays days"

