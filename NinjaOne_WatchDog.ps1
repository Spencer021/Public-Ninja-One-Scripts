<#
.SYNOPSIS
    Service Watchdog Management Script for NinjaOne

.DESCRIPTION
    Manages the deployment and removal of a watchdog script that monitors core NinjaOne agent services.
    Can install, remove, or update the watchdog task and associated files.

.ENVIRONMENT VARIABLES
    $env:action - Specify 'Install' to deploy the watchdog or 'Remove' to uninstall it.
    $env:legacyTaskToRemove - Name of an old/legacy task to remove when installing the new watchdog.
    $env:watchdogDir - Directory for watchdog files. Default: C:\RMM\ninjawatchdog
    $env:taskName - Name of the scheduled task. Default: ServiceWatchdog
    $env:watchdogCustomFieldName - Custom field name in NinjaOne for logging watchdog activity. Default: watchDogActivity

.NOTES
    Original inspiration and portions of the logic from Mikey O'Toole (homotechsual).
    Reference:
    https://github.com/homotechsual/Blog-Scripts/blob/main/NinjaOne%20Scripts/WatchDog.ps1

.REVISION
    Version: 1.1.0
    Last Updated: 2025-02-21
    Author: Spencer A. Heath
    Contributor: Matt Dewart
    Change Log:
      v1.1.0:
      • Added management functionality for install/remove actions and cleanup of superceded tasks.
      • Changed Task trigger to run on startup and repeat hourly indefinitely.
      • Converted here-string script to double quote for variable injection.
      • Added environment variable support for configuration.
      v1.0.0:
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

#region Variable Definitions

$Action = $env:action
$LegacyTaskName = $env:legacyTaskToRemove

$watchdogDir = if ($env:watchdogDirectory) {
    $env:watchdogDirectory.TrimEnd('\')
} else {
    'C:\RMM\ninjawatchdog' 
}

$taskName = if ($env:taskName) {
    $env:taskName.Trim()
} else {
    'NinjaWatchdog' 
}

$watchDogActivity = $env:watchdogCustomFieldName
if ([string]::IsNullOrWhiteSpace($watchDogActivity)) {
    $watchDogActivity = 'watchDogActivity'
}

# Derived Paths
$watchdogPath = Join-Path $watchdogDir 'NinjaWatchdog.ps1'
$watchdogHashPath = Join-Path $watchdogDir 'NinjaWatchdog.hash'

# Event Logging Configuration
$eventSource = 'NinjaOne-Watchdog'
$eventLogName = 'Application'

# Scheduled Task Configuration
$InitialDelayMinutes = 1
$RepeatEveryMinutes = 60
$RunAsUser = 'NT AUTHORITY\SYSTEM'
$RunElevation = 'Highest'
$Executable = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
#endregion

#region Input Validation
# Validate action parameter
if (-not $Action) {
    Write-Host "ERROR: Action not specified. Set `$env:action to 'Install' or 'Remove'" 
    Write-Host "Example: `$env:action = 'Install'" 
    exit 1
}

if ($Action -notin @('Install', 'Remove')) {
    Write-Host "ERROR: Invalid action '$Action'. Valid options are 'Install' or 'Remove'"
    exit 1
}

# Validate watchdog directory path
if ([string]::IsNullOrWhiteSpace($watchdogDir) -or $watchdogDir.Length -gt 248) {
    Write-Host 'ERROR: Invalid watchdog directory path. Path must be valid and under 248 characters'
    exit 1
}

# Validate task name
if ([string]::IsNullOrWhiteSpace($taskName) -or $taskName -match '[<>:"|?*]') {
    Write-Host 'ERROR: Invalid task name. Task name cannot contain special characters: < > : \ | ? *'
    exit 1
}
#endregion

function Remove-LegacyTask {
    param([string]$TaskNameToRemove)
    
    if ([string]::IsNullOrWhiteSpace($TaskNameToRemove)) {
        Write-Host 'Warning: No legacy task name provided'
        return
    }
    
    try {
        if (Get-ScheduledTask -TaskName $TaskNameToRemove -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskName $TaskNameToRemove -Confirm:$false -ErrorAction Stop
            Write-Host "Removed legacy scheduled task: $TaskNameToRemove"
        } else {
            Write-Host "Legacy task '$TaskNameToRemove' not found (already removed or never existed)"
        }
    } catch {
        Write-Host "Warning: Failed to remove legacy task '$TaskNameToRemove' - $($_.Exception.Message)"
    }
}

function Remove-WatchdogComponents {
    param([string]$TaskNameToRemove)
    
    if ([string]::IsNullOrWhiteSpace($TaskNameToRemove)) {
        Write-Host 'ERROR: Task name is required for removal'
        return $false
    }
    
    $success = $true
    
    # Remove scheduled task
    try {
        if (Get-ScheduledTask -TaskName $TaskNameToRemove -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskName $TaskNameToRemove -Confirm:$false -ErrorAction Stop
            Write-Host "Removed scheduled task: $TaskNameToRemove"
        } else {
            Write-Host "Scheduled task '$TaskNameToRemove' not found"
        }
    } catch {
        Write-Host "ERROR: Failed to remove task '$TaskNameToRemove' - $($_.Exception.Message)"
        $success = $false
    }
    if (Test-Path $watchdogDir) {
        try {
            Remove-Item -Path $watchdogDir -Recurse -Force -ErrorAction Stop
            Write-Host "Removed watchdog directory: $watchdogDir"
        } catch {
            Write-Host "ERROR: Failed to remove directory '$watchdogDir' - $($_.Exception.Message)"
            $success = $false
        }
    }
    return $success
}

function Install-WatchdogComponents {
    # Check if task already exists
    $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($existingTask) {
        Write-Host "INFO: Task '$taskName' already exists - will update/replace"
    } else {
        Write-Host "INFO: Creating new task '$taskName'"
    }
    
    if (-not (Test-Path $watchdogDir)) {
        New-Item -Path $watchdogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        Write-Host "Created watchdog directory: $watchdogDir"
    }

    $watchdogScript = @"

`$ErrorActionPreference = 'Stop'

`$services       = 'NinjaRMMAgent','ncstreamer'
`$maxRetries     = 3
`$hashFile       = '$watchdogHashPath'
`$eventSource    = '$eventSource'
`$eventLogName   = '$eventLogName'
`$watchdogField  = '$watchDogActivity'

function Bark {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = `$true)]
        [string]`$Message
    )

    try {
        ## Ensure Ninja module is loaded
        `$ninjaModule = Get-Module 'NJCliPsh'
        if (-not `$ninjaModule) {
            `$ninjaAvailable = Get-Module -ListAvailable -Name 'NJCliPsh'
            if (`$ninjaAvailable) {
                Import-Module 'NJCliPsh' -ErrorAction Stop
            }
            else {
                ## No Ninja module available; bail quietly
                return
            }
        }

        `$payload = [PSCustomObject]@{
            Time     = (Get-Date).ToString('o')
            Hostname = `$env:COMPUTERNAME
            Message  = `$Message
        } | ConvertTo-Json -Compress

        Ninja-Property-Set -Name `$watchdogField -Value `$payload | Out-Null
    }
    catch {
        ## Swallow all, watchdog must not crash due to logging
    }
}

try {
    if (-not (Test-Path `$hashFile)) {
        Bark -Message 'WATCHDOG_HASH_MISSING'
        try {
            if ([System.Diagnostics.EventLog]::SourceExists(`$eventSource)) {
                Write-EventLog -LogName `$eventLogName -Source `$eventSource -EventId 50001 -EntryType Warning -Message 'Watchdog hash file is missing.'
            }
        } catch {}
        exit 1
    }

    `$expected = (Get-Content `$hashFile -Raw).Trim()
    `$current  = (Get-FileHash -Algorithm SHA256 -Path `$MyInvocation.MyCommand.Path).Hash.Trim()

    if (`$expected -ne `$current) {
        Bark -Message 'WATCHDOG_SCRIPT_TAMPERED'
        try {
            if ([System.Diagnostics.EventLog]::SourceExists(`$eventSource)) {
                Write-EventLog -LogName `$eventLogName -Source `$eventSource -EventId 50002 -EntryType Error -Message 'Watchdog script integrity check FAILED. Script has been modified.'
            }
        } catch {}
        exit 1
    }
}
catch {
    ## If integrity check itself fails, fail safe
    Bark -Message "WATCHDOG_INTEGRITY_ERROR: `$(`$_.Exception.Message)"
    exit 1
}

foreach (`$service in `$services) {
    try {
        `$serviceObject = Get-CimInstance -ClassName 'Win32_Service' -Filter "Name = '`$service'"
    }
    catch {
        Bark -Message "SERVICE_QUERY_ERROR: Service=`$service Error=`$(`$_.Exception.Message)"
        continue
    }

    if (-not `$serviceObject) {
        Bark -Message "SERVICE_NOT_FOUND: Service=`$service"
        continue
    }

    ## Ensure start mode is Automatic
    try {
        if (`$serviceObject.StartMode -ne 'Auto') {
            `$null = `$serviceObject | Invoke-CimMethod -MethodName 'ChangeStartMode' -Arguments @{ StartMode = 'Automatic' }
            Bark -Message "SERVICE_STARTMODE_CORRECTED: Service=`$service Previous=`$(`$serviceObject.StartMode) New=Automatic"
        }
    }
    catch {
        Bark -Message "SERVICE_STARTMODE_ERROR: Service=`$service Error=`$(`$_.Exception.Message)"
    }

    `$retries = 0
    while (`$retries -lt `$maxRetries) {
        `$serviceObject = Get-CimInstance -ClassName 'Win32_Service' -Filter "Name = '`$service'"

        if (`$serviceObject.State -ne 'Running') {
            try {
                `$null = `$serviceObject | Invoke-CimMethod -MethodName 'StartService'
                Start-Sleep -Seconds 10
                `$refreshed = Get-CimInstance -ClassName 'Win32_Service' -Filter "Name = '`$service'"

                if (`$refreshed.State -eq 'Running') {
                    Bark -Message "SERVICE_RESTARTED: Service=`$service Attempts=`$(`$retries + 1)"
                    break
                }
                else {
                    `$retries++
                    if (`$retries -ge `$maxRetries) {
                        Bark -Message "SERVICE_FAILED_TO_START: Service=`$service Attempts=`$retries"
                        try {
                            if ([System.Diagnostics.EventLog]::SourceExists(`$eventSource)) {
                                `$msg = "Failed to start service `$service after `$retries retries."
                                Write-EventLog -LogName `$eventLogName -Source `$eventSource -EventId 50003 -EntryType Error -Message `$msg
                            }
                        } catch {}
                    }
                }
            }
            catch {
                `$retries++
                Bark -Message "SERVICE_START_ERROR: Service=`$service Attempt=`$retries Error=`$(`$_.Exception.Message)"
            }
        }
        else {
            ## Already running
            break
        }
    }
}

Bark -Message 'WATCHDOG_RUN_COMPLETED'
"@
    # Write script file with error handling
    try {
        $watchdogScript | Out-File -FilePath $watchdogPath -Encoding UTF8 -Force -ErrorAction Stop
        Write-Host "Created watchdog script: $watchdogPath"
    } catch {
        Write-Host "ERROR: Failed to create watchdog script - $($_.Exception.Message)"
        return $false
    }
    
    # Generate and write hash file
    try {
        $hash = (Get-FileHash -Algorithm SHA256 -Path $watchdogPath -ErrorAction Stop).Hash
        $hash | Out-File -FilePath $watchdogHashPath -Encoding ASCII -Force -ErrorAction Stop
        Write-Host "Created hash file: $watchdogHashPath"
    } catch {
        Write-Host "ERROR: Failed to create hash file - $($_.Exception.Message)"
        return $false
    }

    function Set-SecureAcl {
        param([string]$Path)

        if (-not (Test-Path $Path)) {
            return 
        }

        $acl = Get-Acl $Path
        $acl.SetAccessRuleProtection($true, $false)

        foreach ($ace in $acl.Access) {
            [void]$acl.RemoveAccessRule($ace)
        }

        $rules = @(
            [System.Security.AccessControl.FileSystemAccessRule]::new('SYSTEM', 'FullControl', 'Allow'),
            [System.Security.AccessControl.FileSystemAccessRule]::new('Administrators', 'FullControl', 'Allow'),
            [System.Security.AccessControl.FileSystemAccessRule]::new('Users', 'ReadAndExecute', 'Allow')
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
            New-EventLog -LogName $eventLogName -Source $eventSource -ErrorAction Stop
            Write-Host "Created event log source: $eventSource"
        }
    } catch {
        Write-Host "Warning: Failed to create event log source - $($_.Exception.Message)"
    }

    # Register scheduled task with enhanced error handling
    try {
        $ScriptPath = $watchdogPath
        $TaskName = $taskName
        $taskAction = New-ScheduledTaskAction -Execute $Executable -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`"" -ErrorAction Stop
        $taskTrigger = New-ScheduledTaskTrigger -AtStartup -ErrorAction Stop
        $taskTrigger.Repetition = (New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval ([TimeSpan]::FromMinutes($RepeatEveryMinutes))).Repetition
        $taskPrincipal = New-ScheduledTaskPrincipal -UserId $RunAsUser -RunLevel $RunElevation -ErrorAction Stop
        $task = New-ScheduledTask -Action $taskAction -Trigger $taskTrigger -Principal $taskPrincipal -ErrorAction Stop

        Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force -ErrorAction Stop | Out-Null
        
        Write-Host 'ServiceWatchdog deployed successfully:'
        Write-Host " - Script: $ScriptPath"
        Write-Host " - Task:   $TaskName"
        Write-Host " - Runs:   At startup and every $RepeatEveryMinutes minutes"
        Write-Host ' - Duration: Indefinitely'
        
        return $true
    } catch {
        Write-Host "ERROR: Failed to register scheduled task '$TaskName' - $($_.Exception.Message)"
        return $false
    }
}

# Remove legacy task if specified - to supercede previous versions
if ($LegacyTaskName) {
    Remove-LegacyTask -TaskNameToRemove $LegacyTaskName
}

# Main execution logic

try {
    switch ($Action) {
        'Install' {
            Install-WatchdogComponents
            Write-Host 'ServiceWatchdog installation completed.'
        }
        'Remove' {
            Remove-WatchdogComponents -TaskNameToRemove $taskName
            Write-Host 'ServiceWatchdog components removed successfully.'
        }
    }
} catch {
    Write-Host "ERROR: Unexpected error during $Action operation - $($_.Exception.Message)"
}
