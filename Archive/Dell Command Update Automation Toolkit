<#
.SYNOPSIS
    Automates Dell Command Update (DCU) operations for Dell systems within NinjaOne. 
    This script manages installation, scanning, and updating tasks.

.DESCRIPTION
    This script streamlines Dell system management by interfacing with Dell Command Update (DCU) and NinjaOne. Key features include:
    - Validates the system as a Dell device and removes incompatible "Dell Update*" applications.
    - Downloads and installs the latest DCU version dynamically from Dell’s support site.
    - Performs update scans (general or BIOS/firmware-specific) and applies updates based on user-selected options.
    - Integrates with NinjaOne by setting custom fields for update status and results (except for the general scan, which outputs to CLI only).

    The script’s behavior is controlled via the NinjaOne script variable `pleaseSelectAnOptionToRun`. Users must configure this variable as a dropdown in NinjaOne with the following options:
    - "Install" - Installs DCU after removing incompatible apps.
    - "Remove Incompatible Versions" - Removes conflicting Dell Update apps.
    - "Run Scan" - Scans for all updates and displays results in the CLI.
    - "Run BIOS and Firmware Scan" - Scans for BIOS/firmware updates, setting results in NinjaOne custom fields.
    - "Run Scan And Install All" - Scans and installs all updates.
    - "Run Scan And Install Excluding BIOS and Firmware" - Scans and installs updates, excluding BIOS/firmware.
    - "Run Scan And Install BIOS and Firmware ONLY" - Scans and installs only BIOS/firmware updates.

    Custom fields used:
    - $firmwareBiosUpdateField: Indicates update availability (e.g., "BIOS/Firmware Updates Available", "No Updates Found"). (Defualt: DCU1)
    - $biosFirmwareUpdatesField: Stores scan output or exit code details with descriptions. (Defualt: DCU2)

    Requires a DROP DOWN script variable Named "pleaseSelectAnOptionToRun"
    Required Drop Down Options
       - "Install"
       - "Remove Incompatible Versions"
       - "Run Scan"
       - "Run BIOS and Firmware Scan"
       - "Run Scan And Install All"
       - "Run Scan And Install Excluding BIOS and Firmware"
       - "Run Scan And Install BIOS and Firmware ONLY"

.AUTHOR
    By: Spencer Heath
    DATE: 23 Feb 2025

.GITHUB
    https://github.com/Sp-e-n-c-er

.NOTES
    **Disclaimer**: This script is provided "as is" under the MIT License. 
    Use at your own risk; the author is not responsible for any damages or issues arising from its use. 
    Licensed under MIT: Copyright (c) Spencer Heath. Permission is granted to use, copy, modify, and distribute this software freely, 
    provided the original copyright and this notice are retained. See https://opensource.org/licenses/MIT for full details.
#>

Set-Location -Path $env:SystemRoot
$ProgressPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'
if ([Net.ServicePointManager]::SecurityProtocol -notcontains 'Tls12') {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}
## Update these fields to match your custom field names in NinjaOne
$firmwareBiosUpdateField = "DCU1"
$biosFirmwareUpdatesField = "DCU2"

if ((Get-CimInstance -ClassName Win32_BIOS).Manufacturer -notlike '*Dell*') {
    Write-Output 'Not a Dell system. Aborting...'
    exit 0
}

$RegPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')

function Handle-DCUExitCode {
    param (
        [int]$ExitCode,
        [string]$ScanOutput = ''
    )
    switch ($ExitCode) {
        0 {
            Ninja-Property-Set $firmwareBiosUpdateField "BIOS/Firmware Updates Available"
            Ninja-Property-Set $biosFirmwareUpdatesField $ScanOutput
        }
        { $_ -in 1, 5 } {
            Ninja-Property-Set $firmwareBiosUpdateField "Reboot Required"
            Ninja-Property-Set $biosFirmwareUpdatesField "Exit Code: $ExitCode - Reboot required to complete a previous operation"
        }
        2 {
            Ninja-Property-Set $firmwareBiosUpdateField "Fatal Error"
            Ninja-Property-Set $biosFirmwareUpdatesField "Exit Code: $ExitCode - Dell Command Update returned a fatal error"
        }
        3 {
            Ninja-Property-Set $firmwareBiosUpdateField "Not a Dell System"
            Ninja-Property-Set $biosFirmwareUpdatesField "Exit Code: $ExitCode - Not a Dell system"
        }
        4 {
            Ninja-Property-Set $firmwareBiosUpdateField "Admin Privilege Required"
            Ninja-Property-Set $biosFirmwareUpdatesField "Exit Code: $ExitCode - Admin privileges required"
        }
        6 {
            Ninja-Property-Set $firmwareBiosUpdateField "Currently Running"
            Ninja-Property-Set $biosFirmwareUpdatesField "Exit Code: $ExitCode - Dell Command Update is currently running"
        }
        7 {
            Ninja-Property-Set $firmwareBiosUpdateField "System Model Not Supported"
            Ninja-Property-Set $biosFirmwareUpdatesField "Exit Code: $ExitCode - System model not supported"
        }
        8 {
            Ninja-Property-Set $firmwareBiosUpdateField "No Update Filters Configured"
            Ninja-Property-Set $biosFirmwareUpdatesField "Exit Code: $ExitCode - No update filters configured"
        }
        { $_ -ge 100 -and $_ -le 113 } {
            Ninja-Property-Set $firmwareBiosUpdateField "Input Validation Error"
            Ninja-Property-Set $biosFirmwareUpdatesField "Exit Code: $ExitCode - Input validation error"
        }
        500 {
            Ninja-Property-Set $firmwareBiosUpdateField "No Updates Found"
            Ninja-Property-Set $biosFirmwareUpdatesField "Exit Code: $ExitCode - No updates available"
        }
        { $_ -ge 501 -and $_ -le 503 } {
            Ninja-Property-Set $firmwareBiosUpdateField "Scan Error - Retry"
            Ninja-Property-Set $biosFirmwareUpdatesField "Exit Code: $ExitCode - Scan operation failed"
        }
        { $_ -ge 1000 -and $_ -le 1002 } {
            Ninja-Property-Set $firmwareBiosUpdateField "ApplyUpdates Error - Retry"
            Ninja-Property-Set $biosFirmwareUpdatesField "Exit Code: $ExitCode - Apply updates operation failed"
        }
        { $_ -ge 1505 -and $_ -le 1506 } {
            Ninja-Property-Set $firmwareBiosUpdateField "Configure Error - Retry"
            Ninja-Property-Set $biosFirmwareUpdatesField "Exit Code: $ExitCode - Configure operation failed"
        }
        { $_ -ge 2000 -and $_ -le 2007 } {
            Ninja-Property-Set $firmwareBiosUpdateField "DriverInstall Error - Retry"
            Ninja-Property-Set $biosFirmwareUpdatesField "Exit Code: $ExitCode - Driver install operation failed"
        }
        { $_ -ge 2500 -and $_ -le 2502 } {
            Ninja-Property-Set $firmwareBiosUpdateField "Password Encryption Error"
            Ninja-Property-Set $biosFirmwareUpdatesField "Exit Code: $ExitCode - Password encryption input validation error"
        }
        { $_ -ge 3000 -and $_ -le 3005 } {
            Ninja-Property-Set $firmwareBiosUpdateField "Client Management Service Error"
            Ninja-Property-Set $biosFirmwareUpdatesField "Exit Code: $ExitCode - Dell Client Management Service error"
        }
        default {
            Ninja-Property-Set $firmwareBiosUpdateField "No Valid Exit Code (Received: $ExitCode)"
            Ninja-Property-Set $biosFirmwareUpdatesField "Exit Code: $ExitCode - Unknown error"
        }
    }
}

function Invoke-PreinstallChecks {
    $IncompatibleApps = Get-ChildItem -Path $RegPaths | Get-ItemProperty | Where-Object { $_.DisplayName -like 'Dell Update*' }
    foreach ($App in $IncompatibleApps) {
        Write-Output "Attempting to remove program: [$($App.DisplayName)]"
        $process = Start-Process -FilePath 'cmd.exe' -ArgumentList "/c $($App.UninstallString) /quiet" -Wait -NoNewWindow -PassThru
        if ($process.ExitCode -eq 0) {
            Write-Output "Successfully removed package: [$($App.DisplayName)]"
        } else {
            Write-Warning "Failed to remove package: [$($App.DisplayName)] (Exit Code: $($process.ExitCode))"
            exit 1
        }
    }
}

function Get-DownloadURL {
    $DellURL = 'https://www.dell.com/support/kbdoc/en-us/000177325/dell-command-update'
    $Headers = @{ 
        'accept'          = 'text/html'
        'accept-encoding' = 'gzip'
        'accept-language' = '*'
    }
    try {
        $DellWebPage = Invoke-RestMethod -Uri $DellURL -Headers $Headers -UseBasicParsing
        if ($DellWebPage -match '(https://www\.dell\.com.*driverId=[a-zA-Z0-9]*)') {
            $DownloadPage = Invoke-RestMethod -Uri $Matches[1] -Headers $Headers -UseBasicParsing
            if ($DownloadPage -match '(https://dl\.dell\.com.*Dell-Command-Update.*\.EXE)') {
                $url = $Matches[1]
                if ($url -like "https://dl.dell.com/*.exe") {
                    return $url
                }
            }
        }
        Write-Warning 'Failed to scrape valid URL from Dell website.'
        exit 1
    } 
    catch {
        Write-Warning "URL scraping failed: $_"
        exit 1
    }
}

function Install-DCU {
    $DownloadURL = Get-DownloadURL
    $Installer = "$env:temp\dcu-setup.exe"
    $Version = [version]($DownloadURL | Select-String '[0-9]+\.[0-9]+\.[0-9]+' | ForEach-Object { $_.Matches.Value })
    $AppName = 'Dell Command | Update for Windows Universal'
    $App = Get-ChildItem -Path $RegPaths | Get-ItemProperty | Where-Object { $_.DisplayName -like $AppName } | Select-Object -First 1
    if ($App -and [version]$App.DisplayVersion -ge $Version) {
        Write-Output "Installed version [$($App.DisplayVersion)] is up to date or newer than [$Version]. Skipping install."
    } else {
        Write-Output "Installing Dell Command Update: [$Version]"
        Invoke-WebRequest -Uri $DownloadURL -OutFile $Installer -UserAgent ([Microsoft.PowerShell.Commands.PSUserAgent]::Chrome)
        $process = Start-Process -FilePath $Installer -ArgumentList '/s' -Wait -NoNewWindow -PassThru
        if ($process.ExitCode -eq 0) {
            Write-Output "Successfully installed Dell Command Update: [$Version]"
        } else {
            Write-Warning "Failed to install Dell Command Update (Exit Code: $($process.ExitCode))"
            exit 1
        }
        if (Test-Path $Installer) {
            Remove-Item $Installer -Force
        }
    }
}

function Invoke-DCUandInstall {
    param (
        [ValidateSet('all', 'driver,application', 'firmware,bios')]
        [string]$UpdateType = 'all'
    )
    $DCU = (Resolve-Path "$env:SystemDrive\Program Files*\Dell\CommandUpdate\dcu-cli.exe" -ErrorAction SilentlyContinue).Path
    if (!$DCU) {
        Write-Warning 'Dell Command Update CLI not detected.'
        exit 1
    }
    try {
        $args = "/configure -updatesNotification=disable -userConsent=disable -scheduleAuto -silent"
        Start-Process -FilePath $DCU -ArgumentList $args -Wait -NoNewWindow -PassThru | Out-Null
        Write-Output "====== Scanning for Available Updates ======"
        $scanArgs = "/scan"
        if ($UpdateType -ne 'all') {
            $scanArgs += " -updateType=$UpdateType"
        }
        $scanArgs += " -silent"
        $outputFile = "$env:Temp\Output.txt"
        $scanProcess = Start-Process -FilePath $DCU -ArgumentList $scanArgs -Wait -NoNewWindow -PassThru -RedirectStandardOutput $outputFile
        $scanExitCode = $scanProcess.ExitCode
        $scanOutput = Get-Content -Path $outputFile -Raw
        Remove-Item -Path $outputFile -Force
        Handle-DCUExitCode -ExitCode $scanExitCode -ScanOutput $scanOutput
        if ($scanExitCode -eq 0) {
            Write-Output "====== Applying Updates ======"
            $applyArgs = "/applyUpdates -autoSuspendBitLocker=enable -reboot=disable"
            if ($UpdateType -ne 'all') {
                $applyArgs += " -updateType=$UpdateType"
            }
            $applyProcess = Start-Process -FilePath $DCU -ArgumentList $applyArgs -Wait -NoNewWindow -PassThru
            Handle-DCUExitCode -ExitCode $applyProcess.ExitCode
            if ($applyProcess.ExitCode -eq 0) {
                Write-Output "Updates applied successfully."
            } else {
                Write-Warning "Failed to apply updates (Exit Code: $($applyProcess.ExitCode))"
                exit 1
            }
        } else {
            Write-Warning "Scan failed. Updates not applied."
        }
    } catch {
        Write-Warning "Unable to apply updates using dcu-cli: $_"
        exit 1
    }
}

function Invoke-DCUScan {
    $DCU = (Resolve-Path "$env:SystemDrive\Program Files*\Dell\CommandUpdate\dcu-cli.exe" -ErrorAction SilentlyContinue).Path
    if (!$DCU) {
        Write-Warning 'Dell Command Update CLI not detected.'
        exit 1
    }
    try {
        $args = "/configure -updatesNotification=disable -userConsent=disable -scheduleAuto -silent"
        Start-Process -FilePath $DCU -ArgumentList $args -Wait -NoNewWindow -PassThru | Out-Null
        Write-Output "====== Scanning for Available Updates ======"
        $outputFile = "$env:Temp\Output.txt"
        $process = Start-Process -FilePath $DCU -ArgumentList "/scan -silent" -Wait -NoNewWindow -PassThru -RedirectStandardOutput $outputFile
        $exitCode = $process.ExitCode
        $scanOutput = Get-Content -Path $outputFile -Raw
        Remove-Item -Path $outputFile -Force
        Handle-DCUExitCode -ExitCode $exitCode -ScanOutput $scanOutput
    } catch {
        Write-Warning "Unable to scan using dcu-cli: $_"
        exit 1
    }
}

function Invoke-DCUBiosFirmwareScan {
    $DCU = (Resolve-Path "$env:SystemDrive\Program Files*\Dell\CommandUpdate\dcu-cli.exe" -ErrorAction SilentlyContinue).Path
    if (!$DCU) {
        Write-Warning 'Dell Command Update CLI not detected.'
        exit 1
    }
    try {
        $args = "/configure -updatesNotification=disable -userConsent=disable -scheduleAuto -silent"
        Start-Process -FilePath $DCU -ArgumentList $args -Wait -NoNewWindow -PassThru | Out-Null
        Write-Output "====== Scanning for BIOS and Firmware Updates ======"
        $outputFile = "$env:Temp\Output.txt"
        $process = Start-Process -FilePath $DCU -ArgumentList "/scan -updateType=bios,firmware -silent" -Wait -NoNewWindow -PassThru -RedirectStandardOutput $outputFile
        $exitCode = $process.ExitCode
        $scanOutput = Get-Content -Path $outputFile -Raw
        Remove-Item -Path $outputFile -Force
        Handle-DCUExitCode -ExitCode $exitCode -ScanOutput $scanOutput
    } catch {
        Write-Warning "Unable to scan using dcu-cli: $_"
        exit 1
    }
}

switch ($env:pleaseSelectAnOptionToRun) {
    "Install" {
        Invoke-PreinstallChecks
        Install-DCU
    }
    "Remove Incompatible Versions" {
        Invoke-PreinstallChecks
    }
    "Run Scan" {
        Invoke-DCUScan
    }
    "Run BIOS and Firmware Scan" {
        Invoke-DCUBiosFirmwareScan
    }
    "Run Scan And Install All" {
        Invoke-DCUandInstall -UpdateType 'all'
    }
    "Run Scan And Install Excluding BIOS and Firmware" {
        Invoke-DCUandInstall -UpdateType 'driver,application'
    }
    "Run Scan And Install BIOS and Firmware ONLY" {
        Invoke-DCUandInstall -UpdateType 'firmware,bios'
    }
    default {
        Write-Output "No valid option selected via environment variable 'pleaseSelectionAnOptionToRun'."
        exit 1
    }
}
