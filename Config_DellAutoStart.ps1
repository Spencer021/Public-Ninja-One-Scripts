if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Output "Exiting script, please run as Administrator!"
    exit
}

if ((Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).Manufacturer -notlike "*Dell*") { Write-Output "This script only runs on Dell devices, exiting"; exit }

$AutoOnDay = "EveryDay"
$AutoOnHr = 22
$AutoOnMin = 0
$StagingPath = "C:\RMMTools"
$ModuleName = "DellBIOSProvider"

if (-not (Test-Path -Path $StagingPath)) {
    Write-Output "Creating RMMTools Directory for Staging"
    try {
        New-Item -Path 'C:\' -Name RMMTools -ItemType Directory -ErrorAction Stop | Out-Null
    } catch {
        Write-Error "Failed to create directory $StagingPath : $($_.Exception.Message)"
        exit
    }
} else {
    Write-Output "RMMTools directory already exists"
}

$getModule = Get-Module -Name $ModuleName -ListAvailable | Where-Object { $_.Version -eq "2.9.0" }
if (-not $getModule) {
    Write-Output "Installing DellBIOSProvider Module version 2.9.0"
    try {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
        Find-Module -Name $ModuleName -Repository PSGallery -RequiredVersion 2.9.0 -ErrorAction Stop | Save-Module -Path $StagingPath -Force -ErrorAction Stop
    } catch {
        Write-Error "Failed to install DellBIOSProvider: $($_.Exception.Message)"
        exit
    }
} else {
    Write-Output "DellBIOSProvider version 2.9.0 is already installed"
}

try {
    $env:PSModulePath = $env:PSModulePath + ";$StagingPath"
    Import-Module -Name $ModuleName -RequiredVersion 2.9.0 -ErrorAction Stop
} catch {
    Write-Error "Failed to import DellBIOSProvider: $($_.Exception.Message)"
    exit
}

try {
    $isAdminPassSet = Get-Item -Path DellSmbios:\Security\IsAdminPasswordSet -ErrorAction Stop | Select-Object -ExpandProperty CurrentValue
    $isSystemPassSet = Get-Item -Path DellSmbios:\Security\IsSystemPasswordSet -ErrorAction Stop | Select-Object -ExpandProperty CurrentValue
    if ($isAdminPassSet -eq $true -or $isSystemPassSet -eq $true) {
        Write-Output "BIOS Admin or System Password is set. Please remove it manually via BIOS setup or provide the password."
        exit
    } else {
        Write-Output "No BIOS password set, proceeding with Auto-On configuration"
    }

    $currentAutoOn = Get-Item -Path DellSmbios:\PowerManagement\AutoOn -ErrorAction Stop
    if ($currentAutoOn.CurrentValue -ne $AutoOnDay) {
        Set-Item -Path DellSmbios:\PowerManagement\AutoOn $AutoOnDay -ErrorAction Stop
        Write-Output "Set AutoOn to $AutoOnDay"
    } else {
        Write-Output "AutoOn is already set to $AutoOnDay"
    }

    $currentHour = Get-Item -Path DellSmbios:\PowerManagement\AutoOnHr -ErrorAction Stop
    if ($currentHour.CurrentValue -ne $AutoOnHr) {
        Set-Item -Path DellSmbios:\PowerManagement\AutoOnHr $AutoOnHr -ErrorAction Stop
        Write-Output "Set AutoOnHr to $AutoOnHr"
    } else {
        Write-Output "AutoOnHr is already set to $AutoOnHr"
    }

    $currentMinute = Get-Item -Path DellSmbios:\PowerManagement\AutoOnMn -ErrorAction Stop
    if ($currentMinute.CurrentValue -ne $AutoOnMin) {
        Set-Item -Path DellSmbios:\PowerManagement\AutoOnMn $AutoOnMin -ErrorAction Stop
        Write-Output "Set AutoOnMn to $AutoOnMin"
    } else {
        Write-Output "AutoOnMn is already set to $AutoOnMin"
    }

    Write-Output "`nFinal Settings:"
    Get-Item -Path DellSmbios:\PowerManagement\AutoOn | ForEach-Object { Write-Output "AutoOn: $($_.CurrentValue)" }
    Get-Item -Path DellSmbios:\PowerManagement\AutoOnHr | ForEach-Object { Write-Output "AutoOnHr: $($_.CurrentValue)" }
    Get-Item -Path DellSmbios:\PowerManagement\AutoOnMn | ForEach-Object { Write-Output "AutoOnMn: $($_.CurrentValue)" }

    Write-Output "`nEnsure the system is connected to AC power for Auto-On to function."
} catch {
    Write-Error "Error during Auto-On configuration: $($_.Exception.Message)"
    exit
}
