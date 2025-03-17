param (
    [switch]$Verbose
)

$approvedOU = "OU=Administrators,OU=CORP,DC=corp,DC=contoso,DC=com"
$domainRoot = "DC=corp,DC=contoso,DC=com"

$logFile = "C:\Logs\LocalAdminAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$logDir = Split-Path $logFile -Parent

if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

function Write-Log {
    param (
        [string]$Message,
        [string]$ForegroundColor = "White"
    )
    if ($Verbose) {
        Write-Host $Message -ForegroundColor $ForegroundColor
    }
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $logFile -Append
}

try {
    $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()
}
catch {
    Write-Log "This device is not domain-joined. Exiting script." -ForegroundColor Yellow
    exit 0
}

function Get-LocalAdmins {
    try {
        $admins = net localgroup administrators | 
        Where-Object { 
            $_ -notmatch "command completed" -and 
            $_ -notmatch "Alias name" -and 
            $_ -notmatch "Comment" -and 
            $_ -notmatch "Members" -and 
            $_ -notmatch "^\s*$" -and 
            $_ -notmatch "The command completed successfully" -and
            $_ -notmatch "^-+$"
        }
        return $admins
    }
    catch {
        Write-Log "Failed to get local administrators: $_" -ForegroundColor Red
        return $null
    }
}

function Is-ObjectInApprovedOU {
    param (
        [string]$objectName
    )
    try {
        if ($objectName -match "\\") {
            $objectName = $objectName.Split("\")[1]
        }
        Write-Log "DEBUG: Querying AD for sAMAccountName=$objectName" -ForegroundColor Cyan
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        $searcher.SearchRoot = [ADSI]"LDAP://$domainRoot"
        $searcher.Filter = "(&(sAMAccountName=$objectName)(|(objectCategory=group)(objectCategory=user)))"
        $searcher.PropertiesToLoad.Add("distinguishedName") | Out-Null
        $searcher.PropertiesToLoad.Add("objectCategory") | Out-Null
        $searcher.PropertiesToLoad.Add("sAMAccountName") | Out-Null
        $result = $searcher.FindOne()
        if ($null -ne $result) {
            $objectDN = $result.Properties["distinguishedName"][0]
            $objectType = $result.Properties["objectCategory"][0]
            $foundSAM = $result.Properties["sAMAccountName"][0]
            Write-Log "DEBUG: Found $foundSAM - DN: $objectDN, Type: $objectType" -ForegroundColor Cyan
            $isApproved = $objectDN -like "*$approvedOU"
            Write-Log "DEBUG: Approved check for $foundSAM = $isApproved" -ForegroundColor Cyan
            return $isApproved
        }
        Write-Log "DEBUG: $objectName not found in AD" -ForegroundColor Yellow
        return $false
    }
    catch {
        Write-Log "Error checking object '$objectName' in AD: $_" -ForegroundColor Yellow
        return $false
    }
}

Write-Log "Checking local administrators against approved Administrative OU..." -ForegroundColor Cyan

$localAdmins = Get-LocalAdmins

if ($null -eq $localAdmins) {
    Write-Log "Unable to proceed due to error getting local admins" -ForegroundColor Red
    exit 1
}

$unauthorizedObjects = @()

foreach ($admin in $localAdmins) {
    if ($admin -match "Administrator" -or $admin -match "Domain Admins") {
        Write-Log "Skipping built-in account: $admin" -ForegroundColor Gray
        continue
    }
    $isApproved = Is-ObjectInApprovedOU -objectName $admin
    if (-not $isApproved) {
        Write-Log "UNAUTHORIZED: $admin is not in $approvedOU or its sub-OUs" -ForegroundColor Red
        $unauthorizedObjects += $admin
    }
    else {
        Write-Log "Approved: $admin" -ForegroundColor Green
    }
}

Write-Log "`nSummary:" -ForegroundColor Cyan
if ($unauthorizedObjects.Count -eq 0) {
    Write-Log "No unauthorized admin objects found." -ForegroundColor Green
}
else {
    Write-Log "Found $($unauthorizedObjects.Count) unauthorized admin objects:" -ForegroundColor Red
    $unauthorizedObjects | ForEach-Object { Write-Log "- $_" -ForegroundColor Red }
}

Write-Log "Scan completed on $(Get-Date)" -ForegroundColor Cyan




########### DELETE SCRIPT ###########

$approvedOU = "OU=Administrators,OU=CORP,DC=corp,DC=contoso,DC=com"
$domainRoot = "DC=corp,DC=contoso,DC=com"

function Get-LocalAdmins {
    $admins = net localgroup administrators | 
    Where-Object { 
        $_ -notmatch "command completed" -and 
        $_ -notmatch "Alias name" -and 
        $_ -notmatch "Comment" -and 
        $_ -notmatch "Members" -and 
        $_ -notmatch "^\s*$" -and 
        $_ -notmatch "The command completed successfully" -and
        $_ -notmatch "^-+$"
    }
    return $admins
}

function Is-ObjectInApprovedOU {
    param (
        [string]$objectName
    )
    try {
        if ($objectName -match "\\") {
            $objectName = $objectName.Split("\")[1]
        }
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        $searcher.SearchRoot = [ADSI]"LDAP://$domainRoot"
        $searcher.Filter = "(&(sAMAccountName=$objectName)(|(objectCategory=group)(objectCategory=user)))"
        $searcher.PropertiesToLoad.Add("distinguishedName") | Out-Null
        $searcher.SizeLimit = 1
        $searcher.ServerTimeLimit = 5
        $result = $searcher.FindOne()
        if ($null -ne $result) {
            $objectDN = $result.Properties["distinguishedName"][0]
            return $objectDN -like "*$approvedOU"
        }
        return $false
    }
    catch {
        return $false
    }
}

$success = $true
$localAdmins = Get-LocalAdmins

if ($null -ne $localAdmins) {
    foreach ($admin in $localAdmins) {
        if ($admin -match "Administrator" -or 
            $admin -match "Domain Admins" -or 
            $admin -eq "NT AUTHORITY\SYSTEM" -or 
            $admin -match "S-1-5-") {
            continue
        }
        $isApproved = Is-ObjectInApprovedOU -objectName $admin
        if (-not $isApproved) {
            $result = net localgroup administrators "$admin" /delete 2>&1
            if ($LASTEXITCODE -ne 0 -and $result -notmatch "not a member") {
                $success = $false
            }
        }
    }
}

if ($success) { 
    exit 0 
} 
else { 
    exit 1 
}
