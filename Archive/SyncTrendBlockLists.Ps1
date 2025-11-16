# Constants
$NinjaDocName = "TrendMicro Email Import Data"
$BaseUrl = "https://api.tmcas.trendmicro.com/v1/remediation/mails"
$APITokens = Ninja-Property-Docs-Get-Single $NinjaDocName APItokens

# Retrieve data from Ninja-Property-Docs-Get-Single
$BlockedSenders = @(Ninja-Property-Docs-Get-Single $NinjaDocName blockedsenders -split "`n") | ForEach-Object { $_.Trim() }
$BlockedUrls = @(Ninja-Property-Docs-Get-Single $NinjaDocName blockedurls -split "`n") | ForEach-Object { $_.Trim() }
$BlockedSha256 = @(Ninja-Property-Docs-Get-Single $NinjaDocName blockedsha256 -split "`n") | ForEach-Object { $_.Trim() }
$BlockedSha1 = @(Ninja-Property-Docs-Get-Single $NinjaDocName blockedsha1 -split "`n") | ForEach-Object { $_.Trim() }

# Function to retrieve data from the API
function Get-RemediationRules {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ApiToken
    )

    $Headers = @{
        Authorization  = "Bearer $ApiToken"
        "Content-Type" = "application/json"
    }

    try {
        Invoke-RestMethod -Uri $BaseUrl -Method 'GET' -Headers $Headers
    } catch {
        return $null
    }
}

# Initialize total counts
$TotalAddCount = 0
$TotalRemoveCount = 0

foreach ($APIToken in $APITokens) {
    # Initialize arrays
    $SendAddArray = New-Object System.Collections.ArrayList
    $URLAddArray = New-Object System.Collections.ArrayList
    $Sha256AddArray = New-Object System.Collections.ArrayList
    $Sha1AddArray = New-Object System.Collections.ArrayList

    $SendRemoveArray = New-Object System.Collections.ArrayList
    $URLRemoveArray = New-Object System.Collections.ArrayList
    $Sha256RemoveArray = New-Object System.Collections.ArrayList
    $Sha1RemoveArray = New-Object System.Collections.ArrayList

    # Fetch API data
    $REMRULES = Get-RemediationRules -ApiToken $APIToken
    if (-not $REMRULES) {
        continue
    }

    # Normalize API response
    $BlockedApiSenders = @(if ($REMRULES.Rules -and $REMRULES.Rules.senders) { $REMRULES.Rules.senders } else { @() })
    $BlockedApiUrls = @(if ($REMRULES.Rules -and $REMRULES.Rules.urls) { $REMRULES.Rules.urls } else { @() })
    $BlockedApiSha256 = @(if ($REMRULES.Rules -and $REMRULES.Rules.file256hashes) { $REMRULES.Rules.file256hashes } else { @() })
    $BlockedApiSha1 = @(if ($REMRULES.Rules -and $REMRULES.Rules.filehashes) { $REMRULES.Rules.filehashes } else { @() })

    # Compare blocked senders
    if ($null -eq $BlockedSenders -or $BlockedSenders.Count -eq 0) {
        foreach ($ApiSender in $BlockedApiSenders) {
            $SendRemoveArray.Add($ApiSender) > $null
        }
    } else {
        $SendCompare = Compare-Object -ReferenceObject $BlockedApiSenders -DifferenceObject $BlockedSenders
        foreach ($SendItem in $SendCompare) {
            if ($SendItem.SideIndicator -eq "=>") {
                $SendAddArray.Add($SendItem.InputObject) > $null
            } elseif ($SendItem.SideIndicator -eq "<=") {
                $SendRemoveArray.Add($SendItem.InputObject) > $null
            }
        }
    }

    # Compare blocked URLs
    if ($null -eq $BlockedUrls -or $BlockedUrls.Count -eq 0) {
        foreach ($ApiUrl in $BlockedApiUrls) {
            $URLRemoveArray.Add($ApiUrl) > $null
        }
    } else {
        $UrlCompare = Compare-Object -ReferenceObject $BlockedApiUrls -DifferenceObject $BlockedUrls
        foreach ($UrlItem in $UrlCompare) {
            if ($UrlItem.SideIndicator -eq "=>") {
                $URLAddArray.Add($UrlItem.InputObject) > $null
            } elseif ($UrlItem.SideIndicator -eq "<=") {
                $URLRemoveArray.Add($UrlItem.InputObject) > $null
            }
        }
    }

    # Compare blocked SHA256 hashes (file256hashes)
    if ($null -eq $BlockedSha256 -or $BlockedSha256.Count -eq 0) {
        foreach ($ApiSha256 in $BlockedApiSha256) {
            $Sha256RemoveArray.Add($ApiSha256) > $null
        }
    } else {
        $Sha256Compare = Compare-Object -ReferenceObject $BlockedApiSha256 -DifferenceObject $BlockedSha256
        foreach ($Sha256Item in $Sha256Compare) {
            if ($Sha256Item.SideIndicator -eq "=>") {
                $Sha256AddArray.Add($Sha256Item.InputObject) > $null
            } elseif ($Sha256Item.SideIndicator -eq "<=") {
                $Sha256RemoveArray.Add($Sha256Item.InputObject) > $null
            }
        }
    }

    # Compare blocked SHA1 hashes (filehashes)
    if ($null -eq $BlockedSha1 -or $BlockedSha1.Count -eq 0) {
        foreach ($ApiSha1 in $BlockedApiSha1) {
            $Sha1RemoveArray.Add($ApiSha1) > $null
        }
    } else {
        $Sha1Compare = Compare-Object -ReferenceObject $BlockedApiSha1 -DifferenceObject $BlockedSha1
        foreach ($Sha1Item in $Sha1Compare) {
            if ($Sha1Item.SideIndicator -eq "=>") {
                $Sha1AddArray.Add($Sha1Item.InputObject) > $null
            } elseif ($Sha1Item.SideIndicator -eq "<=") {
                $Sha1RemoveArray.Add($Sha1Item.InputObject) > $null
            }
        }
    }

    # Construct and send Add Rules request
    $AddRules = @{}
    if ($SendAddArray.Count -gt 0) { $AddRules.senders = $SendAddArray }
    if ($URLAddArray.Count -gt 0) { $AddRules.urls = $URLAddArray }
    if ($Sha1AddArray.Count -gt 0) { $AddRules.filehashes = $Sha1AddArray }
    if ($Sha256AddArray.Count -gt 0) { $AddRules.file256hashes = $Sha256AddArray }

    if ($AddRules.Count -gt 0) {
        try {
            Invoke-RestMethod -Uri $BaseUrl -Method 'POST' -Headers @{ Authorization = "Bearer $APIToken"; "Content-Type" = "application/json" } -Body (@{ action_type = "create"; rules = $AddRules } | ConvertTo-Json -Depth 10 -Compress)
            $TotalAddCount += $AddRules.Values | Measure-Object -Sum Count | Select-Object -ExpandProperty Sum
        } catch { }
    }

    # Construct and send Remove Rules request
    $RemoveRules = @{}
    if ($SendRemoveArray.Count -gt 0) { $RemoveRules.senders = $SendRemoveArray }
    if ($URLRemoveArray.Count -gt 0) { $RemoveRules.urls = $URLRemoveArray }
    if ($Sha1RemoveArray.Count -gt 0) { $RemoveRules.filehashes = $Sha1RemoveArray }
    if ($Sha256RemoveArray.Count -gt 0) { $RemoveRules.file256hashes = $Sha256RemoveArray }

    if ($RemoveRules.Count -gt 0) {
        try {
            Invoke-RestMethod -Uri $BaseUrl -Method 'POST' -Headers @{ Authorization = "Bearer $APIToken"; "Content-Type" = "application/json" } -Body (@{ action_type = "delete"; rules = $RemoveRules } | ConvertTo-Json -Depth 10 -Compress)
            $TotalRemoveCount += $RemoveRules.Values | Measure-Object -Sum Count | Select-Object -ExpandProperty Sum
        } catch { }
    }
}

# Output total counts
Write-Output "Total Adds: $TotalAddCount"
Write-Output "Total Removes: $TotalRemoveCount"
