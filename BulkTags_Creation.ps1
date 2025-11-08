<#
.SYNOPSIS
    Mass-creates tags in NinjaOne using the public API and OAuth2 authentication.
    Simplifies bulk tag creation from a CSV file.

.DESCRIPTION
    This script automates the process of creating multiple tags within NinjaOne.
    It uses the official OAuth2 client credentials flow to authenticate and calls
    the /v2/tag endpoint for each tag listed in the CSV file.

    Key features include:
    - Authenticates automatically using the NinjaOne API (no manual token entry).
    - Reads a simple CSV file with columns: Name, Description.
    - Posts each entry as a new tag in NinjaOne.

    Ideal for initial setup, standardizing environments, or migrating tag structures.
    Minimal setup required — just fill in your Client ID, Secret, and CSV path.

.PARAMETERS
    $BaseUrl:       Your NinjaOne region base URL (e.g., https://app.ninjarmm.com)
    $ClientId:      API Client ID from NinjaOne.
    $ClientSecret:  API Client Secret from NinjaOne.
    $CsvPath:       Path to your CSV file containing tag names and descriptions.

.CSV
    | Name     | Description                   |
    |----------|-------------------------------|
    | Critical | Devices needing priority       |
    | Backup   | Endpoints with backup enabled  |

.AUTHOR
    By: Spencer A. Heath
    DATE: 08 November 2025

.GITHUB
    https://github.com/Sp-e-n-c-er

.NOTES
    **Disclaimer**: This script is provided “as is” under the MIT License.
    Use at your own risk. The author assumes no responsibility for damages or issues.
    Licensed under MIT: © Spencer A. Heath. Permission granted to use, copy, modify,
    and distribute freely, provided this notice remains intact.
    See https://opensource.org/licenses/MIT for details.
#>

$BaseUrl      = "https://app.ninjarmm.com"
$ClientId     = ""
$ClientSecret = ""
$CsvPath      = "C:\Temp\test.csv"

$AuthBody = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "monitoring management"
}

$TokenResponse = Invoke-RestMethod -Method Post -Uri "$BaseUrl/ws/oauth/token" -ContentType "application/x-www-form-urlencoded" -Body $AuthBody

$AccessToken = $TokenResponse.access_token

$Headers = @{
    "Authorization" = "Bearer $AccessToken"
    "Accept"        = "application/json"
}

$tags = Import-Csv -Path $CsvPath

foreach ($tag in $tags) {
    $Body = @{
        name        = $tag.Name
        description = $tag.Description
    } | ConvertTo-Json

    Invoke-RestMethod -Method Post -Uri "$BaseUrl/v2/tag" -Headers $Headers -ContentType "application/json" -Body $Body

    Write-Host "Created tag: $($tag.Name)"
}
