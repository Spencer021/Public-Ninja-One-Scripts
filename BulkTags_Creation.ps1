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
