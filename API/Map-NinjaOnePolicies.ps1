$devices = @()
#########################################
#     CUSTOM BRANDING FOR NINJAONE      #
#########################################
$OrganizationName = "YOUR ORG NAME HERE"    #########################################
$Text = "NinjaOne Hierarchy Policy Report"  #     CUSTOM BRANDING FOR NINJAONE      #
$date = Get-Date -Format "yyyy-MM-dd"       #########################################
#########################################
#     CUSTOM BRANDING FOR NINJAONE      #
#########################################
$desktopPath = [Environment]::GetFolderPath("Desktop")
$all = Get-NinjaOnePolicies
$devices = Get-NinjaOneDevices
$TotalPolicyCount = ($all | Measure-Object).Count

# Custom CSS for dark mode styling with custom branding and centered banner
$customCSS = @"
<style>
    body { font-family: Arial, sans-serif; background-color: #2c2c2c; color: #e0e0e0; }
    header {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        text-align: center;
        padding: 20px;
        background-color: #444;
        border-radius: 8px;
        margin: 0 auto 20px auto; /* Centers the header box */
        max-width: 800px; /* Sets a maximum width for the header */
    }
    h1 { font-size: 2em; color: #00bfff; margin: 0; }
    .org-info { font-size: 1.2em; color: #bbb; margin-top: 10px; }
    ul { list-style-type: none; margin-left: 20px; position: relative; }
    li { margin: 10px 0; padding: 10px; background-color: #444; border-radius: 8px; position: relative; }
    li::before { content: ''; position: absolute; top: -10px; left: -20px; width: 2px; height: 100%; background-color: #00bfff; }
    li::after { content: ''; position: absolute; top: 10px; left: -20px; width: 10px; height: 2px; background-color: #00bfff; }
    .policy-name { font-size: 1.2em; color: #ffffff; font-weight: bold; }
    .node-class { font-style: italic; color: #bbb; }
    .device-count { font-size: 1em; color: #aaa; }
</style>
"@

# Create a dictionary for device counts per policy
$deviceCounts = @{ }
foreach ($device in $devices) {
    $policyId = if ($device.policyId) { $device.policyId } else { $device.rolePolicyId }
    if ($policyId) {
        if ($deviceCounts.ContainsKey($policyId)) {
            $deviceCounts[$policyId]++
        } else {
            $deviceCounts[$policyId] = 1
        }
    }
}

# Recursive function to build policy tree
function Get-PolicyTree {
    param ($policies, $parentId = $null)
    $children = $policies | Where-Object { $_.parentPolicyId -eq $parentId }
    foreach ($child in $children) {
        $child | Add-Member -MemberType NoteProperty -Name 'Children' -Value @() -Force
        $policyId = if ($child.id) { $child.id } else { $child.rolePolicyId }
        $deviceCount = if ($deviceCounts.ContainsKey($policyId)) { $deviceCounts[$policyId] } else { 0 }
        $devicesApplied = if ($deviceCount -gt 0) { "True" } else { "False" }
        $child | Add-Member -MemberType NoteProperty -Name 'DeviceCount' -Value $deviceCount -Force
        $child | Add-Member -MemberType NoteProperty -Name 'DevicesApplied' -Value $devicesApplied -Force
        $child.Children = Get-PolicyTree -policies $policies -parentId $child.id
    }
    return $children
}

# Generate policy tree
$policyTree = Get-PolicyTree -policies $all -parentId $null

# Recursive function to convert policy tree to HTML
function Convert-PolicyTreeToHTML {
    param ($policyTree)
    if ($policyTree.Count -eq 0) { return "" }
    $html = "<ul>"
    foreach ($policy in $policyTree) {
        $html += "<li><span class='policy-name'><i class='fa fa-sitemap'></i> $($policy.name)</span><br/>"
        $html += "<span class='node-class'>Node Class: $($policy.nodeClass)</span><br/>"
        $html += "<span class='device-count'>Devices Applied: $($policy.DevicesApplied) (Count: $($policy.DeviceCount))</span>"
        if ($policy.Children.Count -gt 0) {
            $html += Convert-PolicyTreeToHTML -policyTree $policy.Children
        }
        $html += "</li>"
    }
    $html += "</ul>"
    return $html
}

# Generate HTML content from the policy tree
$policyHTML = Convert-PolicyTreeToHTML -policyTree $policyTree

# Create HTML report using PSWriteHTML
Import-Module PSWriteHTML

New-HTML -TitleText "NinjaOne Policy Report" -FilePath "$desktopPath\PolicyReport.html" {

    # Add organization branding at the top
    New-HTMLContent {
        "<header>
            <h1>$OrganizationName</h1>
            <div class='org-info'>$Text</div>
            <div class='org-info'>Generated on: $date</div>
            <div class='org-info'>Total Policies: $TotalPolicyCount</div>
        </header>"
    }
    # Add policy tree
    New-HTMLContent {
        $customCSS
        $policyHTML
    }
}
Write-Host "Policy report generated at: $desktopPath\PolicyReport.html" -ForegroundColor Green
start msedge $desktopPath\PolicyReport.html
