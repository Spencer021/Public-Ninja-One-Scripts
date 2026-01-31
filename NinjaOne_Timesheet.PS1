#requires -Version 5.1
#requires -Modules NinjaOneDocs

<#
.SYNOPSIS
    NinjaOne Timesheet Report Generator

.DESCRIPTION
    Generates comprehensive timesheet reports from NinjaOne ticketing data with improved UI and analytics

.NOTES
    REQUIRED CUSTOMIZATION FOR BRANDING:
    
    1. Application Window Title (Line 21)
       - What: The title bar of the GUI window
       - Change: Replace "NinjaOne – Timesheet Report" 
                 with "[Your Company Name] – Timesheet Report"
       - Purpose: User-facing branding in the application window

    OPTIONAL CUSTOMIZATION (Already Generic, but can be personalized):
    
    - Color Scheme (Line 34, 155): Currently uses #00ACEC (cyan/blue)
      Can be changed to your brand color
    
    - Output File Location: Currently saves to User Root Folder
      ex: "C:\Users\redacted\NinjaOne-Timesheet-2026-01-01_to_2026-01-31.html"
      Modify line 302 if you want a different default location
    
    - Font: Currently uses "Segoe UI"
      Can be changed for brand consistency

    - NinjaOne Instance: Default set to App.NinjaRMM.com
      Can be changed on line 282

    DEPENDENCIES:
    
    - Requires PowerShell 5.1 or higher
    - Requires NinjaOneDocs PowerShell module (must be installed separately)
    - NinjaOne API credentials must be configured on line 283, setup to use 1password vaults or Generic plain text

    FEATURES:
    
    ✅ Interactive date range selection (monthly or custom)
    ✅ Real-time progress indicators
    ✅ Detailed ticket breakdowns with time entries
    ✅ Organization and technician hour summaries
    ✅ Interactive filtering and CSV export
    ✅ Modern, responsive HTML output
    ✅ Automatic browser launch
#>

[CmdletBinding()]
param()

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

## ===========================================================
## UI: Select reporting period
## ===========================================================
$form = New-Object System.Windows.Forms.Form
## Replace For Brandable Reports
$form.Text = "NinjaOne – Timesheet Report"
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 247)
$form.Size = New-Object System.Drawing.Size(500, 400)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.AutoScaleMode = 'Font'

## ===== Header =====
$header = New-Object System.Windows.Forms.Label
$header.Text = "Select Report Period"
## Replace Colors For Brandable Reports
$header.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#00ACEC")
$header.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 14)
$header.AutoSize = $true
$header.Location = New-Object System.Drawing.Point(30, 20)
$form.Controls.Add($header)

$subHeader = New-Object System.Windows.Forms.Label
$subHeader.Text = "Generate detailed timesheet reports from NinjaOne"
$subHeader.ForeColor = [System.Drawing.Color]::Gray
$subHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$subHeader.AutoSize = $true
$subHeader.Location = New-Object System.Drawing.Point(30, 48)
$form.Controls.Add($subHeader)

## ===== Radio Buttons =====
$optMonth = New-Object System.Windows.Forms.RadioButton
$optMonth.Text = "Monthly Report"
$optMonth.Checked = $true
$optMonth.Location = New-Object System.Drawing.Point(30, 85)
$optMonth.Width = 150
$form.Controls.Add($optMonth)

$optCustom = New-Object System.Windows.Forms.RadioButton
$optCustom.Text = "Custom Date Range"
$optCustom.Location = New-Object System.Drawing.Point(200, 85)
$optCustom.Width = 180
$form.Controls.Add($optCustom)

## ===== Month Selection Group =====
$monthPanel = New-Object System.Windows.Forms.Panel
$monthPanel.Location = New-Object System.Drawing.Point(30, 120)
$monthPanel.Size = New-Object System.Drawing.Size(430, 80)
$monthPanel.BorderStyle = "None"
$form.Controls.Add($monthPanel)

$monthLabel = New-Object System.Windows.Forms.Label
$monthLabel.Text = "Month:"
$monthLabel.AutoSize = $true
$monthLabel.Location = New-Object System.Drawing.Point(0, 8)
$monthPanel.Controls.Add($monthLabel)

$monthBox = New-Object System.Windows.Forms.ComboBox
$monthBox.Location = New-Object System.Drawing.Point(90, 5)
$monthBox.Width = 150
$monthBox.DropDownStyle = 'DropDownList'

$months = [System.Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.MonthNames[0..11]
foreach ($m in $months) {
    if ($m) { [void]$monthBox.Items.Add($m) }
}
$monthBox.SelectedIndex = (Get-Date).Month - 1
$monthPanel.Controls.Add($monthBox)

$yearLabel = New-Object System.Windows.Forms.Label
$yearLabel.Text = "Year:"
$yearLabel.AutoSize = $true
$yearLabel.Location = New-Object System.Drawing.Point(0, 43)
$monthPanel.Controls.Add($yearLabel)

$yearUpDown = New-Object System.Windows.Forms.NumericUpDown
$yearUpDown.Location = New-Object System.Drawing.Point(90, 40)
$yearUpDown.Width = 100
$yearUpDown.Minimum = 2020
$yearUpDown.Maximum = 2100
$yearUpDown.Value = (Get-Date).Year
$monthPanel.Controls.Add($yearUpDown)

## ===== Custom Date Range Group =====
$datePanel = New-Object System.Windows.Forms.Panel
$datePanel.Location = New-Object System.Drawing.Point(30, 120)
$datePanel.Size = New-Object System.Drawing.Size(430, 80)
$datePanel.BorderStyle = "None"
$form.Controls.Add($datePanel)

$startLabel = New-Object System.Windows.Forms.Label
$startLabel.Text = "Start Date:"
$startLabel.AutoSize = $true
$startLabel.Location = New-Object System.Drawing.Point(0, 8)
$datePanel.Controls.Add($startLabel)

$startPicker = New-Object System.Windows.Forms.DateTimePicker
$startPicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$startPicker.Location = New-Object System.Drawing.Point(90, 5)
$startPicker.Width = 150
$datePanel.Controls.Add($startPicker)

$endLabel = New-Object System.Windows.Forms.Label
$endLabel.Text = "End Date:"
$endLabel.AutoSize = $true
$endLabel.Location = New-Object System.Drawing.Point(0, 43)
$datePanel.Controls.Add($endLabel)

$endPicker = New-Object System.Windows.Forms.DateTimePicker
$endPicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$endPicker.Location = New-Object System.Drawing.Point(90, 40)
$endPicker.Width = 150
$datePanel.Controls.Add($endPicker)

## ===== Progress Indicator =====
$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Text = ""
$progressLabel.ForeColor = [System.Drawing.Color]::Gray
$progressLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$progressLabel.AutoSize = $false
$progressLabel.Width = 440
$progressLabel.Height = 40
$progressLabel.Location = New-Object System.Drawing.Point(30, 210)
$progressLabel.Visible = $false
$form.Controls.Add($progressLabel)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(30, 245)
$progressBar.Width = 440
$progressBar.Height = 20
$progressBar.Style = 'Continuous'
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

## ===== Buttons =====
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "Generate Report"
$okButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#00ACEC")
$okButton.ForeColor = "White"
$okButton.FlatStyle = "Flat"
$okButton.FlatAppearance.BorderSize = 0
$okButton.Location = New-Object System.Drawing.Point(70, 290)
$okButton.Width = 160
$okButton.Height = 35
$okButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$okButton.Add_Click({
    $okButton.Enabled = $false
    $cancelButton.Enabled = $false
    $progressLabel.Visible = $true
    $progressBar.Visible = $true
    $progressLabel.Text = "Preparing to generate report..."
    $progressBar.Value = 0
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()

    # Process in the button click handler so form stays open
    $script:shouldProcess = $true
    $script:StartDate = $null
    $script:EndDate = $null

    # Capture selection
    if ($optMonth.Checked) {
        $selectedMonth = $monthBox.SelectedIndex + 1
        $selectedYear = [int]$yearUpDown.Value
        $script:StartDate = Get-Date -Year $selectedYear -Month $selectedMonth -Day 1
        $script:EndDate = $script:StartDate.AddMonths(1).AddDays(-1)
    }
    else {
        $script:StartDate = $startPicker.Value.Date
        $script:EndDate = $endPicker.Value.Date
    }

    # Start processing
    & $ProcessReport
})
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.FlatStyle = "Flat"
$cancelButton.Location = New-Object System.Drawing.Point(250, 290)
$cancelButton.Width = 160
$cancelButton.Height = 35
$cancelButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$cancelButton.Add_Click({ $form.Tag = "Cancel"; $form.Close() })
$form.Controls.Add($cancelButton)

## ===== Toggle UI =====
$updateUI = {
    $isMonth = $optMonth.Checked
    $monthPanel.Visible = $isMonth
    $datePanel.Visible = -not $isMonth
    $monthBox.Enabled = $yearUpDown.Enabled = $isMonth
    $startPicker.Enabled = $endPicker.Enabled = -not $isMonth
}
$optMonth.Add_CheckedChanged($updateUI)
$optCustom.Add_CheckedChanged($updateUI)
$updateUI.Invoke()

## ===== Show Form =====
$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })

# Function to update progress
$script:UpdateProgress = {
    param([string]$message, [int]$percent = -1)
    $progressLabel.Text = $message
    if ($percent -ge 0) {
        $progressBar.Value = [Math]::Min(100, $percent)
    }
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

# Define the processing script block
$script:ProcessReport = {
    try {
        & $UpdateProgress "Initializing..." 5

        ## ===========================================================
        ## API CONFIG
        ## ===========================================================
        $NinjaOneInstance = "app.ninjarmm.com"

        & $UpdateProgress "Retrieving credentials..." 10

        ## Use 1Password Login
        ## If Using Plain Text Login, comment this line out
        ##$NinjaOneClientID = op read op://<Enter Vault Info Here>
        ##$NinjaOneClientSecret = op read op://<Enter Vault Info Here>

        ## Use Plain Text Login 
        ## If Using 1Password, comment this line out
        $NinjaOneClientID = " "
        $NinjaOneClientSecret = " "
        

        if ([string]::IsNullOrEmpty($NinjaOneClientID) -or [string]::IsNullOrEmpty($NinjaOneClientSecret)) {
            throw "Failed to retrieve credentials from 1Password"
        }

        $BoardID = 2
        $OutputHtml = "$env:USERPROFILE\NinjaOne-Timesheet-$($script:StartDate.ToString('yyyy-MM-dd'))_to_$($script:EndDate.ToString('yyyy-MM-dd')).html"

        Get-Module -Name NinjaOneDocs -ErrorAction SilentlyContinue | Out-Null
        if (-not (Get-Module -Name NinjaOneDocs)) {
            & $UpdateProgress "Installing NinjaOneDocs module..." 15
            Install-Module -Name NinjaOneDocs -Scope CurrentUser -Force
        }
        & $UpdateProgress "Connecting to NinjaOne..." 20
        Import-Module NinjaOneDocs -ErrorAction Stop
        Connect-NinjaOne -NinjaOneInstance $NinjaOneInstance -NinjaOneClientID $NinjaOneClientID -NinjaOneClientSecret $NinjaOneClientSecret

## ===========================================================
## Helpers
## ===========================================================
function Convert-SecondsToHoursMins {
    param([int]$s)
    $ts = [TimeSpan]::FromSeconds([Math]::Max(0, $s))
    "{0}h {1:00}m" -f [math]::Floor($ts.TotalHours), $ts.Minutes
}

function Convert-SecondsToDecimalHours {
    param([int]$s)
    [math]::Round([Math]::Max(0, $s) / 3600, 2)
}

function Format-SecondsCombo {
    param([int]$s)
    "$((Convert-SecondsToHoursMins $s)) ($((Convert-SecondsToDecimalHours $s))h)"
}

function Get-OrgName {
    param($org)
    if ($null -eq $org) { return "Unassigned" }
    if ($org -is [string]) { return $org }
    if ($org.PSObject.Properties.Name -contains 'name') { return [string]$org.name }
    if ($org.PSObject.Properties.Name -contains 'displayName') { return [string]$org.displayName }
    return [string]$org
}

function ConvertTo-SafeHtml {
    param([string]$text)
    if ([string]::IsNullOrEmpty($text)) { return "" }
    return $text -replace "'", "&#39;" -replace '"', '&quot;' -replace '<', '&lt;' -replace '>', '&gt;'
}

function Get-TimeSecondsFromLog {
    param($log)

    $tte = $log.ticketTimeEntry
    $inSecs  = 0
    $offSecs = 0

    if ($null -ne $tte) {
        $inSecs  = [int]($tte.inHours.timeTracked  -as [int])
        $offSecs = [int]($tte.offHours.timeTracked -as [int])
    }

    if (($inSecs + $offSecs) -le 0) {
        $fallback = [int]($log.timeTracked -as [int])
        if ($fallback -le 0 -and $null -ne $tte.timeTracked) {
            $fallback = [int]($tte.timeTracked -as [int])
        }
        if ($fallback -gt 0) {
            $inSecs = $fallback
            $offSecs = 0
        }
    }

    [pscustomobject]@{
        InSeconds  = [int][Math]::Max(0, $inSecs)
        OffSeconds = [int][Math]::Max(0, $offSecs)
    }
}

function Get-BillingFromBucket {
    param($bucket)

    if ($null -eq $bucket) { return "UNKNOWN" }

    $raw = ($bucket.billing | Out-String).Trim()
    if ([string]::IsNullOrWhiteSpace($raw)) { return "UNKNOWN" }

    switch ($raw) {
        "BILLABLE"      { return "BILLABLE" }
        "NON_BILLABLE"  { return "NON_BILLABLE" }
        default         { return "UNKNOWN" }
    }
}

        $StartUnix = Get-NinjaOneTime -Seconds -Date $script:StartDate
        $EndUnix   = Get-NinjaOneTime -Seconds -Date $script:EndDate

        ## ===========================================================
        ## Technician Map
        ## ===========================================================
        & $UpdateProgress "Fetching technician list..." 25
        $Users = Invoke-NinjaOneRequest -Path 'users' -Method GET -Paginate
        $TechMap = @{}
        foreach ($u in $Users) {
            if ($u.userType -match 'TECHNICIAN') {
                $TechMap[$u.id] = "$($u.firstName) $($u.lastName)".Trim()
            }
        }
        & $UpdateProgress "Found $($TechMap.Count) technicians" 30

        ## ===========================================================
        ## Ticket Pull (Closed only in range)
        ## ===========================================================
        function Get-NinjaTickets {
            param([int]$BoardID, [string]$LastCursor = "0", [int]$PageSize = 1000)

            $Body = @"
{
  "sortBy": [ { "field": "lastUpdated", "direction": "DESC" } ],
  "pageSize": $PageSize,
  "lastCursorId": "$LastCursor"
}
"@

            Invoke-NinjaOneRequest -Path "ticketing/trigger/board/$BoardID/run" -Method POST -Body $Body
        }

        & $UpdateProgress "Fetching tickets..." 35
        $AllTickets = @()
        $LastCursor = 0
        $pageCount = 0
        do {
            $pageCount++
            & $UpdateProgress "Fetching tickets (page $pageCount)..." (35 + ($pageCount * 2))
            $resp = Get-NinjaTickets -BoardID $BoardID -LastCursor $LastCursor
            if ($resp.data.Count -eq 0) { break }
            $AllTickets += $resp.data
            $LastCursor = $resp.metadata.lastCursorId
        } while ($resp.data.Count -gt 0)

        $Tickets = $AllTickets | Where-Object {
            [double]$_.lastUpdated -ge [double]$StartUnix -and
            [double]$_.lastUpdated -le [double]$EndUnix -and
            $_.status.displayName -eq 'Closed'
        }

        & $UpdateProgress "Found $($Tickets.Count) tickets to process" 50

        ## ===========================================================
        ## Aggregate hours + build per-ticket objects
        ## ===========================================================
        $TechnicianHours = @{}
        $TicketHours     = @()
        $ticketCount = 0
        $totalTickets = $Tickets.Count

        foreach ($ticket in $Tickets) {
            $ticketCount++
            if ($ticketCount % 5 -eq 0 -or $ticketCount -eq 1) {
                $percent = 50 + [int](($ticketCount / $totalTickets) * 40)
                & $UpdateProgress "Processing ticket $ticketCount of $totalTickets..." $percent
            }

            $logs = Invoke-NinjaOneRequest -Path "ticketing/ticket/$($ticket.id)/log-entry" -Method GET

            $ticketTechBreakdown = @{}
            $ticketLogEntries    = @()
            $billSecs = 0
            $nonSecs  = 0

            foreach ($log in ($logs | Where-Object { $_.appUserContactType -eq 'TECHNICIAN' })) {

                $time = Get-TimeSecondsFromLog -log $log
                $inSecs  = $time.InSeconds
                $offSecs = $time.OffSeconds

                if (($inSecs + $offSecs) -le 0) { continue }

                $techId = $log.appUserContactId
                if (-not $techId) { continue }
                $techName = if ($TechMap.ContainsKey($techId)) { $TechMap[$techId] } else { "TechID-$techId" }

                if (-not $TechnicianHours.ContainsKey($techName)) {
                    $TechnicianHours[$techName] = @{ BillableSeconds = 0; NonBillableSeconds = 0 }
                }
                if (-not $ticketTechBreakdown.ContainsKey($techName)) {
                    $ticketTechBreakdown[$techName] = @{ BillableSeconds = 0; NonBillableSeconds = 0 }
                }

                $tte = $log.ticketTimeEntry

                if ($inSecs -gt 0) {
                    $inBilling = Get-BillingFromBucket -bucket $tte.inHours
                    if ($inBilling -eq 'BILLABLE') {
                        $TechnicianHours[$techName].BillableSeconds += $inSecs
                        $ticketTechBreakdown[$techName].BillableSeconds += $inSecs
                        $billSecs += $inSecs
                    } else {
                        $TechnicianHours[$techName].NonBillableSeconds += $inSecs
                        $ticketTechBreakdown[$techName].NonBillableSeconds += $inSecs
                        $nonSecs += $inSecs
                    }
                }

                if ($offSecs -gt 0) {
                    $offBilling = Get-BillingFromBucket -bucket $tte.offHours
                    if ($offBilling -eq 'BILLABLE') {
                        $TechnicianHours[$techName].BillableSeconds += $offSecs
                        $ticketTechBreakdown[$techName].BillableSeconds += $offSecs
                        $billSecs += $offSecs
                    } else {
                        $TechnicianHours[$techName].NonBillableSeconds += $offSecs
                        $ticketTechBreakdown[$techName].NonBillableSeconds += $offSecs
                        $nonSecs += $offSecs
                    }
                }

                $billingLabel = "UNKNOWN"
                $inBillingLbl  = Get-BillingFromBucket -bucket $tte.inHours
                $offBillingLbl = Get-BillingFromBucket -bucket $tte.offHours
                if (($inSecs -gt 0) -and ($offSecs -eq 0)) { $billingLabel = $inBillingLbl }
                elseif (($offSecs -gt 0) -and ($inSecs -eq 0)) { $billingLabel = $offBillingLbl }
                elseif ($inBillingLbl -eq $offBillingLbl) { $billingLabel = $inBillingLbl }
                else { $billingLabel = "MIXED" }

                $entryTime = [DateTimeOffset]::FromUnixTimeSeconds([math]::Floor($log.createTime)).DateTime
                $startTime = if ($tte.startDate) { [DateTimeOffset]::FromUnixTimeMilliseconds($tte.startDate).DateTime } else { $null }

                $bodyText = ($log.body -replace '<[^>]+>', '') -replace '\s+', ' '
                if ([string]::IsNullOrWhiteSpace($bodyText)) {
                    $bodyText = ($log.htmlBody -replace '<[^>]+>', '') -replace '\s+', ' '
                }

                $ticketLogEntries += [pscustomobject]@{
                    Technician = $techName
                    Type       = $log.type
                    Billing    = $billingLabel
                    Seconds    = ($inSecs + $offSecs)
                    InSeconds  = $inSecs
                    OffSeconds = $offSecs
                    StartDate  = if ($startTime) { $startTime.ToString('yyyy-MM-dd HH:mm') } else { "" }
                    EntryDate  = $entryTime.ToString('yyyy-MM-dd HH:mm')
                    Summary    = $bodyText
                }
            }

            $openDate  = [DateTimeOffset]::FromUnixTimeSeconds([math]::Floor($ticket.createTime)).DateTime
            $closeDate = if ($ticket.solvedTime) { [DateTimeOffset]::FromUnixTimeSeconds([math]::Floor($ticket.solvedTime)).DateTime } else { $null }
            $duration  = if ($openDate -and $closeDate) { New-TimeSpan -Start $openDate -End $closeDate } else { $null }
            $durationStr = if ($duration) { "{0}d {1}h {2}m" -f $duration.Days, $duration.Hours, $duration.Minutes } else { "N/A" }

            $TicketHours += [pscustomobject]@{
                TicketID           = $ticket.id
                Title              = $ticket.summary
                Organization       = (Get-OrgName $ticket.organization)
                Status             = $ticket.status.displayName
                OpenDate           = $openDate
                CloseDate          = $closeDate
                Duration           = $durationStr
                BillableSeconds    = $billSecs
                NonBillableSeconds = $nonSecs
                TechBreakdown      = $ticketTechBreakdown
                LogEntries         = $ticketLogEntries
            }
        }

        & $UpdateProgress "Building report..." 90

        ## ===========================================================
        ## Build overview tables + rows
        ## ===========================================================
        $AllTechNames = @($TechnicianHours.Keys | Sort-Object)
        $AllOrgs = @($TicketHours.Organization | Sort-Object -Unique | Where-Object { $_ })

        $OrgRows = ''
        $OrgGroups = $TicketHours | Group-Object Organization
        $totalBill = 0
        $totalNon  = 0
        foreach ($g in $OrgGroups) {
            $b = ($g.Group | Measure-Object BillableSeconds -Sum).Sum
            $n = ($g.Group | Measure-Object NonBillableSeconds -Sum).Sum
            $OrgRows += "<tr><td>$($g.Name)</td><td class='billable'>$(Format-SecondsCombo $b)</td><td class='nonbill'>$(Format-SecondsCombo $n)</td><td class='total-col'>$(Format-SecondsCombo ($b+$n))</td></tr>"
            $totalBill += $b
            $totalNon  += $n
        }
        $OrgRows += "<tr class='total-row'><td><strong>Grand Total</strong></td><td class='billable'><strong>$(Format-SecondsCombo $totalBill)</strong></td><td class='nonbill'><strong>$(Format-SecondsCombo $totalNon)</strong></td><td class='total-col'><strong>$(Format-SecondsCombo ($totalBill+$totalNon))</strong></td></tr>"

        $TechRows = ''
        foreach ($t in ($TechnicianHours.GetEnumerator() | Sort-Object { ($_.Value.BillableSeconds + $_.Value.NonBillableSeconds) } -Descending)) {
            $b = $t.Value.BillableSeconds
            $n = $t.Value.NonBillableSeconds
            $TechRows += "<tr><td>$($t.Key)</td><td class='billable'>$(Format-SecondsCombo $b)</td><td class='nonbill'>$(Format-SecondsCombo $n)</td><td class='total-col'>$(Format-SecondsCombo ($b+$n))</td></tr>"
        }
        $TechRows += "<tr class='total-row'><td><strong>Total</strong></td><td class='billable'><strong>$(Format-SecondsCombo $totalBill)</strong></td><td class='nonbill'><strong>$(Format-SecondsCombo $totalNon)</strong></td><td class='total-col'><strong>$(Format-SecondsCombo ($totalBill+$totalNon))</strong></td></tr>"

        $TicketRows = @()
        foreach ($t in $TicketHours) {

            $logHTML = ($t.LogEntries | ForEach-Object {
                "<tr><td>$($_.Technician)</td><td>$($_.Type)</td><td><span class='badge badge-$(($_.Billing).ToLower())'>$($_.Billing)</span></td><td>$(Format-SecondsCombo $_.Seconds)</td><td>$(Format-SecondsCombo $_.InSeconds)</td><td>$(Format-SecondsCombo $_.OffSeconds)</td><td>$($_.StartDate)</td><td>$($_.EntryDate)</td><td class='summary-cell'>$($_.Summary)</td></tr>"
            }) -join "`n"

            $TicketRows += @"
<tr class='ticket-row' data-org='$($t.Organization)' data-techs='$((($t.TechBreakdown.Keys) -join ';'))'>
  <td class='ticket-id'><a href='https://$NinjaOneInstance/#/ticketing/ticket/$($t.TicketID)?boardId=$BoardID' target='_blank'>#$($t.TicketID)</a></td>
  <td class='ticket-title-cell'>
    <div class='ticket-title'>$($t.Title)</div>
    <details class='ticket-details'>
      <summary>View Time Entries ($($t.LogEntries.Count))</summary>
      <div class='details-content'>
        <table class='details-table'>
          <thead><tr><th>Tech</th><th>Type</th><th>Billing</th><th>Total</th><th>In-Hours</th><th>Off-Hours</th><th>Start</th><th>Entry</th><th>Summary</th></tr></thead>
          <tbody>$logHTML</tbody>
        </table>
      </div>
    </details>
  </td>
  <td>$($t.Organization)</td>
  <td class='date-cell'>$($t.OpenDate.ToString('yyyy-MM-dd HH:mm'))</td>
  <td class='date-cell'>$(if($t.CloseDate){$t.CloseDate.ToString('yyyy-MM-dd HH:mm')}else{'N/A'})</td>
  <td class='duration-cell'>$($t.Duration)</td>
  <td><span class='badge badge-status'>$($t.Status)</span></td>
  <td class='billable'>$(Format-SecondsCombo $t.BillableSeconds)</td>
  <td class='nonbill'>$(Format-SecondsCombo $t.NonBillableSeconds)</td>
</tr>
"@
        }
        $TicketRows = $TicketRows -join "`n"

        ## ===========================================================
        ## Enhanced HTML with Modern Design
        ## ===========================================================
        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset='utf-8'/>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NinjaOne Timesheet Report - $($StartDate.ToShortDateString()) to $($EndDate.ToShortDateString())</title>
<style>
:root {
  --primary: #00ACEC;
  --primary-dark: #0699d1;
  --success: #0a7a18;
  --danger: #b50000;
  --warning: #f59e0b;
  --bg-main: #f7f8fa;
  --bg-card: #ffffff;
  --text-primary: #1f2937;
  --text-secondary: #6b7280;
  --border: #e5e7eb;
  --shadow: rgba(0, 0, 0, 0.1);
  --radius: 12px;
}

* { box-sizing: border-box; margin: 0; padding: 0; }

body {
  font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Arial, sans-serif;
  background: var(--bg-main);
  color: var(--text-primary);
  line-height: 1.6;
  padding: 20px;
}

.container {
  max-width: 1400px;
  margin: 0 auto;
}

.header {
  background: linear-gradient(135deg, var(--primary) 0%, var(--primary-dark) 100%);
  color: white;
  padding: 40px;
  border-radius: var(--radius);
  margin-bottom: 30px;
  box-shadow: 0 4px 6px var(--shadow);
}

.header h1 {
  font-size: 32px;
  font-weight: 600;
  margin-bottom: 10px;
}

.header-info {
  display: flex;
  gap: 30px;
  flex-wrap: wrap;
  margin-top: 15px;
  opacity: 0.95;
}

.header-info-item {
  display: flex;
  align-items: center;
  gap: 8px;
}

.header-info-item strong {
  font-weight: 600;
}

.stats-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
  gap: 20px;
  margin-bottom: 30px;
}

.stat-card {
  background: var(--bg-card);
  border-radius: var(--radius);
  padding: 25px;
  box-shadow: 0 2px 8px var(--shadow);
  transition: transform 0.2s, box-shadow 0.2s;
}

.stat-card:hover {
  transform: translateY(-2px);
  box-shadow: 0 4px 12px var(--shadow);
}

.stat-label {
  color: var(--text-secondary);
  font-size: 14px;
  font-weight: 500;
  margin-bottom: 8px;
  text-transform: uppercase;
  letter-spacing: 0.5px;
}

.stat-value {
  font-size: 28px;
  font-weight: 700;
  color: var(--text-primary);
}

.stat-value.billable { color: var(--success); }
.stat-value.nonbill { color: var(--danger); }

.card {
  background: var(--bg-card);
  border-radius: var(--radius);
  box-shadow: 0 2px 8px var(--shadow);
  padding: 30px;
  margin-bottom: 25px;
}

.card h2 {
  font-size: 22px;
  color: var(--text-primary);
  margin-bottom: 20px;
  padding-bottom: 12px;
  border-bottom: 3px solid var(--primary);
  display: inline-block;
}

#controls {
  background: var(--bg-card);
  border-radius: var(--radius);
  padding: 25px;
  margin-bottom: 25px;
  box-shadow: 0 2px 8px var(--shadow);
  display: flex;
  flex-wrap: wrap;
  gap: 15px;
  align-items: center;
}

#controls label {
  display: flex;
  flex-direction: column;
  gap: 6px;
  font-size: 14px;
  font-weight: 500;
  color: var(--text-primary);
}

select, button, input[type=checkbox] {
  padding: 10px 14px;
  border-radius: 8px;
  border: 1px solid var(--border);
  font-size: 14px;
  font-family: inherit;
  transition: all 0.2s;
}

select {
  background: white;
  cursor: pointer;
  min-width: 180px;
}

select:hover, select:focus {
  border-color: var(--primary);
  outline: none;
}

button {
  background: var(--primary);
  color: white;
  border: none;
  cursor: pointer;
  font-weight: 600;
  padding: 10px 20px;
}

button:hover {
  background: var(--primary-dark);
  transform: translateY(-1px);
  box-shadow: 0 2px 8px var(--shadow);
}

.checkbox-label {
  flex-direction: row !important;
  align-items: center;
  gap: 8px !important;
  cursor: pointer;
}

input[type=checkbox] {
  width: 18px;
  height: 18px;
  cursor: pointer;
  margin: 0;
}

table {
  width: 100%;
  border-collapse: collapse;
  font-size: 14px;
}

thead {
  background: linear-gradient(to bottom, #f9fafb 0%, #f3f4f6 100%);
  position: sticky;
  top: 0;
  z-index: 10;
}

th {
  text-align: left;
  padding: 14px 12px;
  font-weight: 600;
  color: var(--text-primary);
  border-bottom: 2px solid var(--border);
  white-space: nowrap;
}

td {
  padding: 12px;
  border-bottom: 1px solid var(--border);
  vertical-align: top;
}

tbody tr {
  transition: background-color 0.15s;
}

tbody tr:hover {
  background-color: #f9fafb;
}

.billable {
  color: var(--success);
  font-weight: 600;
}

.nonbill {
  color: var(--danger);
  font-weight: 600;
}

.total-col {
  color: var(--text-primary);
  font-weight: 600;
}

.total-row {
  background: #f9fafb !important;
  border-top: 2px solid var(--border);
}

.total-row td {
  padding: 16px 12px;
  font-size: 15px;
}

.ticket-id a {
  color: var(--primary);
  text-decoration: none;
  font-weight: 600;
}

.ticket-id a:hover {
  text-decoration: underline;
}

.ticket-title-cell {
  max-width: 400px;
}

.ticket-title {
  font-weight: 600;
  color: var(--text-primary);
  margin-bottom: 8px;
}

.ticket-details {
  margin-top: 8px;
}

.ticket-details summary {
  cursor: pointer;
  color: var(--primary);
  font-weight: 600;
  font-size: 13px;
  padding: 6px 10px;
  border-radius: 6px;
  display: inline-block;
  transition: background-color 0.2s;
}

.ticket-details summary:hover {
  background-color: #f0f9ff;
}

.details-content {
  margin-top: 12px;
  padding: 15px;
  background: #f9fafb;
  border-radius: 8px;
  overflow-x: auto;
  max-height: none;
  min-width: 1000px;
}

.details-table {
  font-size: 13px;
  background: white;
  width: 100%;
  table-layout: fixed;
  border: 1px solid #e5e7eb;
}

.details-table th {
  background: #f3f4f6;
  font-size: 12px;
  padding: 10px 8px;
  position: static;
  white-space: nowrap;
  border-right: 1px solid #e5e7eb;
}

.details-table td {
  padding: 10px 8px;
  white-space: normal;
  word-wrap: break-word;
  overflow-wrap: break-word;
  border-right: 1px solid #e5e7eb;
}

.details-table th:last-child,
.details-table td:last-child {
  border-right: none;
}

.details-table th:nth-child(1), .details-table td:nth-child(1) { width: 12%; } /* Tech */
.details-table th:nth-child(2), .details-table td:nth-child(2) { width: 9%; } /* Type */
.details-table th:nth-child(3), .details-table td:nth-child(3) { width: 9%; } /* Billing */
.details-table th:nth-child(4), .details-table td:nth-child(4) { width: 9%; } /* Total */
.details-table th:nth-child(5), .details-table td:nth-child(5) { width: 9%; } /* In-Hours */
.details-table th:nth-child(6), .details-table td:nth-child(6) { width: 9%; } /* Off-Hours */
.details-table th:nth-child(7), .details-table td:nth-child(7) { width: 11%; } /* Start */
.details-table th:nth-child(8), .details-table td:nth-child(8) { width: 11%; } /* Entry */
.details-table th:nth-child(9), .details-table td:nth-child(9) { width: 21%; } /* Summary */

.details-table .summary-cell {
  word-wrap: break-word;
  white-space: normal;
}

.summary-cell {
  max-width: 400px;
  word-wrap: break-word;
  white-space: normal;
}

.date-cell, .duration-cell {
  white-space: nowrap;
  font-size: 13px;
  color: var(--text-secondary);
}

.badge {
  display: inline-block;
  padding: 4px 10px;
  border-radius: 6px;
  font-size: 12px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.3px;
}

.badge-billable {
  background-color: #d1fae5;
  color: #065f46;
}

.badge-non_billable {
  background-color: #fee2e2;
  color: #991b1b;
}

.badge-unknown, .badge-mixed {
  background-color: #fef3c7;
  color: #92400e;
}

.badge-status {
  background-color: #e0e7ff;
  color: #3730a3;
}

.note {
  background: #fff7ed;
  border-left: 4px solid var(--warning);
  padding: 15px;
  border-radius: 8px;
  margin-bottom: 25px;
  font-size: 14px;
  color: var(--text-secondary);
}

.note strong {
  color: var(--text-primary);
}

.note {
  display: none;
}

@media (max-width: 768px) {
  .header-info {
    flex-direction: column;
    gap: 10px;
  }
  
  .stats-grid {
    grid-template-columns: 1fr;
  }
  
  #controls {
    flex-direction: column;
    align-items: stretch;
  }
  
  select, button {
    width: 100%;
  }
  
  table {
    font-size: 12px;
  }
  
  th, td {
    padding: 8px 6px;
  }
}

@media print {
  body {
    background: white;
  }
  
  #controls, .note {
    display: none;
  }
  
  .card {
    box-shadow: none;
    page-break-inside: avoid;
  }
  
  .ticket-details summary {
    display: none;
  }
  
  .details-content {
    display: block !important;
  }
}
</style>
</head>
<body>
<div class="container">
  <div class="header">
    <h1>🕐 NinjaOne Timesheet Report</h1>
    <div class="header-info">
      <div class="header-info-item">
        <strong>Report Period:</strong> $($StartDate.ToString('MMMM d, yyyy')) – $($EndDate.ToString('MMMM d, yyyy'))
      </div>
      <div class="header-info-item">
        <strong>Generated:</strong> $(Get-Date -Format 'MMMM d, yyyy h:mm tt')
      </div>
      <div class="header-info-item">
        <strong>Total Tickets:</strong> $($Tickets.Count)
      </div>
    </div>
  </div>

  <div class="stats-grid">
    <div class="stat-card">
      <div class="stat-label">💰 Total Billable Hours</div>
      <div class="stat-value billable">$(Convert-SecondsToDecimalHours $totalBill)h</div>
      <div style="font-size:13px;color:var(--text-secondary);margin-top:5px;">$(Convert-SecondsToHoursMins $totalBill)</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">📊 Total Non-Billable Hours</div>
      <div class="stat-value nonbill">$(Convert-SecondsToDecimalHours $totalNon)h</div>
      <div style="font-size:13px;color:var(--text-secondary);margin-top:5px;">$(Convert-SecondsToHoursMins $totalNon)</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">⏱️ Total Hours Logged</div>
      <div class="stat-value">$(Convert-SecondsToDecimalHours ($totalBill + $totalNon))h</div>
      <div style="font-size:13px;color:var(--text-secondary);margin-top:5px;">$(Convert-SecondsToHoursMins ($totalBill + $totalNon))</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">👥 Active Technicians</div>
      <div class="stat-value">$($TechnicianHours.Count)</div>
      <div style="font-size:13px;color:var(--text-secondary);margin-top:5px;">$($AllOrgs.Count) Organizations</div>
    </div>
  </div>

  <div class="note">
    <strong>ℹ️ Note:</strong> Billing classification is derived from <code>ticketTimeEntry.inHours.billing</code> and <code>ticketTimeEntry.offHours.billing</code> fields. Times are displayed as both hours/minutes and decimal hours for easy reference.
  </div>

  <div id='controls'>
    <label>
      👤 Technician
      <select id='techFilter'>
        <option value='__ALL__'>All Technicians</option>
        $($AllTechNames | ForEach-Object {"<option value='$_'>$_</option>"} | Out-String)
      </select>
    </label>
    <label>
      🏢 Organization
      <select id='orgFilter'>
        <option value='__ALL__'>All Organizations</option>
        $($AllOrgs | ForEach-Object {"<option value='$_'>$_</option>"} | Out-String)
      </select>
    </label>
    <label class='checkbox-label'>
      <input type='checkbox' id='hideNoBill' checked/>
      Hide tickets without billable time
    </label>
    <label class='checkbox-label'>
      <input type='checkbox' id='hideNoNonBill'/>
      Hide tickets without non-billable time
    </label>
    <button id='resetFilters'>🔄 Reset Filters</button>
    <button id='exportBtn'>📥 Export to CSV</button>
  </div>

  <div class='card'>
    <h2>🏢 Hours by Organization</h2>
    <table>
      <thead><tr><th>Organization</th><th>Billable</th><th>Non-Billable</th><th>Total</th></tr></thead>
      <tbody>$OrgRows</tbody>
    </table>
  </div>

  <div class='card'>
    <h2>👥 Hours by Technician</h2>
    <table>
      <thead><tr><th>Technician</th><th>Billable</th><th>Non-Billable</th><th>Total</th></tr></thead>
      <tbody>$TechRows</tbody>
    </table>
  </div>

  <div class='card'>
    <h2>🎫 Detailed Ticket Breakdown</h2>
    <table id='ticketTable'>
      <thead><tr>
        <th>ID</th><th>Title</th><th>Organization</th><th>Opened</th><th>Closed</th><th>Duration</th><th>Status</th><th>Billable</th><th>Non-Billable</th>
      </tr></thead>
      <tbody>$TicketRows</tbody>
    </table>
  </div>
</div>

<script>
const techSel = document.getElementById('techFilter');
const orgSel = document.getElementById('orgFilter');
const resetBtn = document.getElementById('resetFilters');
const exportBtn = document.getElementById('exportBtn');
const hideNoBill = document.getElementById('hideNoBill');
const hideNoNonBill = document.getElementById('hideNoNonBill');

function normalize(s) { return (s || '').toLowerCase(); }

function hasZeroCell(row, sel) {
  const el = row.querySelector(sel);
  return el && el.textContent.includes('0h 00m');
}

function applyFilters() {
  const tech = normalize(techSel.value);
  const org = normalize(orgSel.value);
  
  document.querySelectorAll('#ticketTable tbody tr.ticket-row').forEach(r => {
    const rowOrg = normalize(r.getAttribute('data-org'));
    const rowTechs = (r.getAttribute('data-techs') || '').split(';').map(normalize);
    
    const showOrg = (org === '__all__') || (rowOrg === org);
    const showTech = (tech === '__all__') || rowTechs.includes(tech);
    const hideBill = hideNoBill.checked && hasZeroCell(r, '.billable');
    const hideNon = hideNoNonBill.checked && hasZeroCell(r, '.nonbill');
    
    r.style.display = (showOrg && showTech && !hideBill && !hideNon) ? '' : 'none';
  });
}

function exportToCSV() {
  const rows = [['Ticket ID', 'Title', 'Organization', 'Opened', 'Closed', 'Duration', 'Status', 'Billable Hours', 'Non-Billable Hours', 'Time Entry Technician', 'Time Entry Type', 'Time Entry Billing', 'Time Entry Total', 'Time Entry In-Hours', 'Time Entry Off-Hours', 'Time Entry Start', 'Time Entry Created', 'Time Entry Summary']];

  document.querySelectorAll('#ticketTable tbody tr.ticket-row').forEach(row => {
    if (row.style.display !== 'none') {
      const cells = row.querySelectorAll('td');
      const ticketId = cells[0].textContent.trim().replace('#', '');
      const title = cells[1].querySelector('.ticket-title').textContent.trim().replace(/"/g, '""');
      const org = cells[2].textContent.trim().replace(/"/g, '""');
      const opened = cells[3].textContent.trim();
      const closed = cells[4].textContent.trim();
      const duration = cells[5].textContent.trim();
      const status = cells[6].querySelector('.badge') ? cells[6].querySelector('.badge').textContent.trim() : cells[6].textContent.trim();
      const billable = cells[7].textContent.trim();
      const nonBillable = cells[8].textContent.trim();

      // Check for time entries in the details table
      const detailsTable = cells[1].querySelector('.details-table tbody');

      if (detailsTable && detailsTable.rows.length > 0) {
        // Export each time entry as a separate row
        Array.from(detailsTable.rows).forEach(entryRow => {
          const entryCells = entryRow.querySelectorAll('td');
          if (entryCells.length >= 9) {
            rows.push([
              ticketId,
              title,
              org,
              opened,
              closed,
              duration,
              status,
              billable,
              nonBillable,
              entryCells[0].textContent.trim().replace(/"/g, '""'), // Tech
              entryCells[1].textContent.trim().replace(/"/g, '""'), // Type
              entryCells[2].textContent.trim().replace(/"/g, '""'), // Billing
              entryCells[3].textContent.trim().replace(/"/g, '""'), // Total
              entryCells[4].textContent.trim().replace(/"/g, '""'), // In-Hours
              entryCells[5].textContent.trim().replace(/"/g, '""'), // Off-Hours
              entryCells[6].textContent.trim().replace(/"/g, '""'), // Start
              entryCells[7].textContent.trim().replace(/"/g, '""'), // Entry
              entryCells[8].textContent.trim().replace(/"/g, '""')  // Summary
            ]);
          }
        });
      } else {
        // No time entries, just export the ticket row
        rows.push([ticketId, title, org, opened, closed, duration, status, billable, nonBillable, '', '', '', '', '', '', '', '', '']);
      }
    }
  });

  const csvContent = rows.map(row => row.map(cell => '"' + cell + '"').join(',')).join('\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.setAttribute('href', url);
  a.setAttribute('download', 'NinjaOne-Timesheet-Export-' + new Date().toISOString().split('T')[0] + '.csv');
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.URL.revokeObjectURL(url);
}

techSel.onchange = orgSel.onchange = hideNoBill.onchange = hideNoNonBill.onchange = applyFilters;
resetBtn.onclick = () => {
  techSel.value = '__ALL__';
  orgSel.value = '__ALL__';
  hideNoBill.checked = true;
  hideNoNonBill.checked = false;
  applyFilters();
};
exportBtn.onclick = exportToCSV;

document.addEventListener('DOMContentLoaded', applyFilters);
</script>
</body>
</html>
"@

        & $UpdateProgress "Generating HTML file..." 95
        Set-Content -Path $OutputHtml -Value $html -Encoding UTF8

        & $UpdateProgress "Report generated successfully!" 100
        Start-Sleep -Milliseconds 500
        $form.Close()

        try {
            $chromePathes = @(
                "C:\Program Files\Google\Chrome\Application\chrome.exe",
                "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
            )

            $chromePath = $chromePathes | Where-Object { Test-Path $_ } | Select-Object -First 1

            if ($chromePath) {
                Start-Process $chromePath -ArgumentList "`"$OutputHtml`""
            }
            else {
                Start-Process -FilePath $OutputHtml
            }
        }
        catch {
            Start-Process -FilePath $OutputHtml
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Error generating report: $errorMsg", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $form.Close()
    }
}

## ===== Display Form =====
[void]$form.ShowDialog()
