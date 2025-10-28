<#
.SYNOPSIS
Generates a NinjaOne Timesheet Report with technician billable and non-billable hour summaries.

.DESCRIPTION
This PowerShell script connects to the NinjaOne API using stored credentials (via 1Password CLI)
to generate detailed technician timesheet reports. The script allows selection between
Monthly Mode (Month + Year) or a Custom Date Range and outputs a formatted HTML report
with organization and technician summaries.

.OUTPUTS
An HTML file saved to the user's Documents folder (auto-opens in Chrome).

.REQUIREMENTS
- PowerShell 5.1
- NinjaOneDocs module (auto-installs if missing)

.Notes
- Credential input can be changed on line 148

.AUTHOR
Spencer Heath (https://github.com/Sp-e-n-c-er)

.LICENSE
This script is provided "as is" without warranty of any kind, either expressed or implied.
I (Spencer) assumes no liability or responsibility for any damage, loss, or issues arising
from the use, modification, or distribution of this script. Use at your own risk.

#>



#requires -Version 5.1
#requires -Modules NinjaOneDocs

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "NinjaOne – Timesheet Report"
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 247)
$form.Size = New-Object System.Drawing.Size(460, 360)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.AutoScaleMode = 'Font' 

# ===== Header =====
$header = New-Object System.Windows.Forms.Label
$header.Text = "Select the report period"
$header.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#00ACEC")
$header.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 12)
$header.AutoSize = $true
$header.Location = New-Object System.Drawing.Point(30, 25)
$form.Controls.Add($header)

# ===== Radio Buttons =====
$optMonth = New-Object System.Windows.Forms.RadioButton
$optMonth.Text = "Monthly Mode"
$optMonth.Checked = $true
$optMonth.Location = New-Object System.Drawing.Point(30, 70)
$form.Controls.Add($optMonth)

$optCustom = New-Object System.Windows.Forms.RadioButton
$optCustom.Text = "Custom Date Range"
$optCustom.Location = New-Object System.Drawing.Point(180, 70)
$form.Controls.Add($optCustom)

# ===== Month =====
$monthLabel = New-Object System.Windows.Forms.Label
$monthLabel.Text = "Month:"
$monthLabel.AutoSize = $true
$monthLabel.Location = New-Object System.Drawing.Point(30, 110)
$form.Controls.Add($monthLabel)

$monthBox = New-Object System.Windows.Forms.ComboBox
$monthBox.Location = New-Object System.Drawing.Point(120, 105)
$monthBox.Width = 220
$monthBox.DropDownStyle = 'DropDownList'

$months = [System.Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.MonthNames[0..11]
foreach ($m in $months) {
    if ($m) { [void]$monthBox.Items.Add($m) }
}
$monthBox.SelectedIndex = (Get-Date).Month - 1
$form.Controls.Add($monthBox)

# ===== Year =====
$yearLabel = New-Object System.Windows.Forms.Label
$yearLabel.Text = "Year:"
$yearLabel.AutoSize = $true
$yearLabel.Location = New-Object System.Drawing.Point(30, 150)
$form.Controls.Add($yearLabel)

$yearUpDown = New-Object System.Windows.Forms.NumericUpDown
$yearUpDown.Location = New-Object System.Drawing.Point(120, 145)
$yearUpDown.Width = 100
$yearUpDown.Minimum = 2020
$yearUpDown.Maximum = 2100
$yearUpDown.Value = (Get-Date).Year
$form.Controls.Add($yearUpDown)

# ===== Start Date =====
$startLabel = New-Object System.Windows.Forms.Label
$startLabel.Text = "Start Date:"
$startLabel.AutoSize = $true
$startLabel.Location = New-Object System.Drawing.Point(30, 190)
$form.Controls.Add($startLabel)

$startPicker = New-Object System.Windows.Forms.DateTimePicker
$startPicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$startPicker.Location = New-Object System.Drawing.Point(120, 185)
$form.Controls.Add($startPicker)

# ===== End Date =====
$endLabel = New-Object System.Windows.Forms.Label
$endLabel.Text = "End Date:"
$endLabel.AutoSize = $true
$endLabel.Location = New-Object System.Drawing.Point(30, 225)
$form.Controls.Add($endLabel)

$endPicker = New-Object System.Windows.Forms.DateTimePicker
$endPicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$endPicker.Location = New-Object System.Drawing.Point(120, 220)
$form.Controls.Add($endPicker)

# ===== Buttons =====
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "Generate Report"
$okButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#00ACEC")
$okButton.ForeColor = "White"
$okButton.FlatStyle = "Flat"
$okButton.FlatAppearance.BorderSize = 0
$okButton.Location = New-Object System.Drawing.Point(60, 270)
$okButton.Width = 140
$okButton.Add_Click({ $form.Tag = "OK"; $form.Close() })
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.FlatStyle = "Flat"
$cancelButton.Location = New-Object System.Drawing.Point(220, 270)
$cancelButton.Width = 140 
$cancelButton.Add_Click({ $form.Tag = "Cancel"; $form.Close() })
$form.Controls.Add($cancelButton)

# ===== Toggle UI =====
$updateUI = {
    $isMonth = $optMonth.Checked
    $monthBox.Enabled = $yearUpDown.Enabled = $isMonth
    $startPicker.Enabled = $endPicker.Enabled = -not $isMonth
}
$optMonth.Add_CheckedChanged($updateUI)
$optCustom.Add_CheckedChanged($updateUI)
$updateUI.Invoke()

# ===== Show Form =====
$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
if ($form.Tag -ne "OK") { exit }

# ===== Capture Selection =====
if ($optMonth.Checked) {
    $selectedMonth = $monthBox.SelectedIndex + 1
    $selectedYear = [int]$yearUpDown.Value
    $StartDate = Get-Date -Year $selectedYear -Month $selectedMonth -Day 1
    $EndDate = $StartDate.AddMonths(1).AddDays(-1)
}
else {
    $StartDate = $startPicker.Value.Date
    $EndDate = $endPicker.Value.Date
}

Write-Host "Generating report for period $($StartDate.ToShortDateString()) – $($EndDate.ToShortDateString())..."


## ===========================================================
## API CONFIG
## ===========================================================
$NinjaOneInstance     = "app.ninjarmm.com"
$NinjaOneClientID     = op read op://'VaultName'/'ItemName'/Username
$NinjaOneClientSecret = op read op://'VaultName'/'ItemName'/Password
$BoardID              = 2
$OutputHtml           = "$env:USERPROFILE\Documents\NinjaOne-Timesheet-$($StartDate.ToString('yyyy-MM')).html"

Get-Module -Name NinjaOneDocs -ErrorAction SilentlyContinue | Out-Null
if (-not (Get-Module -Name NinjaOneDocs)) {
    Write-Host "NinjaOneDocs module not found. Installing from PSGallery..."
    Install-Module -Name NinjaOneDocs -Scope CurrentUser -Force
}
Import-Module NinjaOneDocs -ErrorAction Stop
Connect-NinjaOne -NinjaOneInstance $NinjaOneInstance -NinjaOneClientID $NinjaOneClientID -NinjaOneClientSecret $NinjaOneClientSecret

## ===========================================================
## Helpers
## ===========================================================
function Convert-SecondsToHoursMins { 
    param([int]$s)
    $ts=[TimeSpan]::FromSeconds([Math]::Max(0,$s));"{0}h {1:00}m" -f [math]::Floor($ts.TotalHours),$ts.Minutes 
}
function Convert-SecondsToDecimalHours { 
    param([int]$s) 
    [math]::Round([Math]::Max(0,$s)/3600,2) 
}
function Format-SecondsCombo { 
    param([int]$s) 
    "$((Convert-SecondsToHoursMins $s)) ($((Convert-SecondsToDecimalHours $s)))" 
}

$StartUnix = Get-NinjaOneTime -Seconds -Date $StartDate
$EndUnix   = Get-NinjaOneTime -Seconds -Date $EndDate

## ===========================================================
## Technician Map
## ===========================================================
$Users = Invoke-NinjaOneRequest -Path 'users' -Method GET -Paginate
$TechMap = @{}
foreach ($u in $Users) {
    if ($u.userType -match 'TECHNICIAN') {
        $TechMap[$u.id] = "$($u.firstName) $($u.lastName)"
    }
}

## ===========================================================
## Ticket Pull (Closed only in range)
## ===========================================================
function Get-NinjaTickets {
    param([int]$BoardID,[string]$LastCursor="0",[int]$PageSize=1000)
    $Body = @"
{
  "sortBy": [ { "field": "lastUpdated", "direction": "DESC" } ],
  "pageSize": $PageSize,
  "lastCursorId": "$LastCursor"
}
"@
    Invoke-NinjaOneRequest -Path "ticketing/trigger/board/$BoardID/run" -Method POST -Body $Body
}

$AllTickets=@();$LastCursor=0
do {
    $resp=Get-NinjaTickets -BoardID $BoardID -LastCursor $LastCursor
    if ($resp.data.Count -eq 0){break}
    $AllTickets+=$resp.data
    $LastCursor=$resp.metadata.lastCursorId
} while ($resp.data.Count -gt 0)

$Tickets=$AllTickets|Where-Object{
    [double]$_.lastUpdated -ge [double]$StartUnix -and
    [double]$_.lastUpdated -le [double]$EndUnix -and
    $_.status.displayName -eq 'Closed'
}

Write-Host "Found $($Tickets.Count) closed tickets."

$TechnicianHours=@{}
$TicketHours=@()

foreach ($ticket in $Tickets) {
    $logs=Invoke-NinjaOneRequest -Path "ticketing/ticket/$($ticket.id)/log-entry" -Method GET
    $ticketTechBreakdown=@{};$ticketLogEntries=@()
    $billSecs=0;$nonSecs=0

    foreach ($log in $logs | Where-Object { $_.appUserContactType -eq 'TECHNICIAN' -and $_.type -eq 'COMMENT'}) {
        $sec=[int]($log.timeTracked -as [int])
        if ($sec -le 0 -and $null -ne $log.ticketTimeEntry.timeTracked){$sec=[int]$log.ticketTimeEntry.timeTracked}
        if ($sec -le 0){continue}
        $billing=if($log.ticketTimeEntry.billing){$log.ticketTimeEntry.billing}else{"UNKNOWN"}
        $techId=$log.appUserContactId;if(-not $techId){continue}
        $techName=if($TechMap.ContainsKey($techId)){$TechMap[$techId]}else{"TechID-$techId"}

        if(-not $TechnicianHours.ContainsKey($techName)){$TechnicianHours[$techName]=@{BillableSeconds=0;NonBillableSeconds=0}}
        if($billing -eq 'BILLABLE'){$TechnicianHours[$techName].BillableSeconds+=$sec;$billSecs+=$sec}else{$TechnicianHours[$techName].NonBillableSeconds+=$sec;$nonSecs+=$sec}

        if(-not $ticketTechBreakdown.ContainsKey($techName)){$ticketTechBreakdown[$techName]=@{BillableSeconds=0;NonBillableSeconds=0}}
        if($billing -eq 'BILLABLE'){$ticketTechBreakdown[$techName].BillableSeconds+=$sec}else{$ticketTechBreakdown[$techName].NonBillableSeconds+=$sec}

        $entryTime=[DateTimeOffset]::FromUnixTimeSeconds([math]::Floor($log.createTime)).DateTime
        $startTime=if($log.ticketTimeEntry.startDate){[DateTimeOffset]::FromUnixTimeMilliseconds($log.ticketTimeEntry.startDate).DateTime}else{$null}
        $bodyText=($log.body -replace '<[^>]+>', '') -replace '\s+', ' '
        $ticketLogEntries+=[pscustomobject]@{
            Technician=$techName; Billing=$billing; Seconds=$sec;
            StartDate=if($startTime){$startTime.ToString('yyyy-MM-dd HH:mm')};
            EntryDate=$entryTime.ToString('yyyy-MM-dd HH:mm');
            Summary=$bodyText
        }
    }

    $openDate=[DateTimeOffset]::FromUnixTimeSeconds([math]::Floor($ticket.createTime)).DateTime
    $closeDate=if($ticket.solvedTime){[DateTimeOffset]::FromUnixTimeSeconds([math]::Floor($ticket.solvedTime)).DateTime}else{$null}
    $duration=if($openDate -and $closeDate){New-TimeSpan -Start $openDate -End $closeDate}else{$null}
    $durationStr=if($duration){"{0}d {1}h {2}m" -f $duration.Days,$duration.Hours,$duration.Minutes}else{"N/A"}

    $TicketHours+=[pscustomobject]@{
        TicketID=$ticket.id; Title=$ticket.summary; Organization=$ticket.organization;
        Status=$ticket.status.displayName;
        OpenDate=$openDate; CloseDate=$closeDate; Duration=$durationStr;
        BillableSeconds=$billSecs; NonBillableSeconds=$nonSecs;
        TechBreakdown=$ticketTechBreakdown; LogEntries=$ticketLogEntries
    }
}

## ===========================================================
## Build overview tables + rows
## ===========================================================
$AllTechNames = @($TechnicianHours.Keys | Sort-Object)
$AllOrgs = @($TicketHours.Organization | Sort-Object -Unique | Where-Object { $_ })

# Org totals
$OrgRows = ''
$OrgGroups = $TicketHours | Group-Object Organization
$totalBill = 0; $totalNon = 0
foreach ($g in $OrgGroups) {
    $b = ($g.Group | Measure-Object BillableSeconds -Sum).Sum
    $n = ($g.Group | Measure-Object NonBillableSeconds -Sum).Sum
    $OrgRows += "<tr><td>$($g.Name)</td><td class='billable'>$(Format-SecondsCombo $b)</td><td class='nonbill'>$(Format-SecondsCombo $n)</td><td>$(Format-SecondsCombo ($b+$n))</td></tr>"
    $totalBill += $b; $totalNon += $n
}
$OrgRows += "<tr class='total'><td>Grand Total</td><td class='billable'>$(Format-SecondsCombo $totalBill)</td><td class='nonbill'>$(Format-SecondsCombo $totalNon)</td><td>$(Format-SecondsCombo ($totalBill+$totalNon))</td></tr>"

# Tech totals
$TechRows = ''
foreach ($t in $TechnicianHours.GetEnumerator() | Sort-Object {($_.Value.BillableSeconds + $_.Value.NonBillableSeconds)} -Descending) {
    $b=$t.Value.BillableSeconds;$n=$t.Value.NonBillableSeconds
    $TechRows += "<tr><td>$($t.Key)</td><td class='billable'>$(Format-SecondsCombo $b)</td><td class='nonbill'>$(Format-SecondsCombo $n)</td><td>$(Format-SecondsCombo ($b+$n))</td></tr>"
}
$TechRows += "<tr class='total'><td>Total</td><td class='billable'>$(Format-SecondsCombo $totalBill)</td><td class='nonbill'>$(Format-SecondsCombo $totalNon)</td><td>$(Format-SecondsCombo ($totalBill+$totalNon))</td></tr>"

# Ticket rows (prebuild as string)
$TicketRows = @()
foreach ($t in $TicketHours) {
    $logHTML = ($t.LogEntries | ForEach-Object {
        "<tr><td>$($_.Technician)</td><td>$($_.Billing)</td><td>$(Format-SecondsCombo $_.Seconds)</td><td>$($_.StartDate)</td><td>$($_.EntryDate)</td><td>$($_.Summary)</td></tr>"
    }) -join "`n"

    $TicketRows += @"
<tr class='ticket-row' data-org='$($t.Organization)' data-techs='$((($t.TechBreakdown.Keys) -join ';'))'>
  <td><a href='https://app.ninjarmm.com/#/ticketing/ticket/$($t.TicketID)?boardId=$BoardID' target='_blank'>$($t.TicketID)</a></td>
  <td><span class='ticket-title'>$($t.Title)</span>
    <details><summary>Details</summary>
      <div><table><thead><tr><th>Tech</th><th>Billing</th><th>Time</th><th>Start</th><th>Entry</th><th>Summary</th></tr></thead><tbody>
        $logHTML
      </tbody></table></div>
    </details>
  </td>
  <td>$($t.Organization)</td>
  <td>$($t.OpenDate.ToString('yyyy-MM-dd HH:mm'))</td>
  <td>$($t.CloseDate.ToString('yyyy-MM-dd HH:mm'))</td>
  <td>$($t.Duration)</td>
  <td>$($t.Status)</td>
  <td class='billable'>$(Format-SecondsCombo $t.BillableSeconds)</td>
  <td class='nonbill'>$(Format-SecondsCombo $t.NonBillableSeconds)</td>
</tr>
"@
}
$TicketRows = $TicketRows -join "`n"

## ===========================================================
## HTML
## ===========================================================
$html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8'/>
<title>NinjaOne Timesheet Report</title>
<style>
body{font-family:'Segoe UI',Roboto,Arial;margin:30px;background:#f7f8fa;color:#222;}
h1{font-size:28px;color:#00ACEC;margin-bottom:5px;}
h2{margin-top:40px;border-left:6px solid #00ACEC;padding-left:10px;color:#444;}
.card{background:#fff;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:20px;margin-bottom:25px;}
table{width:100%;border-collapse:collapse;margin-top:10px;font-size:14px;}
th{position:sticky;top:0;background:#f1f3f5;text-align:left;border-bottom:2px solid #ddd;padding:8px;}
td{padding:8px;border-bottom:1px solid #eee;vertical-align:top;}
tr:hover{background:#f9fafc;}
.billable{color:#0a7a18;font-weight:600;}
.nonbill{color:#b50000;}
.total td{font-weight:bold;background:#f3f3f3;}
#controls{display:flex;flex-wrap:wrap;gap:12px;margin-bottom:15px;align-items:center;}
select,button,input[type=checkbox]{padding:6px 10px;border-radius:6px;border:1px solid #ccc;font-size:13px;}
button{background:#00ACEC;color:#fff;border:none;cursor:pointer;}
button:hover{background:#0699d1;}
a{color:#00ACEC;text-decoration:none;}
a:hover{text-decoration:underline;}
details summary{cursor:pointer;color:#00ACEC;font-weight:600;}
details div{margin-top:8px;}
.ticket-title{font-weight:700;color:#111;}
label.checkbox-label{display:flex;align-items:center;gap:6px;font-size:13px;}
</style>
</head>
<body>
<h1>NinjaOne Timesheet Report</h1>
<p><b>Period:</b> $($StartDate.ToShortDateString()) – $($EndDate.ToShortDateString())</p>
<p><b>Generated:</b> $(Get-Date)</p>

<div id='controls' class='card'>
  <label>Technician:
    <select id='techFilter'>
      <option value='__ALL__'>All Technicians</option>
      $($AllTechNames | ForEach-Object {"<option value='$_'>$_</option>"})
    </select>
  </label>
  <label>Organization:
    <select id='orgFilter'>
      <option value='__ALL__'>All Organizations</option>
      $($AllOrgs | ForEach-Object {"<option value='$_'>$_</option>"})
    </select>
  </label>
  <label class='checkbox-label'><input type='checkbox' id='hideNoBill' checked/> Hide all without billable time</label>
  <label class='checkbox-label'><input type='checkbox' id='hideNoNonBill'/> Hide all without non-billable time</label>
  <button id='resetFilters'>Reset Filters</button>
</div>

<div class='card'>
  <h2>Organization Totals</h2>
  <table>
    <thead><tr><th>Organization</th><th>Billable</th><th>Non-Billable</th><th>Total</th></tr></thead>
    <tbody>$OrgRows</tbody>
  </table>
</div>

<div class='card'>
  <h2>Technician Summary</h2>
  <table>
    <thead><tr><th>Technician</th><th>Billable</th><th>Non-Billable</th><th>Total</th></tr></thead>
    <tbody>$TechRows</tbody>
  </table>
</div>

<div class='card'>
  <h2>Ticket Summary</h2>
  <table id='ticketTable'>
    <thead><tr>
      <th>ID</th><th>Title</th><th>Organization</th><th>Start</th><th>End</th><th>Duration</th><th>Status</th><th>Billable</th><th>Non-Billable</th>
    </tr></thead>
    <tbody>$TicketRows</tbody>
  </table>
</div>

<script>
const techSel=document.getElementById('techFilter'),
      orgSel=document.getElementById('orgFilter'),
      resetBtn=document.getElementById('resetFilters'),
      hideNoBill=document.getElementById('hideNoBill'),
      hideNoNonBill=document.getElementById('hideNoNonBill');

function normalize(s){return (s||'').toLowerCase();}
function hasZeroCell(row,sel){const el=row.querySelector(sel);return el && el.textContent.includes('0h 00m');}

function applyFilters(){
  const tech=normalize(techSel.value), org=normalize(orgSel.value);
  document.querySelectorAll('#ticketTable tbody tr.ticket-row').forEach(r=>{
    const rowOrg=normalize(r.getAttribute('data-org'));
    const rowTechs=(r.getAttribute('data-techs')||'').split(';').map(normalize);
    const showOrg=(org==='__all__')||(rowOrg===org);
    const showTech=(tech==='__all__')||rowTechs.includes(tech);
    const hideBill=hideNoBill.checked && hasZeroCell(r,'.billable');
    const hideNon=hideNoNonBill.checked && hasZeroCell(r,'.nonbill');
    r.style.display=(showOrg && showTech && !hideBill && !hideNon) ? '' : 'none';
  });
}

techSel.onchange=orgSel.onchange=hideNoBill.onchange=hideNoNonBill.onchange=applyFilters;
resetBtn.onclick=()=>{ techSel.value='__ALL__'; orgSel.value='__ALL__'; hideNoBill.checked=true; hideNoNonBill.checked=false; applyFilters(); };
document.addEventListener('DOMContentLoaded',applyFilters);
</script>
</body>
</html>
"@

Set-Content -Path $OutputHtml -Value $html -Encoding UTF8
Start-Process "chrome.exe" $OutputHtml
Write-Host "Report saved to $OutputHtml"
