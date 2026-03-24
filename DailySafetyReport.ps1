<#
.SYNOPSIS
    Daily Safety Incident Report Generator
.DESCRIPTION
    Collects safety incident emails from the last 24 hours and sends a summary report.
    Styled with MRL brand guidelines.
.NOTES
    Author: Claude Code
    Date: 24/03/2026

    REQUIREMENTS:
    - Microsoft Outlook must be installed and configured
    - Outlook should be running when script executes
    - Safety emails from noreply@inxsoftware.com must be in main Inbox

    SCHEDULED TASK:
    - Task Name: DailySafetyIncidentReport
    - Runs daily at 8:00 AM AWST
#>

# Configuration
$SenderEmail = "noreply@inxsoftware.com"
$RecipientEmail = "nellie.grayling@mrl.com.au"
$ReportSubject = "Daily Safety Incident Report - $(Get-Date -Format 'dd/MM/yyyy')"
$ProcessedFolderName = "Safety Incidents - Processed"
$LogoPath = "H:\Other\CSI Logo\CSI logo (black).png"

# Calculate time range (last 24 hours)
$EndTime = Get-Date
$StartTime = $EndTime.AddHours(-24)

# Load System.Web for HTML encoding
Add-Type -AssemblyName System.Web

try {
    # Connect to Outlook
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    $Inbox = $Namespace.GetDefaultFolder(6)  # 6 = Inbox folder

    # Filter emails from the sender within the time range (searching main Inbox)
    $Filter = "[SenderEmailAddress] = '$SenderEmail' AND [ReceivedTime] >= '$($StartTime.ToString("MM/dd/yyyy HH:mm"))'"
    $FilteredItems = $Inbox.Items.Restrict($Filter)

    # Collect incident data - parse INX InControl format
    $Incidents = @()
    $ProcessedEmails = @()  # Store email objects to move after successful send
    foreach ($Email in $FilteredItems) {
        if ($Email.ReceivedTime -ge $StartTime -and $Email.ReceivedTime -le $EndTime) {
            $Body = $Email.Body
            $Subject = $Email.Subject

            # Extract Potential from subject (e.g., "Low Potential", "Minor Potential")
            $Potential = ""
            if ($Subject -match "(Low|Minor|Moderate|Medium|High|Critical)\s+Potential") {
                $Potential = $Matches[1]
            }

            # Parse fields from email body
            $RefNo = if ($Body -match "Reference No:\s*(\d+)") { $Matches[1] } else { "-" }
            $EventSubType = if ($Body -match "Event Sub Type:\s*([^\r\n]+)") { $Matches[1].Trim() } else { "-" }
            $DateReportedStr = if ($Body -match "Date Reported:\s*([^\r\n]+)") { $Matches[1].Trim() } else { "-" }
            $Workgroup = if ($Body -match "Workgroup:\s*([^\r\n]+)") { $Matches[1].Trim() } else { "-" }
            $BriefDesc = if ($Body -match "Brief Description:\s*([^\r\n]+)") { $Matches[1].Trim() } else { "-" }

            # Parse DateReported and check if within report period
            $IncludeInReport = $false
            try {
                # Parse date format like "23-Mar-26" (dd-MMM-yy)
                $DateReportedParsed = [DateTime]::ParseExact($DateReportedStr, "dd-MMM-yy", [System.Globalization.CultureInfo]::InvariantCulture)
                # Check if DateReported falls within the 24-hour report period (date only comparison)
                if ($DateReportedParsed.Date -ge $StartTime.Date -and $DateReportedParsed.Date -le $EndTime.Date) {
                    $IncludeInReport = $true
                }
            } catch {
                # If date parsing fails, exclude from report
                $IncludeInReport = $false
            }

            # Only add to incidents if DateReported is within report period
            if ($IncludeInReport) {
                $Incidents += [PSCustomObject]@{
                    ReceivedTime  = $Email.ReceivedTime.ToString("dd/MM/yyyy HH:mm")
                    Potential     = $Potential
                    RefNo         = $RefNo
                    DateReported  = $DateReportedStr
                    Workgroup     = $Workgroup
                    EventSubType  = $EventSubType
                    BriefDesc     = $BriefDesc
                }
            }

            # Store email object for moving later (move all processed emails, even if not in report)
            $ProcessedEmails += $Email
        }
    }

    # Sort by received time (newest first)
    if ($Incidents.Count -gt 0) {
        $Incidents = @($Incidents | Sort-Object { [DateTime]::ParseExact($_.ReceivedTime, "dd/MM/yyyy HH:mm", $null) } -Descending)
    }

    # Load and encode logo as base64
    $LogoBase64 = ""
    if (Test-Path $LogoPath) {
        $LogoBytes = [System.IO.File]::ReadAllBytes($LogoPath)
        $LogoBase64 = [System.Convert]::ToBase64String($LogoBytes)
    }

    # Build HTML report - MRL Brand Guidelines
    $HtmlHead = @"
<!DOCTYPE html>
<html>
<head>
<style>
    body { font-family: 'Century Gothic', sans-serif; font-size: 10pt; color: #000000; margin: 0; padding: 20px; }
    .header { display: table; width: 100%; margin-bottom: 20px; border-bottom: 3px solid #ce372f; padding-bottom: 15px; }
    .header-logo { display: table-cell; vertical-align: middle; width: 150px; }
    .header-logo img { max-height: 60px; max-width: 140px; }
    .header-title { display: table-cell; vertical-align: middle; }
    h1 { color: #ce372f; font-size: 18pt; font-weight: bold; text-transform: uppercase; margin: 0; }
    table.data { border-collapse: collapse; width: 100%; margin-top: 15px; }
    th { background-color: #ce372f; color: #ffffff; padding: 8px 10px; text-align: left; font-weight: bold; text-transform: uppercase; font-size: 9pt; }
    td { border: 1px solid #e0c09d; padding: 8px 10px; vertical-align: top; font-size: 9pt; }
    tr:nth-child(odd) td { background-color: #ffffff; }
    tr:nth-child(even) td { background-color: #f7f5f2; }
    .summary { background-color: #f1ede7; padding: 15px; margin-bottom: 15px; border-left: 4px solid #ce372f; }
    .summary strong { color: #ce372f; text-transform: uppercase; }
    .no-incidents { color: #000000; font-weight: bold; }
    .footer { font-size: 9pt; color: #544741; margin-top: 20px; padding-top: 10px; border-top: 1px solid #e0c09d; }
    .potential-low { background-color: #4b9ba6; color: #ffffff; padding: 3px 8px; font-weight: bold; font-size: 8pt; }
    .potential-minor { background-color: #998500; color: #ffffff; padding: 3px 8px; font-weight: bold; font-size: 8pt; }
    .potential-moderate { background-color: #c37c59; color: #ffffff; padding: 3px 8px; font-weight: bold; font-size: 8pt; }
    .potential-high { background-color: #ce372f; color: #ffffff; padding: 3px 8px; font-weight: bold; font-size: 8pt; }
    .potential-critical { background-color: #000000; color: #ffffff; padding: 3px 8px; font-weight: bold; font-size: 8pt; }
</style>
</head>
<body>
"@

    # Build header with logo
    $LogoHtml = ""
    if ($LogoBase64) {
        $LogoHtml = "<img src=`"data:image/png;base64,$LogoBase64`" alt=`"CSI Logo`">"
    }

    $HtmlBody = @"
<div class="header">
    <div class="header-logo">$LogoHtml</div>
    <div class="header-title"><h1>Daily Safety Incident Report</h1></div>
</div>
<div class="summary">
    <strong>Report Period:</strong> $($StartTime.ToString("dd/MM/yyyy HH:mm")) - $($EndTime.ToString("dd/MM/yyyy HH:mm")) AWST<br>
    <strong>Total Incidents:</strong> $($Incidents.Count)
</div>
"@

    if ($Incidents.Count -eq 0) {
        $HtmlBody += '<p class="no-incidents">No safety incidents reported in the last 24 hours.</p>'
    } else {
        $HtmlBody += @"
<table class="data">
    <tr>
        <th>Potential</th>
        <th>Ref No</th>
        <th>Date Reported</th>
        <th>Workgroup</th>
        <th>Event Type</th>
        <th>Brief Description</th>
    </tr>
"@
        foreach ($Incident in $Incidents) {
            # Set potential badge class
            $PotentialClass = switch ($Incident.Potential) {
                "Low" { "potential-low" }
                "Minor" { "potential-minor" }
                "Moderate" { "potential-moderate" }
                "Medium" { "potential-moderate" }
                "High" { "potential-high" }
                "Critical" { "potential-critical" }
                default { "potential-low" }
            }
            $PotentialText = if ($Incident.Potential) { $Incident.Potential.ToUpper() } else { "-" }

            $SafeWorkgroup = [System.Web.HttpUtility]::HtmlEncode($Incident.Workgroup)
            $SafeEventType = [System.Web.HttpUtility]::HtmlEncode($Incident.EventSubType)
            $SafeDesc = [System.Web.HttpUtility]::HtmlEncode($Incident.BriefDesc)

            $HtmlBody += @"
    <tr>
        <td><span class="$PotentialClass">$PotentialText</span></td>
        <td>$($Incident.RefNo)</td>
        <td>$($Incident.DateReported)</td>
        <td>$SafeWorkgroup</td>
        <td>$SafeEventType</td>
        <td>$SafeDesc</td>
    </tr>
"@
        }
        $HtmlBody += "</table>"
    }

    $HtmlFooter = @"
<div class="footer">
    This is an automated report generated at $(Get-Date -Format "dd/MM/yyyy HH:mm") AWST.<br>
    Source: Emails from $SenderEmail
</div>
</body>
</html>
"@

    $FullHtml = $HtmlHead + $HtmlBody + $HtmlFooter

    # Create and send the email
    $Mail = $Outlook.CreateItem(0)  # 0 = Mail item
    $Mail.To = $RecipientEmail
    $Mail.Subject = $ReportSubject
    $Mail.HTMLBody = $FullHtml
    $Mail.Send()

    Write-Host "Report sent successfully to $RecipientEmail"
    Write-Host "Incidents found: $($Incidents.Count)"

    # Move processed emails to subfolder (only on successful send)
    if ($ProcessedEmails.Count -gt 0) {
        try {
            # Get or create the processed folder
            $ProcessedFolder = $null
            try {
                $ProcessedFolder = $Inbox.Folders.Item($ProcessedFolderName)
            } catch {
                # Folder doesn't exist, create it
                $ProcessedFolder = $Inbox.Folders.Add($ProcessedFolderName)
                Write-Host "Created folder: $ProcessedFolderName"
            }

            # Move each email to the processed folder
            $movedCount = 0
            foreach ($Email in $ProcessedEmails) {
                $Email.Move($ProcessedFolder) | Out-Null
                $movedCount++
            }
            Write-Host "Moved $movedCount emails to '$ProcessedFolderName'"
        } catch {
            Write-Warning "Failed to move emails: $_"
        }
    }

} catch {
    Write-Error "Failed to generate report: $_"
    exit 1
} finally {
    # Clean up COM objects
    if ($Mail) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Mail) | Out-Null }
    if ($ProcessedFolder) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ProcessedFolder) | Out-Null }
    if ($FilteredItems) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($FilteredItems) | Out-Null }
    if ($Inbox) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Inbox) | Out-Null }
    if ($Namespace) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null }
    if ($Outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
