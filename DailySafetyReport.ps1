<#
.SYNOPSIS
    Daily Safety Incident Report Generator
.DESCRIPTION
    Collects safety incident emails from the last 24 hours and sends a summary report.
    Styled with MRL brand guidelines. Mobile-friendly responsive design.
.NOTES
    Author: Claude Code
    Date: 24/03/2026

    REQUIREMENTS:
    - Microsoft Outlook must be installed and configured
    - Outlook should be running when script executes
    - Safety emails should be in "Safety Incidents" folder (via Outlook rule)
    - Falls back to main Inbox if folder is empty

    SCHEDULED TASK:
    - Task Name: DailySafetyIncidentReport
    - Runs daily at 8:00 AM AWST
#>

# Configuration
$SenderEmail = "noreply@inxsoftware.com"
$RecipientEmail = "nellie.grayling@mrl.com.au"
$ReportSubject = "Daily Safety Incident Report - $(Get-Date -Format 'dd/MM/yyyy')"
$SourceFolderName = "Safety Incidents"
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

    # Try to get Safety Incidents folder, fall back to Inbox
    $SourceFolder = $null
    try {
        $SourceFolder = $Inbox.Folders.Item($SourceFolderName)
        if ($SourceFolder.Items.Count -eq 0) {
            # Folder exists but empty, use Inbox instead
            $SourceFolder = $Inbox
            Write-Host "Safety Incidents folder empty, using Inbox"
        } else {
            Write-Host "Reading from Safety Incidents folder"
        }
    } catch {
        # Folder doesn't exist, use Inbox
        $SourceFolder = $Inbox
        Write-Host "Safety Incidents folder not found, using Inbox"
    }

    # Filter emails from the sender within the time range
    $Filter = "[SenderEmailAddress] = '$SenderEmail' AND [ReceivedTime] >= '$($StartTime.ToString("MM/dd/yyyy HH:mm"))'"
    $FilteredItems = $SourceFolder.Items.Restrict($Filter)

    # Collect incident data - parse INX InControl format
    $Incidents = @()
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

    # Build HTML report - MRL Brand Guidelines with Mobile-Responsive Design
    $HtmlHead = @"
<!DOCTYPE html>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
    /* Base styles */
    body {
        font-family: 'Century Gothic', Arial, sans-serif;
        font-size: 14px;
        color: #000000;
        margin: 0;
        padding: 15px;
        -webkit-text-size-adjust: 100%;
    }

    /* Header - stacks on mobile */
    .header {
        width: 100%;
        margin-bottom: 20px;
        border-bottom: 3px solid #ce372f;
        padding-bottom: 15px;
    }
    .header-logo {
        display: inline-block;
        vertical-align: middle;
        margin-right: 15px;
        margin-bottom: 10px;
    }
    .header-logo img {
        max-height: 50px;
        max-width: 120px;
        width: auto;
    }
    .header-title {
        display: inline-block;
        vertical-align: middle;
    }
    h1 {
        color: #ce372f;
        font-size: 18px;
        font-weight: bold;
        text-transform: uppercase;
        margin: 0;
    }

    /* Summary box */
    .summary {
        background-color: #f1ede7;
        padding: 15px;
        margin-bottom: 15px;
        border-left: 4px solid #ce372f;
        font-size: 14px;
    }
    .summary strong {
        color: #ce372f;
        text-transform: uppercase;
    }
    .no-incidents {
        color: #000000;
        font-weight: bold;
        padding: 20px;
        text-align: center;
    }

    /* Mobile card layout for incidents */
    .incident-card {
        background-color: #ffffff;
        border: 1px solid #e0c09d;
        border-left: 4px solid #ce372f;
        margin-bottom: 15px;
        padding: 15px;
        border-radius: 4px;
    }
    .incident-card:nth-child(even) {
        background-color: #f7f5f2;
    }
    .incident-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 10px;
        flex-wrap: wrap;
        gap: 8px;
    }
    .incident-ref {
        font-weight: bold;
        font-size: 16px;
        color: #000000;
    }
    .incident-row {
        margin-bottom: 8px;
        font-size: 13px;
    }
    .incident-label {
        font-weight: bold;
        color: #544741;
        display: inline-block;
        min-width: 100px;
    }
    .incident-value {
        color: #000000;
    }
    .incident-desc {
        margin-top: 10px;
        padding-top: 10px;
        border-top: 1px solid #e0c09d;
        font-size: 13px;
    }

    /* Potential badges */
    .potential-badge {
        display: inline-block;
        padding: 5px 12px;
        font-weight: bold;
        font-size: 12px;
        border-radius: 3px;
        text-transform: uppercase;
    }
    .potential-low { background-color: #4b9ba6; color: #ffffff; }
    .potential-minor { background-color: #998500; color: #ffffff; }
    .potential-moderate { background-color: #c37c59; color: #ffffff; }
    .potential-high { background-color: #ce372f; color: #ffffff; }
    .potential-critical { background-color: #000000; color: #ffffff; }

    /* Footer */
    .footer {
        font-size: 12px;
        color: #544741;
        margin-top: 20px;
        padding-top: 10px;
        border-top: 1px solid #e0c09d;
    }

    /* Desktop table view - hidden on mobile by default */
    @media screen and (min-width: 768px) {
        h1 { font-size: 22px; }
        .header-logo img { max-height: 60px; max-width: 140px; }
    }
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
        # Mobile-friendly card layout
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
<div class="incident-card">
    <div class="incident-header">
        <span class="incident-ref">Ref: $($Incident.RefNo)</span>
        <span class="potential-badge $PotentialClass">$PotentialText</span>
    </div>
    <div class="incident-row">
        <span class="incident-label">Date Reported:</span>
        <span class="incident-value">$($Incident.DateReported)</span>
    </div>
    <div class="incident-row">
        <span class="incident-label">Workgroup:</span>
        <span class="incident-value">$SafeWorkgroup</span>
    </div>
    <div class="incident-row">
        <span class="incident-label">Event Type:</span>
        <span class="incident-value">$SafeEventType</span>
    </div>
    <div class="incident-desc">
        <span class="incident-label">Description:</span><br>
        <span class="incident-value">$SafeDesc</span>
    </div>
</div>
"@
        }
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

} catch {
    Write-Error "Failed to generate report: $_"
    exit 1
} finally {
    # Clean up COM objects
    if ($Mail) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Mail) | Out-Null }
    if ($SourceFolder -and $SourceFolder -ne $Inbox) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($SourceFolder) | Out-Null }
    if ($FilteredItems) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($FilteredItems) | Out-Null }
    if ($Inbox) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Inbox) | Out-Null }
    if ($Namespace) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null }
    if ($Outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
