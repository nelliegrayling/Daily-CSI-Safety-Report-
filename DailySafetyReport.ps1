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
$ReportsFolder = "C:\Users\Nellie.grayling\Daily CSI Safety Report\Reports"
$ChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"

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

    # Build HTML report - MRL Brand Guidelines with Landscape Layout
    $HtmlHead = @"
<!DOCTYPE html>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
    @page { size: landscape; margin: 15mm; }
    body { font-family: 'Century Gothic', sans-serif; font-size: 10pt; color: #000000; margin: 0; padding: 20px; -webkit-text-size-adjust: 100%; }
    .header { width: 100%; margin-bottom: 20px; }
    .header-logo { margin-bottom: 10px; }
    .header-logo img { max-height: 90px; max-width: 210px; }
    .header-line { border-bottom: 3px solid #ce372f; margin-bottom: 10px; }
    h1 { color: #000000; font-size: 18pt; font-weight: bold; margin: 0 0 15px 0; }
    table.data { border-collapse: collapse; width: 100%; margin-top: 15px; }
    thead { border-bottom: 3px solid #ce372f; }
    th { background-color: #000000; color: #ffffff; padding: 8px 10px; text-align: left; font-weight: bold; text-transform: uppercase; font-size: 9pt; }
    td { border: 1px solid #e0e0e0; padding: 8px 10px; vertical-align: top; font-size: 9pt; }
    tbody tr:nth-child(odd) td { background-color: #ffffff; }
    tbody tr:nth-child(even) td { background-color: #f5f5f5; }
    .summary { background-color: #f1ede7; padding: 15px; margin-bottom: 15px; border-left: 4px solid #ce372f; }
    .summary strong { color: #ce372f; text-transform: uppercase; }
    .no-incidents { color: #000000; font-weight: bold; }
    .footer { font-size: 9pt; color: #544741; margin-top: 20px; padding-top: 10px; border-top: 1px solid #e0c09d; }
    .potential-badge { display: inline-block; padding: 5px 14px; font-weight: bold; font-size: 9pt; border-radius: 3px; white-space: nowrap; text-align: center; min-width: 60px; }
    .potential-low { background-color: #4b9ba6; color: #ffffff; }
    .potential-minor { background-color: #998500; color: #ffffff; }
    .potential-moderate { background-color: #c37c59; color: #ffffff; }
    .potential-high { background-color: #ce372f; color: #ffffff; }
    .potential-critical { background-color: #000000; color: #ffffff; }

    /* Mobile responsive styles */
    @media screen and (max-width: 768px) {
        body { padding: 10px; }
        h1 { font-size: 14pt; }
        .header { text-align: center; }
        table.data, table.data thead, table.data tbody, table.data th, table.data td, table.data tr { display: block; }
        table.data thead tr { position: absolute; top: -9999px; left: -9999px; }
        table.data tr { border: 1px solid #e0c09d; border-left: 4px solid #ce372f; margin-bottom: 15px; background-color: #ffffff; }
        table.data tr:nth-child(even) { background-color: #f7f5f2; }
        table.data td { border: none; border-bottom: 1px solid #e0c09d; padding: 10px; padding-left: 40%; position: relative; text-align: left; }
        table.data td:last-child { border-bottom: none; }
        table.data td:before { content: attr(data-label); position: absolute; left: 10px; width: 35%; padding-right: 10px; font-weight: bold; color: #544741; text-transform: uppercase; font-size: 8pt; }
    }

    /* Print styles for landscape */
    @media print {
        @page { size: landscape; margin: 10mm; }
        body { padding: 0; }
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

    # Format date for header (DD Month YYYY)
    $ReportDateFormatted = $EndTime.ToString("dd MMMM yyyy")

    $HtmlBody = @"
<div class="header">
    <div class="header-logo">$LogoHtml</div>
    <div class="header-line"></div>
    <h1>Daily Safety Incident Report $ReportDateFormatted</h1>
</div>
<div class="summary">
    <strong>Report Period:</strong> $($StartTime.ToString("dd/MM/yyyy HH:mm")) - $($EndTime.ToString("dd/MM/yyyy HH:mm")) AWST<br>
    <strong>Total Incidents:</strong> $($Incidents.Count)
</div>
"@

    if ($Incidents.Count -eq 0) {
        $HtmlBody += '<p class="no-incidents">No safety incidents reported in the last 24 hours.</p>'
    } else {
        # Table layout
        $HtmlBody += @"
<table class="data">
    <thead>
        <tr>
            <th>Potential</th>
            <th>Ref No</th>
            <th>Date Reported</th>
            <th>Workgroup</th>
            <th>Event Type</th>
            <th>Brief Description</th>
        </tr>
    </thead>
    <tbody>
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
        <td data-label="Potential"><span class="potential-badge $PotentialClass">$PotentialText</span></td>
        <td data-label="Ref No">$($Incident.RefNo)</td>
        <td data-label="Date Reported">$($Incident.DateReported)</td>
        <td data-label="Workgroup">$SafeWorkgroup</td>
        <td data-label="Event Type">$SafeEventType</td>
        <td data-label="Description">$SafeDesc</td>
    </tr>
"@
        }
        $HtmlBody += "</tbody></table>"
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

    # Generate PDF filename with timestamp
    $PdfFileName = "Daily Safety Incident Report - $(Get-Date -Format 'yyyy-MM-dd_HHmmss').pdf"
    $PdfPath = Join-Path $ReportsFolder $PdfFileName

    # Create temp HTML file for conversion
    $TempHtmlPath = Join-Path $env:TEMP "SafetyReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $FullHtml | Out-File -FilePath $TempHtmlPath -Encoding UTF8

    # Convert HTML to PDF using Chrome headless (landscape A4)
    & $ChromePath --headless --disable-gpu --no-pdf-header-footer --print-to-pdf="$PdfPath" --print-to-pdf-no-header $TempHtmlPath 2>$null
    Start-Sleep -Seconds 3  # Wait for PDF to be written

    if (-not (Test-Path $PdfPath)) {
        throw "Failed to generate PDF file"
    }

    Write-Host "PDF saved to: $PdfPath"

    # Create and send the email with PDF attachment
    $Mail = $Outlook.CreateItem(0)  # 0 = Mail item
    $Mail.To = $RecipientEmail
    $Mail.Subject = $ReportSubject
    $Mail.Body = "Please find attached the Daily Safety Incident Report.`n`nReport Period: $($StartTime.ToString('dd/MM/yyyy HH:mm')) - $($EndTime.ToString('dd/MM/yyyy HH:mm')) AWST`nTotal Incidents: $($Incidents.Count)`n`nThis is an automated report."
    $Mail.Attachments.Add($PdfPath)
    $Mail.Send()

    # Clean up temp HTML file
    Remove-Item -Path $TempHtmlPath -Force -ErrorAction SilentlyContinue

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
