================================================================================
DAILY CSI SAFETY REPORT - DOCUMENTATION
================================================================================

Created: 24/03/2026
Author: Claude Code

================================================================================
OVERVIEW
================================================================================

This script automatically sends a daily summary of safety incident emails
at 8:00 AM AWST. It reads Safety Flash emails from INX InControl
(noreply@inxsoftware.com) and compiles them into a branded HTML report.

================================================================================
FILES
================================================================================

DailySafetyReport.ps1    - Main PowerShell script
README.txt               - This documentation file

================================================================================
CONFIGURATION
================================================================================

Current Settings:
  - Source Emails:     noreply@inxsoftware.com
  - Recipient:         nellie.grayling@mrl.com.au
  - Schedule:          8:00 AM daily (AWST)
  - Time Range:        Previous 24 hours
  - Source Folder:     Main Inbox
  - Processed Folder:  Inbox\Safety Incidents - Processed
  - Logo:              H:\Other\CSI Logo\CSI logo (black).png

Incident Filtering:
  - Only incidents where "Date Reported" falls within the 24-hour report
    period are included in the report
  - Emails received within 24hrs but with older "Date Reported" are still
    moved to the processed folder but NOT included in the report
  - This ensures the report only shows incidents that occurred within the
    reporting window

To change settings, edit the Configuration section at the top of the script:
  $SenderEmail          - Email address to filter by
  $RecipientEmail       - Where to send the report
  $ReportSubject        - Email subject line
  $ProcessedFolderName  - Folder to move emails after processing
  $LogoPath             - Path to the header logo image

================================================================================
EMAIL PROCESSING
================================================================================

After successfully sending the report, the script automatically:
  1. Creates "Safety Incidents - Processed" subfolder (if it doesn't exist)
  2. Moves all reported emails to that folder

This ensures:
  - Emails are only moved if the report sends successfully
  - Previously reported incidents won't appear in future reports
  - You have an archive of all processed safety emails

To disable this feature:
  Comment out or remove the "Move processed emails" section in the script

================================================================================
REPORT FORMAT
================================================================================

Header:
  - CSI Mining Services logo (black version) on the left
  - Report title on the right
  - Red underline separator

The report includes a table with these columns:
  - Potential       Color-coded badge (Low/Minor/Medium/High/Critical)
  - Ref No          INX reference number
  - Date Reported   When incident was reported
  - Workgroup       Site/team location
  - Event Type      Injury type, Asset Damage, etc.
  - Brief Desc      Summary of the incident

Styling follows MRL Brand Guidelines:
  - Font: Century Gothic
  - Primary colour: MinRes Red (#ce372f)
  - Background: Sand (#f1ede7)
  - Alternating rows: White / 60% Sand

Potential Badge Colours:
  - Low:      Teal (#4b9ba6)
  - Minor:    Gold (#998500)
  - Medium:   Copper (#c37c59)
  - High:     Red (#ce372f)
  - Critical: Black (#000000)

================================================================================
HOW TO RUN MANUALLY
================================================================================

1. Open PowerShell
2. Run:

   powershell -ExecutionPolicy Bypass -File "C:\Users\Nellie.grayling\Daily CSI Safety Report\DailySafetyReport.ps1"

Or double-click the script file (if PowerShell execution is enabled).

================================================================================
SCHEDULED TASK
================================================================================

A Windows Scheduled Task named "DailySafetyIncidentReport" runs this script
daily at 8:00 AM AWST.

Current Task Settings:
  - Task Name:          DailySafetyIncidentReport
  - Schedule:           Daily at 8:00 AM AWST (Perth, Western Australia)
  - Time Zone:          W. Australia Standard Time (UTC+08:00)
  - Status:             Ready (Enabled)
  - Start When Available: Yes (runs at next opportunity if PC was off at 8am)
  - Run On Battery:     Yes

Requirements for scheduled task to run:
  - PC must be powered on (or will run when next available)
  - Outlook must be running or able to start
  - User must be logged in (runs under your user account)

To view/modify the task:
  1. Press Windows + R
  2. Type "taskschd.msc" and press Enter
  3. Find "DailySafetyIncidentReport" in the list

PowerShell commands:

  View status:
    Get-ScheduledTask -TaskName "DailySafetyIncidentReport"

  Run now:
    Start-ScheduledTask -TaskName "DailySafetyIncidentReport"

  Disable:
    Disable-ScheduledTask -TaskName "DailySafetyIncidentReport"

  Enable:
    Enable-ScheduledTask -TaskName "DailySafetyIncidentReport"

To recreate the scheduled task:

$Action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Users\Nellie.grayling\Daily CSI Safety Report\DailySafetyReport.ps1"'
$Trigger = New-ScheduledTaskTrigger -Daily -At '08:00'
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries
Register-ScheduledTask -TaskName 'DailySafetyIncidentReport' -Action $Action -Trigger $Trigger -Settings $Settings -Description 'Sends daily safety incident report at 8am AWST'

================================================================================
REQUIREMENTS
================================================================================

  - Microsoft Outlook installed and configured with your email account
  - Outlook should be running at 8:00 AM (or will run when PC next available)
  - Safety emails must be in main Inbox (not a subfolder)
  - PC must be powered on

================================================================================
TROUBLESHOOTING
================================================================================

PROBLEM: Report not sending
  - Check Outlook is open and logged in
  - Check your Sent Items folder for the report
  - Run script manually to see error messages

PROBLEM: 0 incidents but you know there are emails
  - Verify emails are in main Inbox (not a subfolder)
  - Check sender is exactly "noreply@inxsoftware.com"
  - Emails must be received within last 24 hours

PROBLEM: Scheduled task not running
  - Open Task Scheduler and check task history
  - Verify task status is "Ready"
  - Check PC was on at 8:00 AM

PROBLEM: Only some incidents showing
  - Check emails haven't been moved to a subfolder
  - Verify all emails are from the same sender address

================================================================================
CUSTOMISATION
================================================================================

To change recipient:
  Edit line: $RecipientEmail = "new.email@mrl.com.au"

To change run time:
  Open Task Scheduler > Right-click task > Properties > Triggers > Edit

To change the logo:
  Edit line: $LogoPath = "C:\path\to\new\logo.png"
  Supported formats: PNG, JPG
  Recommended size: 140px wide x 60px tall max

To add more potential levels:
  1. Add to the regex pattern on the line with "if ($Subject -match..."
  2. Add a new case in the $PotentialClass switch statement
  3. Add a new CSS class in $HtmlHead if needed

================================================================================
