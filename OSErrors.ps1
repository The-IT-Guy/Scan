# Add Windows Forms .NET assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to generate HTML content with "The IT Guy" Documentation Styling
function Generate-HTMLReport {
    param (
        [string]$htmlPath,
        [string]$content
    )

    # HTML structure matching your corporate documentation theme
    $htmlContent = @"
<!DOCTYPE html>
<html lang='en'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>Operating System Health Report</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #ffffff; /* Consistent White Background */
            margin: 20px;
            color: #000000; /* Consistent Black Text */
        }
        h1 {
            color: #0056b3; /* The IT Guy Blue */
            text-align: center;
            border-bottom: 2px solid #0056b3;
            padding-bottom: 10px;
        }
        .section-title {
            color: #333333;
            font-size: 18px;
            margin-top: 25px;
            font-weight: bold;
            border-left: 5px solid #0056b3;
            padding-left: 10px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
        }
        th, td {
            padding: 10px;
            border: 1px solid #cccccc;
            text-align: left;
        }
        th {
            background-color: #0056b3;
            color: #ffffff;
        }
        pre {
            background-color: #e9ecef;
            padding: 15px;
            border-radius: 5px;
            overflow-x: auto;
            white-space: pre-wrap;
            word-wrap: break-word;
            font-family: 'Consolas', monospace;
            border: 1px solid #ccc;
        }
        p {
            margin: 10px 0;
            line-height: 1.5;
        }
    </style>
</head>
<body>
    <h1>Operating System Health Report</h1>
    <div>
        $content
    </div>
</body>
</html>
"@

    # Write the HTML content to the specified file path
    $htmlContent | Out-File -Encoding UTF8 $htmlPath
}

# --- Diagnostic Functions ---

function Check-DiskErrors {
    $diskReport = "<div class='section-title'>Disk Errors (CHKDSK)</div>"
    $drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
    foreach ($drive in $drives) {
        $driveLetter = $drive.DeviceID
        $diskReport += "<p>Running CHKDSK on drive $driveLetter...</p>"
        $chkdskResult = cmd /c "chkdsk $driveLetter /scan" | Out-String
        $diskReport += "<pre>$chkdskResult</pre>"
    }
    return $diskReport
}

function Run-SFC {
    $sfcReport = "<div class='section-title'>System File Checker (SFC)</div>"
    $result = cmd /c "sfc /scannow" | Out-String
    $sfcReport += "<pre>$result</pre>"
    
    if ($result -like "*found corrupt files*") {
        $sfcReport += "<p style='color: #d9534f;'><strong>Note:</strong> SFC found and repaired corrupt system files.</p>"
    }
    return $sfcReport
}

function Check-EventLogErrors {
    $eventLogReport = "<div class='section-title'>Critical Event Logs (Last 7 Days)</div>"
    try {
        $criticalEvents = Get-WinEvent -LogName System -FilterHashtable @{Level=1; StartTime=(Get-Date).AddDays(-7)} -MaxEvents 20 -ErrorAction SilentlyContinue
        
        if ($null -eq $criticalEvents) {
            $eventLogReport += "<p>No critical system errors found in the event logs.</p>"
        } else {
            $eventLogReport += "<table><tr><th>Error ID</th><th>Time</th><th>Message</th></tr>"
            foreach ($event in $criticalEvents) {
                $eventLogReport += "<tr><td>$($event.Id)</td><td>$($event.TimeCreated)</td><td>$($event.Message)</td></tr>"
            }
            $eventLogReport += "</table>"
        }
    } catch {
        $eventLogReport += "<p>Could not retrieve event logs.</p>"
    }
    return $eventLogReport
}

function Run-DISMCheck {
    $dismReport = "<div class='section-title'>DISM Image Health Check</div>"
    $result = cmd /c "DISM /Online /Cleanup-Image /CheckHealth" | Out-String
    $dismReport += "<pre>$result</pre>"
    return $dismReport
}

# --- Main Execution ---

Write-Host "Running OS Health Check... Please wait." -ForegroundColor Cyan

$fullContent = Check-DiskErrors
$fullContent += Run-SFC
$fullContent += Check-EventLogErrors
$fullContent += Run-DISMCheck

# Save the Report
$htmlPath = "$env:USERPROFILE\Desktop\OSErrorCheckReport.html"
Generate-HTMLReport -htmlPath $htmlPath -content $fullContent

# Prompt User to View Report
$title = "OS Scan Complete"
$message = "The Operating System Health Check is finished. Would you like to view the report now?"
$buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
$icon = [System.Windows.Forms.MessageBoxIcon]::Information

$result = [System.Windows.Forms.MessageBox]::Show($message, $title, $buttons, $icon)

if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
    Start-Process $htmlPath
} else {
    Write-Host "Report saved to Desktop as OSErrorCheckReport.html"
}
