# Add Windows Forms .NET assembly
Add-Type -AssemblyName System.Windows.Forms

# Function to generate HTML content
function Generate-HTMLReport {
    param (
        [string]$htmlPath,
        [string]$content
    )

    # HTML structure with white background and black text
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
            background-color: #ffffff;
            color: #000000;
            margin: 20px;
        }
        h1 {
            color: #0056b3;
            text-align: center;
        }
        .section-title {
            color: #333;
            font-size: 18px;
            margin-top: 20px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
        }
        th, td {
            padding: 10px;
            border: 1px solid #ccc;
            text-align: left;
        }
        th {
            background-color: #0056b3;
            color: #fff;
        }
        pre {
            background-color: #e9ecef;
            padding: 10px;
            border-radius: 5px;
            overflow-x: auto;
            white-space: pre-wrap;
            word-wrap: break-word;
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

# Function to check for disk errors using CHKDSK
function Check-DiskErrors {
    $diskReport = "<h2 class='section-title'>Disk Errors</h2>"
    $drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
    foreach ($drive in $drives) {
        $driveLetter = $drive.DeviceID
        $diskReport += "<p>Running CHKDSK on drive $driveLetter...</p>"
        $chkdskResult = cmd /c "chkdsk $driveLetter /scan"
        $diskReport += "<pre>$chkdskResult</pre>"
    }
    return $diskReport
}

# Function to check for system file corruption using SFC
function Run-SFC {
    $sfcReport = "<h2 class='section-title'>System File Checker (SFC)</h2>"
    $result = cmd /c "sfc /scannow"
    $sfcReport += "<pre>$result</pre>"

    if ($result -like "*Windows Resource Protection found corrupt files*") {
        $sfcReport += "<p>SFC found and repaired corrupt system files.</p>"
    } elseif ($result -like "*Windows Resource Protection did not find any integrity violations*") {
        $sfcReport += "<p>No integrity violations found by SFC.</p>"
    }
    return $sfcReport
}

# Function to check the health of the system image using DISM
function Run-DISMCheck {
    $dismReport = "<h2 class='section-title'>DISM Health Check</h2>"
    $result = cmd /c "DISM /Online /Cleanup-Image /CheckHealth"
    $dismReport += "<pre>$result</pre>"

    if ($result -like "*The component store is repairable*") {
        $dismReport += "<p>Component store corruption detected. Attempting repair...</p>"
        $repairResult = cmd /c "DISM /Online /Cleanup-Image /RestoreHealth"
        $dismReport += "<pre>$repairResult</pre>"
    } elseif ($result -like "*No component store corruption detected*") {
        $dismReport += "<p>No component store corruption detected.</p>"
    }
    return $dismReport
}

# Function to check event logs for critical system errors
function Check-EventLogErrors {
    $eventLogReport = "<h2 class='section-title'>Event Log Errors</h2>"
    $criticalEvents = Get-WinEvent -LogName System -FilterHashtable @{Level=1; StartTime=(Get-Date).AddDays(-7)} -MaxEvents 20

    if ($criticalEvents.Count -eq 0) {
        $eventLogReport += "<p>No critical system errors found in the event logs.</p>"
    } else {
        $eventLogReport += "<table><tr><th>Error ID</th><th>Time</th><th>Message</th></tr>"
        foreach ($event in $criticalEvents) {
            $eventLogReport += "<tr><td>$($event.Id)</td><td>$($event.TimeCreated)</td><td>$($event.Message)</td></tr>"
        }
        $eventLogReport += "</table>"
    }
    return $eventLogReport
}

# Function to generate the full system health report
function Generate-FullReport {
    $fullReport = "<h2 class='section-title'>Operating System Health Check</h2>"
    $fullReport += Check-DiskErrors
    $fullReport += Run-SFC
    $fullReport += Check-EventLogErrors
    $fullReport += Run-DISMCheck
    return $fullReport
}

# Generate the HTML report
$reportContent = Generate-FullReport
$htmlPath = "$env:USERPROFILE\Desktop\OperatingSystemErrorCheckReport.html"
Generate-HTMLReport -htmlPath $htmlPath -content $reportContent

# Open the generated HTML report in the default web browser
Start-Process $htmlPath

Write-Host "Operating System Error Check Report has been generated and saved to your desktop."