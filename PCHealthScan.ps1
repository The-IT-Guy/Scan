# Add Windows Forms .NET assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to generate HTML content with "The IT Guy" Documentation Styling
function Generate-HTMLReport {
    param (
        [string]$htmlPath,
        [string]$content
    )

    # HTML structure matching your core branding (White background, Black text, Blue accents)
    $htmlContent = @"
<!DOCTYPE html>
<html lang='en'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>PC Health Report</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #ffffff;
            margin: 20px;
            color: #000000;
        }
        h1 {
            color: #0056b3;
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
        .info-p {
            margin: 5px 0;
            font-size: 14px;
        }
    </style>
</head>
<body>
    <h1>PC Health Report</h1>
    <div>
        $content
    </div>
</body>
</html>
"@

    # Write the HTML content to the specified file path
    $htmlContent | Out-File -Encoding UTF8 $htmlPath
}

# Function to run all the system checks
function Run-SystemScan {
    $reportContent = ""

    # -------- System Summary --------
    $reportContent += "<div class='section-title'>System Summary</div>"
    $os = Get-WmiObject -Class Win32_OperatingSystem
    $reportContent += "<p class='info-p'><strong>Operating System:</strong> $($os.Caption) ($($os.Version))</p>"
    $reportContent += "<p class='info-p'><strong>Total Physical Memory:</strong> $([math]::Round($os.TotalVisibleMemorySize / 1MB, 2)) GB</p>"
    $reportContent += "<p class='info-p'><strong>Free Physical Memory:</strong> $([math]::Round($os.FreePhysicalMemory / 1MB, 2)) GB</p>"
    $uptime = (Get-Date) - [System.Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)
    $reportContent += "<p class='info-p'><strong>System Uptime:</strong> $([math]::Round($uptime.TotalHours, 2)) hours</p>"

    # -------- Physical Disk Health --------
    $reportContent += "<div class='section-title'>Physical Disk Health</div>"
    $disks = Get-PhysicalDisk | Select-Object DeviceID, MediaType, OperationalStatus, HealthStatus
    $reportContent += "<table><tr><th>Device ID</th><th>Media Type</th><th>Status</th><th>Health</th></tr>"
    foreach ($disk in $disks) {
        $reportContent += "<tr><td>$($disk.DeviceID)</td><td>$($disk.MediaType)</td><td>$($disk.OperationalStatus)</td><td>$($disk.HealthStatus)</td></tr>"
    }
    $reportContent += "</table>"

    # -------- Memory Status --------
    $reportContent += "<div class='section-title'>Hardware Memory Status</div>"
    $memoryInfo = Get-WmiObject -Class Win32_PhysicalMemory
    $reportContent += "<table><tr><th>Device</th><th>Capacity (GB)</th><th>Speed (MHz)</th><th>Status</th></tr>"
    foreach ($mem in $memoryInfo) {
        $capacityGB = [math]::round($mem.Capacity / 1GB, 2)
        $reportContent += "<tr><td>$($mem.DeviceLocator)</td><td>$capacityGB GB</td><td>$($mem.Speed) MHz</td><td>$($mem.Status)</td></tr>"
    }
    $reportContent += "</table>"

    # -------- Disk Errors (CHKDSK) --------
    $reportContent += "<div class='section-title'>Disk Integrity (CHKDSK Scan)</div>"
    $drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
    foreach ($drive in $drives) {
        $driveLetter = $drive.DeviceID
        $reportContent += "<p class='info-p'>Scanning Drive $driveLetter...</p>"
        $chkdskResult = cmd /c "chkdsk $driveLetter /scan" | Out-String
        $reportContent += "<pre>$chkdskResult</pre>"
    }

    # -------- Windows Event Log Errors --------
    $reportContent += "<div class='section-title'>Recent System Errors (Event Log)</div>"
    try {
        $errorEvents = Get-WinEvent -LogName System -FilterHashtable @{Level=2} -MaxEvents 10 -ErrorAction SilentlyContinue
        if ($null -eq $errorEvents) {
            $reportContent += "<p>No recent system errors found.</p>"
        } else {
            $reportContent += "<table><tr><th>Time</th><th>ID</th><th>Message</th></tr>"
            foreach ($event in $errorEvents) {
                $reportContent += "<tr><td>$($event.TimeCreated)</td><td>$($event.Id)</td><td>$($event.Message)</td></tr>"
            }
            $reportContent += "</table>"
        }
    } catch {
        $reportContent += "<p>Unable to retrieve system event logs.</p>"
    }

    # -------- Resource Usage --------
    $reportContent += "<div class='section-title'>High Resource Usage Processes</div>"
    $processes = Get-Process | Where-Object { $_.CPU -gt 50 -or $_.WorkingSet -gt 500MB } | Select-Object Name, CPU, WorkingSet
    if ($processes) {
        $reportContent += "<table><tr><th>Process</th><th>CPU</th><th>Memory (MB)</th></tr>"
        foreach ($process in $processes) {
            $memMB = [math]::round($process.WorkingSet / 1MB, 2)
            $reportContent += "<tr><td>$($process.Name)</td><td>$($process.CPU)</td><td>$memMB MB</td></tr>"
        }
        $reportContent += "</table>"
    } else {
        $reportContent += "<p>No high-usage processes detected.</p>"
    }

    return $reportContent
}

# 1. Execute Scan
Write-Host "Starting PC Health Scan..." -ForegroundColor Cyan
$fullReport = Run-SystemScan

# 2. Save Report
$htmlPath = "$env:USERPROFILE\Desktop\PCHealthReport.html"
Generate-HTMLReport -htmlPath $htmlPath -content $fullReport

# 3. Prompt User to View Report
$title = "Scan Complete"
$message = "PC Health Report generated on Desktop. Would you like to view it now?"
$buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
$icon = [System.Windows.Forms.MessageBoxIcon]::Information

$result = [System.Windows.Forms.MessageBox]::Show($message, $title, $buttons, $icon)

if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
    Start-Process $htmlPath
} else {
    Write-Host "Report saved to Desktop as PCHealthReport.html"
}
