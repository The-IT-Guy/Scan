# Add Windows Forms .NET assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to generate HTML content
function Generate-HTMLReport {
    param (
        [string]$htmlPath,
        [string]$content
    )

    # HTML structure matching theme from other reports
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
            background-color: #f4f4f4;
            margin: 20px;
            color: #333;
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
    $reportContent += "<h2 class='section-title'>System Summary</h2>"
    $os = Get-WmiObject -Class Win32_OperatingSystem
    $reportContent += "<p>Operating System: $($os.Caption) $($os.Version)</p>"
    $reportContent += "<p>Total Physical Memory: $([math]::Round($os.TotalVisibleMemorySize / 1MB, 2)) GB</p>"
    $reportContent += "<p>Free Physical Memory: $([math]::Round($os.FreePhysicalMemory / 1MB, 2)) GB</p>"
    $uptime = (Get-Date) - [System.Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)
    $reportContent += "<p>System Uptime: $([math]::Round($uptime.TotalHours, 2)) hours</p>"

    # -------- Hardware Health Check --------
    $reportContent += "<h2 class='section-title'>Physical Disk Health</h2>"
    $disks = Get-PhysicalDisk | Select-Object DeviceID, MediaType, OperationalStatus, HealthStatus
    $reportContent += "<table><tr><th>Device ID</th><th>Media Type</th><th>Status</th><th>Health</th></tr>"
    foreach ($disk in $disks) {
        $reportContent += "<tr><td>$($disk.DeviceID)</td><td>$($disk.MediaType)</td><td>$($disk.OperationalStatus)</td><td>$($disk.HealthStatus)</td></tr>"
    }
    $reportContent += "</table>"

    # -------- Memory Status --------
    $reportContent += "<h2 class='section-title'>Memory Status</h2>"
    $memoryInfo = Get-WmiObject -Class Win32_PhysicalMemory
    $reportContent += "<table><tr><th>Device</th><th>Capacity (GB)</th><th>Speed (MHz)</th><th>Status</th></tr>"
    foreach ($mem in $memoryInfo) {
        $capacityGB = [math]::round($mem.Capacity / 1GB, 2)
        $reportContent += "<tr><td>$($mem.DeviceLocator)</td><td>$capacityGB</td><td>$($mem.Speed)</td><td>$($mem.Status)</td></tr>"
    }
    $reportContent += "</table>"

    # -------- Disk Errors (CHKDSK) --------
    $reportContent += "<h2 class='section-title'>Disk Errors (CHKDSK)</h2>"
    $drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
    foreach ($drive in $drives) {
        $driveLetter = $drive.DeviceID
        $reportContent += "<p>Running CHKDSK on drive $driveLetter...</p>"
        $chkdskResult = cmd /c "chkdsk $driveLetter /scan"
        $reportContent += "<pre>$chkdskResult</pre>"
    }

    # -------- Installed Software --------
    $reportContent += "<h2 class='section-title'>Installed Software</h2>"
    $softwareList = Get-WmiObject -Class Win32_Product | Select-Object Name, Version, Vendor
    $reportContent += "<table><tr><th>Name</th><th>Version</th><th>Vendor</th></tr>"
    foreach ($software in $softwareList) {
        $reportContent += "<tr><td>$($software.Name)</td><td>$($software.Version)</td><td>$($software.Vendor)</td></tr>"
    }
    $reportContent += "</table>"

    # -------- Windows Event Log Errors --------
    $reportContent += "<h2 class='section-title'>Windows Event Log Errors</h2>"
    $errorEvents = Get-WinEvent -LogName System -FilterHashtable @{Level=2} -MaxEvents 10
    if ($errorEvents.Count -eq 0) {
        $reportContent += "<p>No recent critical system errors found.</p>"
    } else {
        foreach ($event in $errorEvents) {
            $reportContent += "<p>Error ID: $($event.Id), Time: $($event.TimeCreated), Message: $($event.Message)</p>"
        }
    }

    # -------- High CPU/Memory Usage Processes --------
    $reportContent += "<h2 class='section-title'>High CPU/Memory Usage Processes</h2>"
    $processes = Get-Process | Where-Object { $_.CPU -gt 80 -or $_.WorkingSet -gt 1GB } | Select-Object Name, CPU, WorkingSet
    foreach ($process in $processes) {
        $reportContent += "<p>Process: $($process.Name), CPU: $($process.CPU), Memory: $([math]::round($process.WorkingSet / 1MB, 2)) MB</p>"
    }

    # Generate HTML report file
    $htmlPath = "$env:USERPROFILE\Desktop\SystemScanReport.html"
    Generate-HTMLReport -htmlPath $htmlPath -content $reportContent

    # Automatically open the generated HTML report in the default web browser
    Start-Process $htmlPath
}

# Run the system scan
Run-SystemScan