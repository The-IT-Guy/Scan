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
    <title>Hardware Scan Report</title>
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
    <h1>Hardware Scan Report</h1>
    <div>
        $content
    </div>
</body>
</html>
"@

    # Write the HTML content to the specified file path
    $htmlContent | Out-File -Encoding UTF8 $htmlPath
}

# Function to run all the hardware system checks
function Run-HardwareScan {
    $reportContent = ""

    # -------- Physical Disk Health --------
    $reportContent += "<h2 class='section-title'>Physical Disk Health</h2>"
    $disks = Get-PhysicalDisk | Select-Object DeviceID, MediaType, OperationalStatus, HealthStatus
    if ($disks) {
        $reportContent += "<table><tr><th>DeviceID</th><th>MediaType</th><th>OperationalStatus</th><th>HealthStatus</th></tr>"
        foreach ($disk in $disks) {
            $reportContent += "<tr><td>$($disk.DeviceID)</td><td>$($disk.MediaType)</td><td>$($disk.OperationalStatus)</td><td>$($disk.HealthStatus)</td></tr>"
        }
        $reportContent += "</table>"
    } else {
        $reportContent += "<p>No physical disk data available.</p>"
    }

    # -------- Memory Status --------
    $reportContent += "<h2 class='section-title'>Memory Status</h2>"
    $memoryInfo = Get-WmiObject -Class Win32_PhysicalMemory
    if ($memoryInfo) {
        $reportContent += "<table><tr><th>Device</th><th>Capacity (GB)</th><th>Speed (MHz)</th><th>Status</th></tr>"
        foreach ($mem in $memoryInfo) {
            $capacityGB = [math]::round($mem.Capacity / 1GB, 2)
            $reportContent += "<tr><td>$($mem.DeviceLocator)</td><td>$capacityGB</td><td>$($mem.Speed)</td><td>$($mem.Status)</td></tr>"
        }
        $reportContent += "</table>"
    } else {
        $reportContent += "<p>No memory information available.</p>"
    }

    # -------- System Information --------
    $reportContent += "<h2 class='section-title'>System Information</h2>"
    $computerSystem = Get-WmiObject -Class Win32_ComputerSystem
    $os = Get-WmiObject -Class Win32_OperatingSystem
    $bios = Get-WmiObject -Class Win32_BIOS
    if ($computerSystem -and $os -and $bios) {
        $reportContent += "<p>Computer Name: $($computerSystem.Name)</p>"
        $reportContent += "<p>Manufacturer: $($computerSystem.Manufacturer)</p>"
        $reportContent += "<p>Model: $($computerSystem.Model)</p>"
        $reportContent += "<p>BIOS Version: $($bios.SMBIOSBIOSVersion)</p>"
        $reportContent += "<p>OS Name: $($os.Caption) $($os.Version)</p>"
    } else {
        $reportContent += "<p>System information not available.</p>"
    }

    # -------- Disk Errors (CHKDSK) --------
    $reportContent += "<h2 class='section-title'>Disk Errors (CHKDSK)</h2>"
    $drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
    if ($drives) {
        foreach ($drive in $drives) {
            $driveLetter = $drive.DeviceID
            $reportContent += "<p>Running CHKDSK on drive $driveLetter...</p>"
            $chkdskResult = cmd /c "chkdsk $driveLetter /scan"
            $reportContent += "<pre>$chkdskResult</pre>"
        }
    } else {
        $reportContent += "<p>No drives found for CHKDSK.</p>"
    }

    return $reportContent
}

# Run the hardware scan and generate the HTML report
$reportContent = Run-HardwareScan
$htmlPath = "$env:USERPROFILE\Desktop\HardwareScanReport.html"
Generate-HTMLReport -htmlPath $htmlPath -content $reportContent

# Open the report in the default browser
Start-Process $htmlPath

Write-Host "Hardware Scan Report has been generated and saved to your desktop."