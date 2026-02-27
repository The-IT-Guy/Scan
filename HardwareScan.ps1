# Add Windows Forms .NET assembly
Add-Type -AssemblyName System.Windows.Forms

# Function to generate HTML content
function Generate-HTMLReport {
    param (
        [string]$htmlPath,
        [string]$content
    )

    # HTML structure matching theme with white background and black text, title removed
    $htmlContent = @"
<!DOCTYPE html>
<html lang='en'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title></title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #ffffff; /* White background */
            margin: 20px;
            color: #000000; /* Black text */
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
    <div>
        $content
    </div>
</body>
</html>
"@

    # Write the HTML content to the specified file path
    $htmlContent | Out-File -Encoding UTF8 $htmlPath
# Function to run all the hardware system checks
function Run-HardwareScan {
    $reportContent = ""

    # -------- System Information --------
    $reportContent += "<div class='section-title'>System Information</div>"
    $computerSystem = Get-WmiObject -Class Win32_ComputerSystem
    $os = Get-WmiObject -Class Win32_OperatingSystem
    $bios = Get-WmiObject -Class Win32_BIOS
    
    if ($computerSystem -and $os -and $bios) {
        $reportContent += "<p class='info-p'><strong>Computer Name:</strong> $($computerSystem.Name)</p>"
        $reportContent += "<p class='info-p'><strong>Manufacturer:</strong> $($computerSystem.Manufacturer)</p>"
        $reportContent += "<p class='info-p'><strong>Model:</strong> $($computerSystem.Model)</p>"
        $reportContent += "<p class='info-p'><strong>BIOS Version:</strong> $($bios.SMBIOSBIOSVersion)</p>"
        $reportContent += "<p class='info-p'><strong>OS Name:</strong> $($os.Caption) ($($os.Version))</p>"
    }

    # -------- Physical Disk Health --------
    $reportContent += "<div class='section-title'>Physical Disk Health</div>"
    $disks = Get-PhysicalDisk | Select-Object DeviceID, MediaType, OperationalStatus, HealthStatus
    if ($disks) {
        $reportContent += "<table><tr><th>DeviceID</th><th>MediaType</th><th>OperationalStatus</th><th>HealthStatus</th></tr>"
        foreach ($disk in $disks) {
            $reportContent += "<tr><td>$($disk.DeviceID)</td><td>$($disk.MediaType)</td><td>$($disk.OperationalStatus)</td><td>$($disk.HealthStatus)</td></tr>"
        }
        $reportContent += "</table>"
    }

    # -------- Memory Status --------
    $reportContent += "<div class='section-title'>Memory Status</div>"
    $memoryInfo = Get-WmiObject -Class Win32_PhysicalMemory
    if ($memoryInfo) {
        $reportContent += "<table><tr><th>Device</th><th>Capacity (GB)</th><th>Speed (MHz)</th><th>Status</th></tr>"
        foreach ($mem in $memoryInfo) {
            $capacityGB = [math]::round($mem.Capacity / 1GB, 2)
            $reportContent += "<tr><td>$($mem.DeviceLocator)</td><td>$capacityGB GB</td><td>$($mem.Speed) MHz</td><td>$($mem.Status)</td></tr>"
        }
        $reportContent += "</table>"
    }

    # -------- Disk Errors (CHKDSK) --------
    $reportContent += "<div class='section-title'>Disk Errors (CHKDSK Scan)</div>"
    $drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
    if ($drives) {
        foreach ($drive in $drives) {
            $driveLetter = $drive.DeviceID
            $reportContent += "<p class='info-p'>Results for Drive $driveLetter :</p>"
            $chkdskResult = cmd /c "chkdsk $driveLetter /scan" | Out-String
            $reportContent += "<pre>$chkdskResult</pre>"
        }
    }

    return $reportContent
}

# 1. Execute Scan
$reportContent = Run-HardwareScan

# 2. Save Report
$htmlPath = "$env:USERPROFILE\Desktop\HardwareScanReport.html"
Generate-HTMLReport -htmlPath $htmlPath -content $reportContent

# 3. Prompt User to View Report
$title = "Scan Complete"
$message = "Hardware Scan is finished. Would you like to view the HTML report now?"
$buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
$icon = [System.Windows.Forms.MessageBoxIcon]::Information

$result = [System.Windows.Forms.MessageBox]::Show($message, $title, $buttons, $icon)

if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
    Start-Process $htmlPath
} else {
    Write-Host "Report saved to Desktop as HardwareScanReport.html"
}


