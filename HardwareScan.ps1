# Add Windows Forms .NET assembly
Add-Type -AssemblyName System.Windows.Forms

# Function to generate HTML content
function Generate-HTMLReport {
    param (
        [string]$htmlPath,
        [string]$content
    )

    # EXACT Battery Health Theme: White bg, Black text, #0078D4 Blue
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
            background-color: #ffffff;
            margin: 20px;
            color: #000000;
        }
        h1 {
            color: #0078D4;
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
            background-color: #0078D4;
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

    $htmlContent | Out-File -Encoding UTF8 $htmlPath
}

function Run-HardwareScan {
    $reportContent = "<h1>Hardware Scan Report</h1>"
    $reportContent += "<h2 class='section-title'>Physical Disk Health</h2>"
    $disks = Get-PhysicalDisk | Select-Object DeviceID, MediaType, OperationalStatus, HealthStatus
    $reportContent += "<table><tr><th>DeviceID</th><th>MediaType</th><th>Status</th><th>Health</th></tr>"
    foreach ($disk in $disks) {
        $reportContent += "<tr><td>$($disk.DeviceID)</td><td>$($disk.MediaType)</td><td>$($disk.OperationalStatus)</td><td>$($disk.HealthStatus)</td></tr>"
    }
    $reportContent += "</table>"

    $reportContent += "<h2 class='section-title'>Memory Status</h2>"
    $memoryInfo = Get-WmiObject -Class Win32_PhysicalMemory
    $reportContent += "<table><tr><th>Device</th><th>Capacity (GB)</th><th>Speed (MHz)</th></tr>"
    foreach ($mem in $memoryInfo) {
        $cap = [math]::round($mem.Capacity / 1GB, 2)
        $reportContent += "<tr><td>$($mem.DeviceLocator)</td><td>$cap</td><td>$($mem.Speed)</td></tr>"
    }
    $reportContent += "</table>"
    return $reportContent
}

$reportContent = Run-HardwareScan
$htmlPath = "$env:USERPROFILE\Desktop\HardwareScanReport.html"
Generate-HTMLReport -htmlPath $htmlPath -content $reportContent

if ([System.Windows.Forms.MessageBox]::Show("Scan Complete. View HTML report?", "The IT Guy", [System.Windows.Forms.MessageBoxButtons]::YesNo) -eq "Yes") {
    Start-Process $htmlPath
}
