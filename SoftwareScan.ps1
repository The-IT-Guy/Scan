# Add Windows Forms .NET assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to generate HTML content with "The IT Guy" Styling
function Generate-HTMLReport {
    param (
        [string]$htmlPath,
        [string]$content
    )

    $htmlContent = @"
<!DOCTYPE html>
<html lang='en'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>Software & Security Risk Report</title>
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
        .risk-high { color: #d9534f; font-weight: bold; }
        .risk-update { color: #f0ad4e; font-weight: bold; }
        .risk-safe { color: #5cb85c; }
        pre {
            background-color: #e9ecef;
            padding: 15px;
            border-radius: 5px;
            overflow-x: auto;
            white-space: pre-wrap;
            font-family: 'Consolas', monospace;
            border: 1px solid #ccc;
        }
    </style>
</head>
<body>
    <h1>Software & Security Risk Report</h1>
    <div>
        $content
    </div>
</body>
</html>
"@

    $htmlContent | Out-File -Encoding UTF8 $htmlPath
}

# Function to check for outdated software using Winget
function Get-OutdatedSoftware {
    $updateReport = "<div class='section-title'>Available Software Updates (via WinGet)</div>"
    try {
        # Check for updates and capture output
        $updates = winget upgrade | Out-String
        if ($updates -like "*No installed package have available updates*" -or $updates -eq "") {
            $updateReport += "<p class='risk-safe'>All tracked applications are up to date.</p>"
        } else {
            $updateReport += "<pre>$updates</pre>"
            $updateReport += "<p class='risk-update'>Action Required: Run 'winget upgrade --all' to update these items.</p>"
        }
    } catch {
        $updateReport += "<p>WinGet not found or unavailable for update scanning.</p>"
    }
    return $updateReport
}

# Function to check for risk/deprecated software
function Check-SoftwareRisk {
    param ($Name, $Vendor)
    
    # List of keywords for high-risk or deprecated software
    $riskKeywords = @("Flash Player", "Silverlight", "QuickTime", "uTorrent", "CCleaner", "Java 6", "Java 7", "Ask Toolbar")
    
    foreach ($risk in $riskKeywords) {
        if ($Name -like "*$risk*") {
            return "<span class='risk-high'>High Risk / Deprecated</span>"
        }
    }
    return "<span class='risk-safe'>Standard</span>"
}

# Function to retrieve installed software
function Get-InstalledSoftware {
    Write-Host "Scanning installed software and checking risks..." -ForegroundColor Cyan
    
    # Using registry for faster and more comprehensive results than Win32_Product
    $paths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
    $softwareList = Get-ItemProperty $paths | Select-Object DisplayName, DisplayVersion, Publisher | Where-Object { $_.DisplayName -ne $null } | Sort-Object DisplayName

    $reportContent = "<div class='section-title'>Software Inventory & Risk Audit</div>"
    $reportContent += "<table><tr><th>Name</th><th>Version</th><th>Publisher</th><th>Risk Status</th></tr>"

    foreach ($software in $softwareList) {
        $riskStatus = Check-SoftwareRisk -Name $software.DisplayName -Vendor $software.Publisher
        $reportContent += "<tr><td>$($software.DisplayName)</td><td>$($software.DisplayVersion)</td><td>$($software.Publisher)</td><td>$riskStatus</td></tr>"
    }
    $reportContent += "</table>"
    return $reportContent
}

# --- Execution ---

# 1. Gather Data
$os = Get-WmiObject -Class Win32_OperatingSystem
$osReport = "<div class='section-title'>Operating System Details</div><p>Current OS: $($os.Caption) (Version: $($os.Version))</p>"

$updateReport = Get-OutdatedSoftware
$inventoryReport = Get-InstalledSoftware

# 2. Generate and Save
$fullReport = $osReport + $updateReport + $inventoryReport
$htmlPath = "$env:USERPROFILE\Desktop\SoftwareRiskReport.html"
Generate-HTMLReport -htmlPath $htmlPath -content $fullReport

# 3. Prompt User
$title = "Software Scan Complete"
$message = "The Software Audit and Risk Scan is finished. Would you like to view the report now?"
$buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
$icon = [System.Windows.Forms.MessageBoxIcon]::Warning

$result = [System.Windows.Forms.MessageBox]::Show($message, $title, $buttons, $icon)

if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
    Start-Process $htmlPath
} else {
    Write-Host "Report saved to Desktop as SoftwareRiskReport.html"
}
