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
    <title>Installed Software Report</title>
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
    <h1>Installed Software Report</h1>
    <div>
        $content
    </div>
</body>
</html>
"@

    # Write the HTML content to the specified file path
    $htmlContent | Out-File -Encoding UTF8 $htmlPath
}

# Function to check compatibility of software version with OS version
function Check-Compatibility {
    param (
        [string]$softwareName,
        [string]$softwareVersion
    )

    # Example of potentially incompatible software version check
    if ($softwareName -like "*Flash*" -or $softwareVersion -like "1.*") {
        return "Potentially Incompatible"
    } else {
        return "Compatible"
    }
}

# Function to retrieve installed software from the system
function Get-InstalledSoftware {
    Write-Host "Retrieving list of installed software..." -ForegroundColor Cyan
    $softwareList = Get-WmiObject -Class Win32_Product | Select-Object Name, Version, Vendor

    $reportContent = "<h2 class='section-title'>Installed Software</h2>"
    $reportContent += "<table><tr><th>Name</th><th>Version</th><th>Vendor</th><th>Compatibility</th></tr>"

    foreach ($software in $softwareList) {
        if ($software.Name) {
            $compatibility = Check-Compatibility -softwareName $software.Name -softwareVersion $software.Version
            $reportContent += "<tr><td>$($software.Name)</td><td>$($software.Version)</td><td>$($software.Vendor)</td><td>$compatibility</td></tr>"
        }
    }
    $reportContent += "</table>"
    return $reportContent
}

# Function to check the Windows OS version
function Get-OSVersion {
    $os = Get-WmiObject -Class Win32_OperatingSystem
    $osSummary = "<h2 class='section-title'>Operating System Details</h2>"
    $osSummary += "<p>Current OS: $($os.Caption) Version: $($os.Version)</p>"
    return $osSummary
}

# Run OS version check and software scan
$osReport = Get-OSVersion
$softwareReport = Get-InstalledSoftware

# Combine the report content
$fullReport = $osReport + $softwareReport

# Generate HTML report file
$htmlPath = "$env:USERPROFILE\Desktop\SoftwareScanReport.html"
Generate-HTMLReport -htmlPath $htmlPath -content $fullReport

# Open the generated HTML report in the default web browser
Start-Process $htmlPath