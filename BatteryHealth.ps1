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
}

# Function to check battery health and generate an HTML report
function Check-BatteryHealth {
    # Define the output file path for the battery report
    $batteryReportPath = "$env:USERPROFILE\Desktop\battery-report.html"

    # Generate the battery report using powercfg
    powercfg /batteryreport /output $batteryReportPath

    # Check if the report was generated successfully
    if (Test-Path $batteryReportPath) {
        # Read the content of the battery report
        $reportContent = Get-Content $batteryReportPath -Raw

        # Generate an HTML formatted report with consistent styling
        $htmlPath = "$env:USERPROFILE\Desktop\BatteryHealthReport.html"
        Generate-HTMLReport -htmlPath $htmlPath -content $reportContent

        # Open the generated HTML report in the default web browser
        Start-Process $htmlPath
    } else {
        [System.Windows.Forms.MessageBox]::Show("Failed to generate battery report.", "Error")
    }
}

# Execute the battery health check
Check-BatteryHealth