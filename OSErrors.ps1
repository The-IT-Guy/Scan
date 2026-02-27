# Add Windows Forms .NET assembly
Add-Type -AssemblyName System.Windows.Forms

# Function to generate HTML content
function Generate-HTMLReport {
    param (
        [string]$htmlPath,
        [string]$content
    )

    # HTML structure matching EXACT theme with white background and black text
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
            color: #0056b3; /* Battery Report Blue */
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
            background-color: #0056b3; /* Battery Report Blue */
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

# --- Diagnostic Functions ---

function Check-OSHealth {
    $report = "<h1>Operating System Health Report</h1>"
    
    # SFC Scan
    $report += "<h2 class='section-title'>System File Checker (SFC)</h2>"
    $sfc = cmd /c "sfc /verifyonly" | Out-String
    $report += "<pre>$sfc</pre>"

    # Event Logs
    $report += "<h2 class='section-title'>Critical Event Log Errors</h2>"
    try {
        $events = Get-WinEvent -LogName System -FilterHashtable @{Level=1; StartTime=(Get-Date).AddDays(-7)} -MaxEvents 10 -ErrorAction SilentlyContinue
        if ($events) {
            $report += "<table><tr><th>Time</th><th>ID</th><th>Message</th></tr>"
            foreach ($e in $events) {
                $report += "<tr><td>$($e.TimeCreated)</td><td>$($e.Id)</td><td>$($e.Message)</td></tr>"
            }
            $reportContent += "</table>"
        } else {
            $report += "<p>No critical system errors found in the last 7 days.</p>"
        }
    } catch { $report += "<p>Unable to access event logs.</p>" }

    return $report
}

# --- Execution ---

$content = Check-OSHealth
$htmlPath = "$env:USERPROFILE\Desktop\OSErrorCheckReport.html"
Generate-HTMLReport -htmlPath $htmlPath -content $content

# User Prompt matching the professional toolkit flow
$title = "Scan Complete"
$message = "OS Error Check is finished. Would you like to view the HTML report now?"
$result = [System.Windows.Forms.MessageBox]::Show($message, $title, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)

if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
    Start-Process $htmlPath
}
