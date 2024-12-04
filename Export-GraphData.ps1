param (
    [Parameter(Mandatory = $true)]
    [object]$Tokens,

    [Parameter(Mandatory = $true)]
    [string]$OutputFilePath,

    [Parameter(Mandatory = $false)]
    [string]$ReportType = "JSON"
)

function Invoke-GraphRequest {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Endpoint,
        [string]$AccessToken
    )

    $Headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }

    try {
        $Response = Invoke-RestMethod -Uri $Endpoint -Headers $Headers -Method GET
        return $Response
    } catch {
        if ($_.Exception.Response.StatusCode -eq 429) {
            $RetryAfter = $_.Exception.Response.Headers["Retry-After"]
            Write-Warning "Rate limit reached. Retrying after $RetryAfter seconds..."
            Start-Sleep -Seconds $RetryAfter
            return Invoke-GraphRequest -Endpoint $Endpoint -AccessToken $AccessToken
        } elseif ($_.Exception.Response.StatusCode -eq 403) {
            Write-Warning "Access forbidden for $Endpoint. Check permissions."
            return $null
        } else {
            Write-Error "Failed to fetch data from ${Endpoint}: $_"
            return $null
        }
    }
}

$RequiredScopes = @{
    "https://graph.microsoft.com/v1.0/me"                    = @("User.Read", "User.ReadWrite", "User.Read.All")
    "https://graph.microsoft.com/v1.0/me/messages"           = @("Mail.Read", "Mail.ReadWrite")
    "https://graph.microsoft.com/v1.0/me/drive/root/children" = @("Files.Read", "Files.ReadWrite")
    "https://graph.microsoft.com/v1.0/security/alerts"       = @("SecurityEvents.Read.All", "SecurityEvents.ReadWrite.All")
    "https://graph.microsoft.com/v1.0/groups"                = @("Group.Read.All", "Group.ReadWrite.All")
    "https://graph.microsoft.com/v1.0/applications"          = @("Application.Read.All", "Application.ReadWrite.OwnedBy")
    "https://graph.microsoft.com/v1.0/policies/conditionalAccessPolicies" = @("Policy.Read.All", "Policy.ReadWrite.ConditionalAccess")
}

$TokenScopes = $Tokens.scope -split " "

$ExportedData = @{}
$ReportData = @()

foreach ($Endpoint in $RequiredScopes.Keys) {
    Start-Sleep -Milliseconds 500
    $MissingScopes = @()
    foreach ($Scope in $RequiredScopes[$Endpoint]) {
        if (-not ($TokenScopes -contains $Scope)) {
            $MissingScopes += $Scope
        }
    }

    if ($MissingScopes.Count -eq 0) {
        Write-Output "Success: $Endpoint - All required permissions are present."
        Write-Output "Fetching data from $Endpoint..."
        $Data = Invoke-GraphRequest -Endpoint $Endpoint -AccessToken $Tokens.access_token
        if ($Data) {
            $ExportedData[$Endpoint] = $Data
            $ReportData += @{
                Endpoint = $Endpoint
                Status   = "Success"
                Data     = $Data | ConvertTo-Json -Depth 10
            }
        } else {
            Write-Warning "No data retrieved for $Endpoint."
            $ReportData += @{
                Endpoint = $Endpoint
                Status   = "No Data"
                Data     = "N/A"
            }
        }
    } else {
        Write-Warning "Warning: $Endpoint - Missing permissions: $($MissingScopes -join ', ')"
        $ReportData += @{
            Endpoint = $Endpoint
            Status   = "Missing Permissions"
            Data     = $MissingScopes -join ', '
        }
    }
}

if ($ReportType -eq "HTML") {
    Write-Output "Generating HTML report..."
    $HTMLReport = @"
<html>
<head>
    <title>Microsoft Graph API Export Report</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f4f4f9;
            color: #333;
        }
        h1 {
            color: #4CAF50;
            text-align: center;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin-top: 20px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
            background-color: #fff;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 12px;
            text-align: left;
        }
        th {
            background-color: #4CAF50;
            color: white;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        tr:hover {
            background-color: #f1f1f1;
        }
        .status-success {
            color: #4CAF50;
            font-weight: bold;
        }
        .status-missing {
            color: #E53935;
            font-weight: bold;
        }
        .status-no-data {
            color: #FFB300;
            font-weight: bold;
        }
        pre {
            font-family: Consolas, monospace;
            white-space: pre-wrap;
            word-wrap: break-word;
            background-color: #f8f8f8;
            padding: 10px;
            border-radius: 5px;
            overflow-x: auto;
        }
    </style>
</head>
<body>
    <h1>Microsoft Graph API Export Report</h1>
    <table>
        <thead>
            <tr>
                <th>Endpoint</th>
                <th>Status</th>
                <th>Data / Missing Permissions</th>
            </tr>
        </thead>
        <tbody>
"@

    foreach ($Entry in $ReportData) {
        $StatusClass = if ($Entry.Status -eq "Success") {
            "status-success"
        } elseif ($Entry.Status -eq "Missing Permissions") {
            "status-missing"
        } else {
            "status-no-data"
        }

        $HTMLReport += "<tr>"
        $HTMLReport += "<td>$($Entry.Endpoint)</td>"
        $HTMLReport += "<td class='$StatusClass'>$($Entry.Status)</td>"
        $HTMLReport += "<td><pre>$($Entry.Data)</pre></td>"
        $HTMLReport += "</tr>"
    }

    $HTMLReport += @"
        </tbody>
    </table>
</body>
</html>
"@

    $HTMLReport | Out-File -FilePath $OutputFilePath -Encoding UTF8
    Write-Output "HTML report saved to $OutputFilePath"
} elseif ($ReportType -eq "JSON") {
    Write-Output "Exporting data to JSON..."
    $ExportedData | ConvertTo-Json -Depth 10 | Out-File -FilePath $OutputFilePath -Encoding UTF8
    Write-Output "JSON report saved to $OutputFilePath"
} else {
    Write-Warning "Invalid report type specified. Use 'JSON' or 'HTML'."
}
