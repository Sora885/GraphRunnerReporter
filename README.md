# GraphRunnerReporter

**GraphRunnerReporter** is a PowerShell script designed to interact with Microsoft Graph API. It allows you to extract data from various endpoints, validate token permissions, and generate detailed reports in JSON or HTML format.

## Features

- **Token Permission Validation**: Checks if the provided token includes the necessary scopes for each endpoint.
- **Data Export**: Fetches data from Microsoft Graph API endpoints like `me`, `messages`, `drive`, and more.
- **Reports**: Generates reports in:
  - **JSON**: Consolidated data in JSON format.
  - **HTML**: Interactive and styled HTML report with statuses and fetched data.

## Prerequisites

1. **PowerShell**: Ensure PowerShell is installed on your system.
2. **Microsoft Graph Token**:
   - Obtain a token with appropriate scopes using tools like [GraphRunner](https://github.com/dafthack/GraphRunner) or Azure CLI.
3. **Required Permissions**:
   - Ensure your token has the necessary permissions for the endpoints you want to query (see [Endpoints and Permissions](#endpoints-and-permissions)).

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/GraphRunnerReporter.git
   cd GraphRunnerReporter

## Usage Guide

### Step 1: Download and Run GraphRunner

To start, download and execute **GraphRunner** to obtain your Microsoft Graph API token:

```powershell
IEX (iwr 'https://raw.githubusercontent.com/dafthack/GraphRunner/main/GraphRunner.ps1')
```
### Step 2: Authenticate and Obtain Tokens

Run the following command to authenticate and retrieve your tokens:
```
Get-GraphTokens
```
### Step 3: Run GraphRunnerReporter

Once the $tokens variable is defined, use the following commands to generate your reports:
## HTML Report
```
.\Export-GraphData.ps1 -Tokens $tokens -OutputFilePath "GraphReport.html" -ReportType "HTML"
```
## JSON Report
```
.\Export-GraphData.ps1 -Tokens $tokens -OutputFilePath "ExportedData.json" -ReportType "JSON"
```

### Parameters
```
-Tokens : Object containing the access_token and scope.
-OutputFilePath : Path to save the generated report.
-ReportType : Report format (JSON or HTML)
```
### Endpoint and Permissions

Endpoint : https://graph.microsoft.com/v1.0/me
Permissions requises : User.Read, User.ReadWrite, User.Read.All

Endpoint : https://graph.microsoft.com/v1.0/me/messages
Permissions requises : Mail.Read, Mail.ReadWrite

Endpoint : https://graph.microsoft.com/v1.0/me/drive/root/children
Permissions requises : Files.Read, Files.ReadWrite

Endpoint : https://graph.microsoft.com/v1.0/security/alerts
Permissions requises : SecurityEvents.Read.All, SecurityEvents.ReadWrite.All

Endpoint : https://graph.microsoft.com/v1.0/groups
Permissions requises : Group.Read.All, Group.ReadWrite.All

Endpoint : https://graph.microsoft.com/v1.0/applications
Permissions requises : Application.Read.All, Application.ReadWrite.OwnedBy

Endpoint : https://graph.microsoft.com/v1.0/policies/conditionalAccessPolicies
Permissions requises : Policy.Read.All, Policy.ReadWrite.ConditionalAccess

