<#
.SYNOPSIS
This PowerShell script retrieves user access rights for Power BI workspaces 
and reports, and exports the information to a CSV file.

.DESCRIPTION
The script connects to the Power BI Service account, retrieves information 
about workspaces, and for each report in the workspace, fetches user access 
rights. The collected data is then exported to a CSV file for further analysis.

.PARAMETER outputPath
Specifies the path where the CSV file will be saved.

#>

# Define the default output path for the CSV file
param (
    [string]$outputPath = "C:\Users\Peng Yu\Downloads"
)

# Function to check if a PowerShell module exists and install/update it if necessary
function Assert-ModuleExists {
    param(
        [string]$ModuleName
    )

    $module = Get-Module $ModuleName -ListAvailable -ErrorAction SilentlyContinue

    if (!$module) {
        Write-Host "Installing module $ModuleName ..."
        Install-Module -Name $ModuleName -Force -Scope CurrentUser
        Write-Host "Module installed"
    }
    else {
        Write-Host "Module $ModuleName found."

        if ($module.Version -lt '1.0.0' -or $module.Version -le '1.0.410') {
            Write-Host "Updating module $ModuleName ..."
            Update-Module -Name $ModuleName -Force -ErrorAction Stop
            Write-Host "Module updated"
        }
    }
}

# Ensure that the required Power BI module is installed and updated
Assert-ModuleExists -ModuleName "MicrosoftPowerBIMgmt"

# Connect to Power BI Service Account
Connect-PowerBIServiceAccount

# Get all Power BI workspaces in the organization
$allWorkspaces = Get-PowerBIWorkspace -Scope Organization

# Initialize an array to store user access rights information
$userAccessRight = @()

# Loop through each workspace
ForEach ($workspace in ($allWorkspaces)) {
    # Prepare the body for the API request
    $body = "{'workspaces': ['$($workspace.Id)']}"
    $url_getscanId = "https://api.powerbi.com/v1.0/myorg/admin/workspaces/getInfo?lineage=True&datasourceDetails=True&datasetSchema=True&datasetExpressions=True"
    
    # Invoke API to get the scan ID
    $GetScanId = Invoke-PowerBIRestMethod -Url $url_getscanId -Method Post -Body $body | ConvertFrom-Json
    
    # Construct URL to get scan result based on scan ID
    $url = "https://api.powerbi.com/v1.0/myorg/admin/workspaces/scanResult/$($GetScanId.id)"
    
    # Invoke API to get the scan result
    $GetScanResult = Invoke-PowerBIRestMethod -Url $url -Method Get | ConvertFrom-Json

    # Extract workspace information
    $workspaceId = $GetScanResult.workspaces.id
    $workspaceName = $GetScanResult.workspaces.name

    # Loop through each report in the workspace
    foreach ($report in $GetScanResult.workspaces.reports) {
        $reportId = $report.id
        $reportName = $report.name
        Write-Host "report name is $($reportName)"
        
        # Construct URL to get users for a specific report
        $UrlReportId = "https://api.powerbi.com/v1.0/myorg/admin/reports/$reportId/users"
        
        # Invoke API to get users for the report
        $getUsers = Invoke-PowerBIRestMethod -Url $UrlReportId -Method GET | ConvertFrom-Json

        # Loop through each user for the report
        foreach ($getUser in $getUsers.value) {
            # Create a custom object to store user access rights information
            $userAccessRightInfo = [PSCustomObject]@{
                "workspaceId" = $workspaceId
                "workspaceName" = $workspaceName
                "reportId" = $reportId
                "reportName" = $reportName
                "reportUserAccessRight" = $getUser.reportUserAccessRight
                "displayName" = $getUser.displayName
                "identifier" = $getUser.identifier
                "graphId" = $getUser.graphId
                "principalType" = $getUser.principalType
                "userType" = $getUser.userType
            }

            # Add user access rights information to the array
            $userAccessRight += $userAccessRightInfo
        }
    }
}

# Export the user access rights information to a CSV file
$userAccessRight | Export-Csv -Path "$outputPath\PBI_WSUsersAccessRights.csv" -NoTypeInformation

Write-Host "Completed"
