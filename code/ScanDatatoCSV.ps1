<#
.SYNOPSIS
This PowerShell script scans Power BI workspaces, dashboards, reports, 
datasets, tables, columns, and measures, and exports the scanned 
information to CSV files.

.DESCRIPTION
The script utilizes the Power BI Management module to retrieve 
information about workspaces and their contents and then processes 
the obtained data to generate CSV files for dashboards, reports, 
datasets, tables, columns, and measures.

.PARAMETER outputPath
Specifies the path where the CSV files will be saved.

#>

param (
    [string]$outputPath = "C:\Users\Peng Yu\Downloads"
)

# Function to install or update a PowerShell module
function Install-OrUpdate-Module([string]$ModuleName) {
    $module = Get-Module $ModuleName -ListAvailable -ErrorAction SilentlyContinue

    # Check if the module is not installed or if its version is not within the specified range
    if (!$module -or ($module.Version -ne '1.0.0' -and $module.Version -le '1.0.410')) {
        $action = if (!$module) { "Installing" } else { "Updating" }
        Write-Host "$action module $ModuleName ..."
        Install-Module -Name $ModuleName -Force -Scope CurrentUser
        Write-Host "Module $ModuleName $action complete"
    }
}

# Install or update the required PowerShell module for Power BI management
Install-OrUpdate-Module -ModuleName "MicrosoftPowerBIMgmt"
Connect-PowerBIServiceAccount

# Arrays to store scanned information
$ScannerDashboards = @()
$ScannerReportsAndDatasets = @()
$ScannerTablesAndColumns = @()
$ScannerMeasures = @()
$dataset_unique = @()

# Get all workspaces in the organization
$allWorkspaces = Get-PowerBIWorkspace -Scope Organization

# Loop through each workspace
ForEach ($workspace in ($allWorkspaces)) {
    # Prepare the request body for scanning workspace details
    $body = "{'workspaces': ['$($workspace.Id)']}"
    $url_getscanId = "https://api.powerbi.com/v1.0/myorg/admin/workspaces/getInfo?lineage=True&datasourceDetails=True&datasetSchema=True&datasetExpressions=True"
    $GetScanId = Invoke-PowerBIRestMethod -Url $url_getscanId -Method Post -Body $body | ConvertFrom-Json

    # Get the scan results for the workspace
    $url = "https://api.powerbi.com/v1.0/myorg/admin/workspaces/scanResult/$($GetScanId.id)"
    $GetScanResult = Invoke-PowerBIRestMethod -Url $url -Method Get | ConvertFrom-Json

    # Extract workspace information
    $workspaceId = $GetScanResult.workspaces.id
    $workspaceName = $GetScanResult.workspaces.name 
    $type = $GetScanResult.workspaces.type 
    $state = $GetScanResult.workspaces.state 
    $isOnDedicatedCapacity = $GetScanResult.workspaces.isOnDedicatedCapacity

    Write-Host $workspaceName

    # Process dashboards in the workspace
    if ($GetScanResult.workspaces.dashboards.Length -ne 0){
        foreach ($dashboard in $GetScanResult.workspaces.dashboards){
            $dashboardsId = $dashboard.id
            $displayName = $dashboard.displayName
            $isReadOnly = $dashboard.isReadOnly
            $dashboard_sensitivityLabel_labelId = $dashboard.sensitivityLabel.labelId

            if ($dashboard.tiles.Length -ne 0){
                # Process tiles in the dashboard
                foreach ($tile in $dashboard.tiles){
                    $tiles_id = $tile.id
                    $tiles_title = $tile.title
                    $tiles_reportId = $tile.reportId
                    $tiles_datasetId = $tile.datasetId 

                    # Create an object to store dashboard information
                    $ScannerDashboardsInfo = [PSCustomObject]@{
                        "workspaceId" = $workspaceId
                        "dashboardsId" = $dashboardsId
                        "displayName" = $displayName
                        "isReadOnly" = $isReadOnly
                        "tiles_id" = $tiles_id
                        "tiles_title" = $tiles_title
                        "tiles_reportId" = $tiles_reportId
                        "tiles_datasetId" = $tiles_datasetId
                        "dashboard_sensitivityLabel" = $dashboard_sensitivityLabel_labelId
                    }

                    # Add the dashboard information to the array
                    $ScannerDashboards += $ScannerDashboardsInfo
                }
            } 
            else {
                # If no tiles in the dashboard
                $ScannerDashboardsInfo = [PSCustomObject]@{
                    "workspaceId" = $workspaceId
                    "dashboardsId" = $dashboardsId
                    "displayName" = $displayName
                    "isReadOnly" = $isReadOnly
                    "tiles_id" = ""
                    "tiles_title" = ""
                    "tiles_reportId" = ""
                    "tiles_datasetId" = ""
                    "dashboard_sensitivityLabel" = $dashboard_sensitivityLabel_labelId
                }

                # Add the dashboard information to the array
                $ScannerDashboards += $ScannerDashboardsInfo
            }
        }
    }

    # Process reports in the workspace
    foreach ($report in $GetScanResult.workspaces.reports){
        $reportId = $report.id
        $reportName = $report.name
        $report_datasetsId = $report.datasetId
        $createdDateTime = $report.createdDateTime
        $modifiedDateTime = $report.modifiedDateTime
        $modifiedBy = $report.modifiedBy
        $reportType = $report.reportType
        $createdBy = $report.createdBy
        $modifiedById = $report.modifiedById
        $createdById = $report.createdById
        $endorsement = $report.endorsementDetails.endorsement
        $endorsement_certifiedBy = $report.endorsementDetails.certifiedBy
        $report_sensitivityLabel = $report.sensitivityLabel.labelId

        # Handle the case where there are no datasets associated with the report
        if ($report_datasetsId.Count -eq 0){
            $report_datasetsId = "wrong_report_dataset_Id"
        }

        # Loop through each dataset associated with the report
        foreach ($report_datasetId in $report_datasetsId){
            $getDatasets = $GetScanResult.workspaces.datasets | Where-Object { $_.id -eq $report_datasetId}

            # Handle the case where there are no datasets found
            if ($getDatasets.count -eq 0){
                $getDatasets = "wrong_dataset"
            }

            # Loop through each dataset
            foreach ($dataset in $getDatasets){
                $datasetsId = $dataset.id
                $datasetsName = $dataset.name
                $configuredBy = $dataset.configuredBy
                $configuredById = $dataset.configuredById
                $isEffectiveIdentityRequired = $dataset.isEffectiveIdentityRequired
                $isEffectiveIdentityRolesRequired = $dataset.isEffectiveIdentityRolesRequired
                $targetStorageMode = $dataset.targetStorageMode
                $createdDate_Dataset = $dataset.createdDate
                $contentProviderType = $dataset.contentProviderType
                $datasourceUsages_datasourceInstanceId = $dataset.datasourceUsages.datasourceInstanceId

                # Handle the case where there are no datasource instances associated with the dataset
                if ($datasourceUsages_datasourceInstanceId.Count -eq 0){
                    $datasourceUsages_datasourceInstanceId = "wrong_datasourceInstance_Id"
                }

                # Loop through each datasource instance
                foreach ($datasourceInstanceId in $datasourceUsages_datasourceInstanceId) {
                    $getDatasources = $GetScanResult.datasourceInstances | Where-Object { $_.datasourceId -eq $datasourceInstanceId }

                    # Handle the case where there are no datasource instances found
                    if ($getDatasources.count -eq 0){
                        $getDatasources = "wrong_datasource"
                    }

                    # Loop through each datasource
                    foreach ($datasource in $getDatasources){
                        $datasourceId = $datasource.datasourceId
                        $datasourceType = $datasource.datasourceType
                        $connectionDetails = $datasource.connectionDetails
                        $gatewayId = $datasource.gatewayId

                        # Create an object to store report and dataset information
                        $ScannerReportsAndDatasetsInfo = [PSCustomObject]@{
                            "workspaceId" = $workspaceId
                            "workspaceName" = $workspaceName
                            "reportId" = $reportId
                            "reportName" = $reportName
                            "datasetsId" = $datasetsId
                            "datasetsName" = $datasetsName
                            "datasourceId" = $datasourceId
                            "gatewayId" = $gatewayId
                            "createdBy" = $createdBy
                            "createdDate_Dataset" = $createdDate_Dataset
                            "createdDateTime" = $createdDateTime
                            "createdById" = $createdById
                            "modifiedBy" = $modifiedBy
                            "modifiedDateTime" = $modifiedDateTime
                            "modifiedById" = $modifiedById
                            "type" = $type
                            "state" = $state
                            "isOnDedicatedCapacity" = $isOnDedicatedCapacity
                            "isEffectiveIdentityRequired" = $isEffectiveIdentityRequired
                            "isEffectiveIdentityRolesRequired" = $isEffectiveIdentityRolesRequired
                            "reportType" = $reportType
                            "datasourceType" = $datasourceType
                            "targetStorageMode" = $targetStorageMode
                            "configuredBy" = $configuredBy
                            "configuredById" = $configuredById
                            "contentProviderType" = $contentProviderType
                            "connectionDetails" = $connectionDetails
                            "endorsement" = $endorsement
                            "endorsement_certifiedBy" = $endorsement_certifiedBy
                            "report_sensitivityLabel" = $report_sensitivityLabel
                        }

                        # Add the report and dataset information to the array
                        $ScannerReportsAndDatasets += $ScannerReportsAndDatasetsInfo
                    }
                }

                if ($dataset_unique -notcontains $datasetsId) {
                    $dataset_unique += $datasetsId

                    # Process tables in the dataset
                    if ($dataset.tables.Length -ne 0){
                        foreach ($table in $dataset.tables){
                            $tablesName = $table.name
                            $source = $table.source.expression -replace '\s{2,}', ' ' -replace '"', "``" -replace "'", "``"
                            $isHidden = $table.isHidden
                            $tableDescription = $table.description

                            # Process measures in the table
                            if ($table.measures.Length -ne 0){
                                foreach ($measure in $table.measures){
                                    $measureName = $measure.name
                                    $measureExpression = $measure.expression -replace '\s{2,}', ' ' -replace '"', "``" -replace "'", "``"
                                    $measureDescription = $measure.description

                                    # Create an object to store measure information
                                    $ScannerMeasuresInfo = [PSCustomObject]@{
                                        "datasetsId" = $datasetsId
                                        "datasetsName" = $datasetsName
                                        "tablesName" = $tablesName
                                        "measureName" = $measureName
                                        "measureDescription" = $measureDescription
                                        "measureExpression" = $measureExpression
                                    }

                                    # Add the measure information to the array
                                    $ScannerMeasures += $ScannerMeasuresInfo
                                }
                            }

                            # Process columns in the table
                            if ($table.columns.Length -ne 0){
                                foreach ($column in $table.columns){
                                    $columnsName = $column.name
                                    $columnsDataType = $column.dataType
                                    $columnsIsHidden = $column.isHidden
                                    $columnsType = $column.columnType

                                    # Create an object to store table and column information
                                    $ScannerTablesAndColumnsInfo = [PSCustomObject]@{
                                        "datasetsId" = $datasetsId
                                        "datasetsName" = $datasetsName
                                        "tablesName" = $tablesName
                                        "tableDescription" = $tableDescription
                                        "columns_name" = $columnsName
                                        "columns_dataType" = $columnsDataType
                                        "isHidden" = $columnsIsHidden
                                        "columns_type" = $columnsType
                                        "source" = $source
                                    }

                                    # Add the table and column information to the array
                                    $ScannerTablesAndColumns += $ScannerTablesAndColumnsInfo
                                }
                            } 
                            # If no columns in the table
                            else {
                                $ScannerTablesAndColumnsInfo = [PSCustomObject]@{
                                    "datasetsId" = $datasetsId
                                    "datasetsName" = $datasetsName
                                    "tablesName" = $tablesName
                                    "tableDescription" = $tableDescription
                                    "columns_name" = ""
                                    "columns_dataType" = ""
                                    "isHidden" = $isHidden
                                    "columns_type" = ""
                                    "source" = $source
                                }

                                # Add the table information to the array
                                $ScannerTablesAndColumns += $ScannerTablesAndColumnsInfo   
                            }
                        }
                    }
                }
            }
        }
    }
}

# Export scanned information to CSV files
$ScannerDashboards | Export-Csv -Path "$outputPath\PBI_ScannerDashboards.csv" -NoTypeInformation
$ScannerReportsAndDatasets | Export-Csv -Path "$outputPath\PBI_ScannerReportsAndDatasets.csv" -NoTypeInformation
$ScannerTablesAndColumns | Export-Csv -Path "$outputPath\PBI_ScannerTablesAndColumns.csv" -NoTypeInformation
$ScannerMeasures | Export-Csv -Path "$outputPath\PBI_ScannerMeasures.csv" -NoTypeInformation

Write-Host "completed"

