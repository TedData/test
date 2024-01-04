
# Parameter that need modification
$sqlUsername = "SuperAdmin"
$sqlPassword = "SuperAdmin"
$output_path = "C:\Users\Peng Yu\Downloads"
$serverName = "DESKTOP-46A7LA5"
$databaseName = "PBI_Inventory"
$tables = @()
#$tables = @()  # download whole tables in the database





$connStr = "Server=$serverName;Database=$databaseName;User Id=$sqlUsername;Password=$sqlPassword"
$connection = New-Object System.Data.SqlClient.SqlConnection($connStr)
$connection.Open()
$TodaysDate = Get-Date -Format "yyyyMMdd"
try {
    if (-not $tables) {
        $queryTable = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'"
        $tables = Invoke-Sqlcmd -ServerInstance $serverName -Database $databaseName -Username $sqlUsername -Password $sqlPassword -Query $queryTable -TrustServerCertificate | Select-Object -ExpandProperty TABLE_NAME
    } 
    foreach ($table in $tables) {
        $outputFile = "$output_path\$table"+"_$TodaysDate.csv"
        $query = "SELECT * FROM $table"
        $command = $connection.CreateCommand()
        $command.CommandText = $query
        $dataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $command
        $dataTable = New-Object System.Data.DataTable
        $dataAdapter.Fill($dataTable)
        $dataTable | Export-Csv -Path $outputFile -NoTypeInformation
   }
    Write-Host "Exported table names to $outputFile."
}
catch {
    Write-Host "Error: $_"
}
