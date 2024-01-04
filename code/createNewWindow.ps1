Add-Type -AssemblyName System.Windows.Forms

$inputBox1 = New-Object System.Windows.Forms.TextBox
$inputBox1.Location = New-Object Drawing.Point @(420, 25)
$inputBox1.Size = New-Object Drawing.Size @(450, 2000)
$inputBox1.Multiline = $false
$inputBox1.Text = "C:\Users\Peng Yu\Downloads"

$inputBox2 = New-Object System.Windows.Forms.TextBox
$inputBox2.Location = New-Object Drawing.Point @(420, 245)
$inputBox2.Size = New-Object Drawing.Size @(450, 2000)
$inputBox2.Multiline = $false
$inputBox2.Text = "2023-12-06"

$form = New-Object Windows.Forms.Form
$form.Text = "TechnologyCue"
$form.Size = New-Object Drawing.Size @(950, 850)

$label1 = New-Object Windows.Forms.Label
$label1.Location = New-Object Drawing.Point @(10, 20)
$label1.Size = New-Object Drawing.Size @(400, 150)
$label1.Text = "Output Path:"

$label2 = New-Object Windows.Forms.Label
$label2.Location = New-Object Drawing.Point @(10, 240)
$label2.Size = New-Object Drawing.Size @(400, 150)
$label2.Text = "End Date:"

$form.Font = New-Object System.Drawing.Font("Arial", 26)  
$label1.Font = New-Object System.Drawing.Font("Arial", 26)  
$label2.Font = New-Object System.Drawing.Font("Arial", 26)  
$inputBox1.Font = New-Object System.Drawing.Font("Arial", 26)  
$inputBox2.Font = New-Object System.Drawing.Font("Arial", 26)  

$form.Controls.Add($label1)
$form.Controls.Add($inputBox1)
$form.Controls.Add($label2)
$form.Controls.Add($inputBox2)

$okButton = New-Object Windows.Forms.Button
$okButton.Location = New-Object Drawing.Point @(90, 540)
$okButton.Size = New-Object Drawing.Size @(300, 100)
$okButton.Text = "OK"
$okButton.DialogResult = [Windows.Forms.DialogResult]::OK

$cancelButton = New-Object Windows.Forms.Button
$cancelButton.Location = New-Object Drawing.Point @(490, 540)
$cancelButton.Size = New-Object Drawing.Size @(300, 100)
$cancelButton.Text = "Cancel"
$cancelButton.DialogResult = [Windows.Forms.DialogResult]::Cancel

$form.AcceptButton = $okButton
$form.CancelButton = $cancelButton

$form.Controls.Add($okButton)
$form.Controls.Add($cancelButton)

$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [Windows.Forms.DialogResult]::OK) {
    $userInput1 = $inputBox1.Text
    $userInput2 = $inputBox2.Text
} else {
    Write-Host "User canceled the input."
}


$outputPath = $userInput1
$endDate = $userInput2


# Connect to the Power BI service account
Connect-PowerBIServiceAccount

# Get the current date and time for reference
$retrieveDate = Get-Date 

# Construct the path for the CSV file
$activityLogsPath = Join-Path -Path $outputPath -ChildPath "ActivityLogs.csv"

# Convert start and end date strings to DateTime objects
$startDate = (Get-Date $endDate).AddDays(-30)
$endDate = (Get-Date $endDate).AddDays(1)
if ($startDate -lt $retrieveDate.AddDays(-30)) {
    $startDate = $retrieveDate.AddDays(-30)
}

# Initialize the loop with the start date
$currentDate = Get-Date $startDate
$activityLog = @()

# Loop through each day in the specified date range
while ($currentDate -le $endDate) {
    # Format the current date to create the start and end datetime strings
    $dateStr = $currentDate.ToString("yyyy-MM-dd")
    Write-Host $dateStr
    $startDt = $dateStr + 'T00:00:00.000'
    $endDt = $dateStr + 'T23:59:59.999'

    # Define parameters for retrieving Power BI activity logs
    $activityLogsParams = @{
        StartDateTime = $startDt
        EndDateTime   = $endDt
    }

    # Retrieve and convert Power BI activity logs from JSON
    $activityLogs = Get-PowerBIActivityEvent @activityLogsParams | ConvertFrom-Json

    # Select relevant properties and add a 'RetrieveDate' property
    $activityLogSchema = $activityLogs | Select-Object Id, RecordType, CreationTime, Operation, OrganizationId, UserType, UserKey, Workload, `
        UserId, ClientIP, UserAgent, Activity, ItemName, WorkspaceName, DatasetName, ReportName, `
        WorkspaceId, CapacityId, CapacityName, AppName, ObjectId, DatasetId, ReportId, IsSuccess, `
        ReportType, RequestId, ActivityId, AppReportId, DistributionMethod, ConsumptionMethod, `
        @{Name="RetrieveDate"; Expression={$retrieveDate.ToString("yyyy-MM-ddThh:mm:ss")}}

    # Add the activity log data to the array
    $activityLog += $activityLogSchema

    # Move to the next day
    $currentDate = $currentDate.AddDays(1)
}

$activityLog | Export-Csv $ActivityLogsPath -NoTypeInformation
