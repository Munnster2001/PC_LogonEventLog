# Load the necessary assemblies for creating a GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Check if the script is running with administrator privileges
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    [System.Windows.Forms.MessageBox]::Show("This script must be run with Administrator privileges to access the Security event log.", "Permission Denied", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Advanced Logon History Viewer'
$form.Size = New-Object System.Drawing.Size(1200, 550)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false

# Create controls for user input
$computerLabel = New-Object System.Windows.Forms.Label
$computerLabel.Text = 'Remote Computer Name:'
$computerLabel.Location = New-Object System.Drawing.Point(10, 20)
$computerLabel.Size = New-Object System.Drawing.Size(150, 20)

$computerNameTextBox = New-Object System.Windows.Forms.TextBox
$computerNameTextBox.Location = New-Object System.Drawing.Point(160, 20)
$computerNameTextBox.Size = New-Object System.Drawing.Size(220, 20)
$computerNameTextBox.Text = 'localhost' # Defaults to localhost

$daysLabel = New-Object System.Windows.Forms.Label
$daysLabel.Text = 'Days (e.g., 20):'
$daysLabel.Location = New-Object System.Drawing.Point(10, 50)
$daysLabel.Size = New-Object System.Drawing.Size(150, 20)

$daysTextBox = New-Object System.Windows.Forms.TextBox
$daysTextBox.Location = New-Object System.Drawing.Point(160, 50)
$daysTextBox.Size = New-Object System.Drawing.Size(220, 20)
$daysTextBox.Text = '3'

# Checkbox to exclude system accounts
$excludeSystemCheckBox = New-Object System.Windows.Forms.CheckBox
$excludeSystemCheckBox.Text = 'Exclude System Accounts'
$excludeSystemCheckBox.Location = New-Object System.Drawing.Point(400, 20)
$excludeSystemCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$excludeSystemCheckBox.Checked = $true

# Create a Log Box to provide detailed feedback
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Text = 'Ready.'
$logTextBox.Location = New-Object System.Drawing.Point(10, 420)
$logTextBox.Size = New-Object System.Drawing.Size(1160, 90)
$logTextBox.Multiline = $true
$logTextBox.ReadOnly = $true
$logTextBox.ScrollBars = 'Vertical'

# Create a progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 400)
$progressBar.Size = New-Object System.Drawing.Size(1160, 15)
$progressBar.Style = 'Marquee'
$progressBar.Visible = $false

# Create the "Get History" button
$getHistoryButton = New-Object System.Windows.Forms.Button
$getHistoryButton.Text = 'Get Advanced History'
$getHistoryButton.Location = New-Object System.Drawing.Point(160, 80)
$getHistoryButton.Size = New-Object System.Drawing.Size(150, 30)

# Create the "Export to CSV" button
$exportCsvButton = New-Object System.Windows.Forms.Button
$exportCsvButton.Text = 'Export to CSV'
$exportCsvButton.Location = New-Object System.Drawing.Point(320, 80)
$exportCsvButton.Size = New-Object System.Drawing.Size(150, 30)

# Create the "Clear Form" button
$clearFormButton = New-Object System.Windows.Forms.Button
$clearFormButton.Text = 'Clear Form'
$clearFormButton.Location = New-Object System.Drawing.Point(480, 80)
$clearFormButton.Size = New-Object System.Drawing.Size(150, 30)

# Create a DataGridView to display the results
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, 120)
$dataGridView.Size = New-Object System.Drawing.Size(760, 260)
$dataGridView.AutoSizeColumnsMode = 'Fill'
$dataGridView.ReadOnly = $true
$dataGridView.SelectionMode = 'FullRowSelect'
$dataGridView.AutoGenerateColumns = $false

# Create a text box to display event details
$eventDetailsTextBox = New-Object System.Windows.Forms.TextBox
$eventDetailsTextBox.Location = New-Object System.Drawing.Point(790, 120)
$eventDetailsTextBox.Size = New-Object System.Drawing.Size(380, 260)
$eventDetailsTextBox.Multiline = $true
$eventDetailsTextBox.ReadOnly = $true
$eventDetailsTextBox.ScrollBars = 'Vertical'

# Define the action for when a row in the DataGridView is selected
$dataGridView.Add_SelectionChanged({
    if ($dataGridView.SelectedRows.Count -gt 0) {
        $selectedRow = $dataGridView.SelectedRows[0].DataBoundItem
        if ($selectedRow) {
            $details = @"
Time: $($selectedRow.Time)
Event Type: $($selectedRow.'Event Type')
User: $($selectedRow.User)
Logon Type: $($selectedRow.'Logon Type')
Logon Title: $($selectedRow.'Logon Title')
Source Address: $($selectedRow.'Source Address')

--- Subject Details ---
Subject Security ID: $($selectedRow.'Subject Security ID')
Subject Account Name: $($selectedRow.'Subject Account Name')
Subject Account Domain: $($selectedRow.'Subject Account Domain')
Subject Logon ID: $($selectedRow.'Subject Logon ID')

--- Authentication Details ---
Logon Process: $($selectedRow.'Logon Process')
Authentication Package: $($selectedRow.'Authentication Package')
Transited Services: $($selectedRow.'Transited Services')
Package Name: $($selectedRow.'Package Name')
Key Length: $($selectedRow.'Key Length')
"@
            $eventDetailsTextBox.Text = $details
        }
    }
})

# Define the action for the "Get History" button
$getHistoryButton.Add_Click({
    # Disable buttons and start progress indicators
    $getHistoryButton.Enabled = $false
    $exportCsvButton.Enabled = $false
    $clearFormButton.Enabled = $false
    
    $logTextBox.Clear()
    $logTextBox.AppendText("`n[$(Get-Date)] Starting event log query...")
    $progressBar.Visible = $true
    [System.Windows.Forms.Application]::DoEvents() # Force UI refresh

    $Computer = $computerNameTextBox.Text
    $Days = [int]$daysTextBox.Text
    $excludeSystem = $excludeSystemCheckBox.Checked

    $dataGridView.DataSource = $null
    $dataGridView.Columns.Clear() # Clear existing columns
    $eventDetailsTextBox.Text = ''

    if ([string]::IsNullOrWhiteSpace($Computer)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a computer name.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $logTextBox.AppendText("`n[$(Get-Date)] Error: No computer name specified.")
        $progressBar.Visible = $false
        $getHistoryButton.Enabled = $true
        $exportCsvButton.Enabled = $true
        $clearFormButton.Enabled = $true
        return
    }

    try {
        $logTextBox.AppendText("`n[$(Get-Date)] Querying events from '$Computer' for the last $Days days.")
        $Result = New-Object System.Collections.ArrayList
        
        $LogonTypeMap = @{
            '2'  = "Interactive (Local)"
            '3'  = "Network (Remote Share)"
            '4'  = "Batch (Scheduled Task)"
            '5'  = "Service (Service Logon)"
            '7'  = "Unlock (Workstation Unlock)"
            '8'  = "Network Cleartext (Legacy)"
            '10' = "RemoteInteractive (RDP/TS)"
            '11' = "CachedInteractive (Local)"
        }
        
        $logonEventIDs = @(4624, 4625, 4634, 4647, 4800, 4801, 4672)
        $startTime = (Get-Date).AddDays(-$Days)

        try {
            $logTextBox.AppendText("`n[$(Get-Date)] Running Get-WinEvent with filter...")
            
            $filterHashtable = @{
                LogName = 'Security'
                StartTime = $startTime
                ID = $logonEventIDs
            }

            if ($Computer -ne 'localhost') {
                $filterHashtable.Add('ComputerName', $Computer)
            }
            
            # Run the query and check for data
            $Events = Get-WinEvent -FilterHashtable $filterHashtable -ErrorAction Stop
            
            if (-not $Events) {
                $logTextBox.AppendText("`n[$(Get-Date)] No events found for the specified criteria.")
                [System.Windows.Forms.MessageBox]::Show("No events were found for the specified criteria. Try increasing the number of days or check your event log.", "No Events Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $progressBar.Visible = $false
                $getHistoryButton.Enabled = $true
                $exportCsvButton.Enabled = $true
                $clearFormButton.Enabled = $true
                return
            }
        }
        catch {
            $logTextBox.AppendText("`n[$(Get-Date)] An error occurred while retrieving event logs. The error was: $_")
            [System.Windows.Forms.MessageBox]::Show("An error occurred while retrieving event logs from '$Computer'. The error was: $_`n`nCommon issues include: Incorrect computer name, lack of administrative permissions, or Windows Firewall blocking access on the remote machine.", "Event Log Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $progressBar.Visible = $false
            $getHistoryButton.Enabled = $true
            $exportCsvButton.Enabled = $true
            $clearFormButton.Enabled = $true
            return
        }

        $logTextBox.AppendText("`n[$(Get-Date)] Found $($Events.Count) events. Processing...")
        [System.Windows.Forms.Application]::DoEvents() # Force UI refresh
        
        $eventCount = 0
        $updateFrequency = 100 # Refresh the UI every 100 events

        ForEach ($Event in $Events) {
            $EventID = $Event.Id
            
            switch ($EventID) {
                4624 { $ET = "Logon Successful" }
                4625 { $ET = "Logon Failed" }
                4634 { $ET = "Logoff" }
                4647 { $ET = "User-initiated Logoff" }
                4800 { $ET = "Workstation Locked" }
                4801 { $ET = "Workstation Unlocked" }
                4672 { $ET = "Special Logon" }
                default { $ET = "Unknown Event" }
            }

            [xml]$eventXML = $Event.ToXml()
            $userNode = $eventXML.Event.EventData.Data | Where-Object { $_.Name -eq 'TargetUserName' }
            
            if (-not $userNode) {
                $userNode = $eventXML.Event.EventData.Data | Where-Object { $_.Name -eq 'SubjectUserName' }
            }
            
            $User = if ($userNode) {
                if ($userNode.'#text' -eq 'SYSTEM' -or $userNode.'#text' -like '*$') {
                    'SYSTEM ACCOUNT'
                } else {
                    $userNode.'#text'
                }
            } else {
                'N/A'
            }

            if ($excludeSystem -and $User -eq 'SYSTEM ACCOUNT') { continue }

            $LogonType = 'N/A'
            $LogonTitle = 'N/A'
            $SourceAddress = 'N/A'
            $LogonProcess = 'N/A'
            $AuthenticationPackage = 'N/A'
            $TransitedServices = 'N/A'
            $PackageName = 'N/A'
            $KeyLength = 'N/A'

            # Parse Logon Type and Source Address for applicable events
            if ($Event.Message -match 'Logon Type:\s*(\d+)') {
                $LogonType = $Matches[1]
                if ($LogonTypeMap.ContainsKey($LogonType)) {
                    $LogonTitle = $LogonTypeMap[$LogonType]
                }
            }
            if ($Event.Message -match 'Source Network Address:\s*([\d\.]+|[-])') {
                $SourceAddress = $Matches[1]
            }

            if ($EventID -eq 4624) {                
                $LogonProcess = ($eventXML.Event.EventData.Data | Where-Object { $_.Name -eq 'LogonProcessName' }).'#text'
                $AuthenticationPackage = ($eventXML.Event.EventData.Data | Where-Object { $_.Name -eq 'AuthenticationPackageName' }).'#text'
                $TransitedServices = ($eventXML.Event.EventData.Data | Where-Object { $_.Name -eq 'TransitedServices' }).'#text'
                $PackageName = ($eventXML.Event.EventData.Data | Where-Object { $_.Name -eq 'PackageName' }).'#text'
                $KeyLength = ($eventXML.Event.EventData.Data | Where-Object { $_.Name -eq 'KeyLength' }).'#text'
            }
            
            $subjectSecurityId = ($eventXML.Event.EventData.Data | Where-Object { $_.Name -eq 'SubjectUserSid' }).'#text'
            $subjectAccountName = ($eventXML.Event.EventData.Data | Where-Object { $_.Name -eq 'SubjectUserName' }).'#text'
            $subjectAccountDomain = ($eventXML.Event.EventData.Data | Where-Object { $_.Name -eq 'SubjectDomainName' }).'#text'
            $subjectLogonId = ($eventXML.Event.EventData.Data | Where-Object { $_.Name -eq 'SubjectLogonId' }).'#text'

            $Result.Add([PSCustomObject]@{
                Time = $Event.TimeCreated
                'Event Type' = $ET
                User = $User
                'Logon Type' = $LogonType
                'Logon Title' = $LogonTitle
                'Source Address' = $SourceAddress
                'Subject Security ID' = $subjectSecurityId
                'Subject Account Name' = $subjectAccountName
                'Subject Account Domain' = $subjectAccountDomain
                'Subject Logon ID' = $subjectLogonId
                'Logon Process' = $LogonProcess
                'Authentication Package' = $AuthenticationPackage
                'Transited Services' = $TransitedServices
                'Package Name' = $PackageName
                'Key Length' = $KeyLength
            }) | Out-Null
            
            $eventCount++
            if ($eventCount % $updateFrequency -eq 0) {
                [System.Windows.Forms.Application]::DoEvents() # Keep UI responsive during processing
            }
        }
        
        $logTextBox.AppendText("`n[$(Get-Date)] Finished processing events. Displaying data.")
        $logTextBox.AppendText("`n[$(Get-Date)] Result count: $($Result.Count)")
        
        # --- Updated logic to ensure DataGridView displays data ---
        $dataGridView.DataSource = $null # Clear the DataSource first
        $dataGridView.Columns.Clear() # Clear existing columns
        if ($Result.Count -gt 0) {
            # Convert to DataTable for better DataGridView compatibility, including all properties
            $dataTable = New-Object System.Data.DataTable
            $dataTable.Columns.Add("Time", [System.DateTime])
            $dataTable.Columns.Add("Event Type", [string])
            $dataTable.Columns.Add("User", [string])
            $dataTable.Columns.Add("Logon Type", [string])
            $dataTable.Columns.Add("Logon Title", [string])
            $dataTable.Columns.Add("Source Address", [string])
            $dataTable.Columns.Add("Subject Security ID", [string])
            $dataTable.Columns.Add("Subject Account Name", [string])
            $dataTable.Columns.Add("Subject Account Domain", [string])
            $dataTable.Columns.Add("Subject Logon ID", [string])
            $dataTable.Columns.Add("Logon Process", [string])
            $dataTable.Columns.Add("Authentication Package", [string])
            $dataTable.Columns.Add("Transited Services", [string])
            $dataTable.Columns.Add("Package Name", [string])
            $dataTable.Columns.Add("Key Length", [string])

            foreach ($item in $Result) {
                $row = $dataTable.NewRow()
                $row["Time"] = $item.Time
                $row["Event Type"] = $item.'Event Type'
                $row["User"] = $item.User
                $row["Logon Type"] = $item.'Logon Type'
                $row["Logon Title"] = $item.'Logon Title'
                $row["Source Address"] = $item.'Source Address'
                $row["Subject Security ID"] = $item.'Subject Security ID'
                $row["Subject Account Name"] = $item.'Subject Account Name'
                $row["Subject Account Domain"] = $item.'Subject Account Domain'
                $row["Subject Logon ID"] = $item.'Subject Logon ID'
                $row["Logon Process"] = $item.'Logon Process'
                $row["Authentication Package"] = $item.'Authentication Package'
                $row["Transited Services"] = $item.'Transited Services'
                $row["Package Name"] = $item.'Package Name'
                $row["Key Length"] = $item.'Key Length'
                $dataTable.Rows.Add($row)
            }

            # Add displayed columns
            $colTime = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $colTime.DataPropertyName = "Time"
            $colTime.HeaderText = "Time"
            $dataGridView.Columns.Add($colTime)

            $colEventType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $colEventType.DataPropertyName = "Event Type"
            $colEventType.HeaderText = "Event Type"
            $dataGridView.Columns.Add($colEventType)

            $colUser = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $colUser.DataPropertyName = "User"
            $colUser.HeaderText = "User"
            $dataGridView.Columns.Add($colUser)

            $colLogonType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $colLogonType.DataPropertyName = "Logon Type"
            $colLogonType.HeaderText = "Logon Type"
            $dataGridView.Columns.Add($colLogonType)

            $colLogonTitle = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $colLogonTitle.DataPropertyName = "Logon Title"
            $colLogonTitle.HeaderText = "Logon Title"
            $dataGridView.Columns.Add($colLogonTitle)

            $colSourceAddress = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $colSourceAddress.DataPropertyName = "Source Address"
            $colSourceAddress.HeaderText = "Source Address"
            $dataGridView.Columns.Add($colSourceAddress)

            $dataGridView.DataSource = $dataTable
            $dataGridView.Sort($dataGridView.Columns[0], [System.ComponentModel.ListSortDirection]::Descending)
            $dataGridView.Refresh() # Force UI refresh
            $logTextBox.AppendText("`n[$(Get-Date)] DataGridView bound with $($dataTable.Rows.Count) rows.")
        } else {
            $logTextBox.AppendText("`n[$(Get-Date)] No data to display in DataGridView.")
        }
        # --- End of updated logic ---
        
    } catch {
        $logTextBox.AppendText("`n[$(Get-Date)] An unknown error occurred. The error was: $_")
        [System.Windows.Forms.MessageBox]::Show("An unknown problem occurred. The error was: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $progressBar.Visible = $false
        $getHistoryButton.Enabled = $true
        $exportCsvButton.Enabled = $true
        $clearFormButton.Enabled = $true
    }
})

# Define the action for the "Export to CSV" button
$exportCsvButton.Add_Click({
    if ($dataGridView.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("There is no data to export.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    $saveFileDialog.FileName = "LogonHistory-$(Get-Date -Format yyyy-MM-dd).csv"

    if ($saveFileDialog.ShowDialog() -eq 'OK') {
        try {
            $dataGridView.DataSource | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show("Data successfully exported to `"$($saveFileDialog.FileName)`"", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("An error occurred during export: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Define the action for the "Clear Form" button
$clearFormButton.Add_Click({
    $computerNameTextBox.Text = 'localhost'
    $daysTextBox.Text = '3'
    $excludeSystemCheckBox.Checked = $true
    $dataGridView.DataSource = $null
    $dataGridView.Columns.Clear()
    $eventDetailsTextBox.Text = ''
    $logTextBox.Clear()
    $logTextBox.AppendText('Ready.')
})

# Add controls to the form
$form.Controls.Add($computerLabel)
$form.Controls.Add($computerNameTextBox)
$form.Controls.Add($daysLabel)
$form.Controls.Add($daysTextBox)
$form.Controls.Add($excludeSystemCheckBox)
$form.Controls.Add($getHistoryButton)
$form.Controls.Add($exportCsvButton)
$form.Controls.Add($clearFormButton)
$form.Controls.Add($logTextBox)
$form.Controls.Add($progressBar)
$form.Controls.Add($dataGridView)
$form.Controls.Add($eventDetailsTextBox)

# Show the form
$form.ShowDialog()
