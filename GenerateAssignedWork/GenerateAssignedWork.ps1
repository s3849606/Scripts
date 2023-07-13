<# SCRIPT IS NOT TO BE DISTRIBUTED OR USED OUTSIDE OF SERVICE DESK TECH MENTORS

This script is used only by Service Deks Tech Mentors to delegate work to Service Desk Agents
Script is to be used within Powershell ISE with no administrator privileges

Version = v1.0.7
Author: James Zhou (UWBYR)

The script takes two sets of data (Number of Agents and Work to be delegated).
It will evenly distribute the work provided and save a tempoary copy of an excle sheet with the data formatted on each row for nubmer of agents.
When the GUI is closed, the script will clean up and delete any files created during the runtime of the script.
#>

# Import packages
Add-Type -AssemblyName System.Windows.Forms

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Generate Assigned Task"
$form.Size = New-Object System.Drawing.Size(420, 505)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$fileSaved = $false

# Create a label for the number of agents
$numOfAgentsLabel = New-Object System.Windows.Forms.Label
$numOfAgentsLabel.Location = New-Object System.Drawing.Point(83, 20)
$numOfAgentsLabel.Size = New-Object System.Drawing.Size(100, 20)
$numOfAgentsLabel.Text = "Number of Agents:"
$form.Controls.Add($numOfAgentsLabel)

# Create a textbox for entering the number of agents
$numOfAgentsInputField = New-Object System.Windows.Forms.TextBox
$numOfAgentsInputField.Location = New-Object System.Drawing.Point(80, 40)
$numOfAgentsInputField.Size = New-Object System.Drawing.Size(110, 20)
$numOfAgentsInputField.TextAlign = "Center"
$numOfAgentsInputField.MaxLength = 3
$numOfAgentsInputField.Text = "1"
$form.Controls.Add($numOfAgentsInputField)

# Create a decrease button
$decreaseNumOfAgentsLabel = New-Object System.Windows.Forms.Button
$decreaseNumOfAgentsLabel.Location = New-Object System.Drawing.Point(55, 40)
$decreaseNumOfAgentsLabel.Size = New-Object System.Drawing.Size(30, 20)
$decreaseNumOfAgentsLabel.Text = "v"
$decreaseNumOfAgentsLabel.Add_Click(
    {
        $numOfAgents = [int]$numOfAgentsInputField.Text
        if ($numOfAgents -gt 1) {
            $numOfAgents--
            $numOfAgentsInputField.Text = $numOfAgents.ToString()
        }
    }
)
$form.Controls.Add($decreaseNumOfAgentsLabel)

# Create an increase button
$increaseNumOfAgentsLabel = New-Object System.Windows.Forms.Button
$increaseNumOfAgentsLabel.Location = New-Object System.Drawing.Point(185, 40)
$increaseNumOfAgentsLabel.Size = New-Object System.Drawing.Size(30, 20)
$increaseNumOfAgentsLabel.Text = "^"
$increaseNumOfAgentsLabel.Add_Click(
    {
        $numOfAgents = [int]$numOfAgentsInputField.Text
        $numOfAgents++
        $numOfAgentsInputField.Text = $numOfAgents.ToString()
    }
)
$form.Controls.Add($increaseNumOfAgentsLabel)

# Create a checkbox for randomizing
$randomizeCheckBox = New-Object System.Windows.Forms.CheckBox
$randomizeCheckBox.Text = "Randomize Work"
$randomizeCheckBox.Location = New-Object System.Drawing.Point(250, 40)
$randomizeCheckBox.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($randomizeCheckBox)

# Create a label for the tasks input
$assignedWorkLabel = New-Object System.Windows.Forms.Label
$assignedWorkLabel.Location = New-Object System.Drawing.Point(20, 75)
$assignedWorkLabel.Size = New-Object System.Drawing.Size(360, 20)
$assignedWorkLabel.Text = "Paste all Assigned Work here:"
$assignedWorkLabel.TextAlign = "MiddleCenter"
$form.Controls.Add($assignedWorkLabel)

# Create a textbox for entering the tasks
$assignedWorkTextField = New-Object System.Windows.Forms.TextBox
$assignedWorkTextField.Location = New-Object System.Drawing.Point(20, 95)
$assignedWorkTextField.Size = New-Object System.Drawing.Size(360, 300)
$assignedWorkTextField.Multiline = $true
$assignedWorkTextField.ScrollBars = "None"
$form.Controls.Add($assignedWorkTextField)

# Calculate the maximum visible lines based on the size of the text field
$assignedWorkTextField.Add_TextChanged(
    {
        $lineHeight = $assignedWorkTextField.Font.GetHeight()
        $maxVisibleLines = [math]::Floor($assignedWorkTextField.Height / $lineHeight)
    
        # Set the scrollbars based on the text content
        if ($assignedWorkTextField.Lines.Length -gt $maxVisibleLines) {
            $assignedWorkTextField.ScrollBars = "Vertical"
        }
        else {
            $assignedWorkTextField.ScrollBars = "None"
        }
    }
)

# Create a button to generate an Excel sheet containing assigned work
$generateSheetButton = New-Object System.Windows.Forms.Button
$generateSheetButton.Location = New-Object System.Drawing.Point(100, 410)
$generateSheetButton.Size = New-Object System.Drawing.Size(100, 30)
$generateSheetButton.Text = "Generate Sheet"
$form.Controls.Add($generateSheetButton)

# Create a button to delete the generated Excel sheet
$deleteSheetButton = New-Object System.Windows.Forms.Button
$deleteSheetButton.Location = New-Object System.Drawing.Point(220, 410)
$deleteSheetButton.Size = New-Object System.Drawing.Size(100, 30)
$deleteSheetButton.Text = "Delete Sheet"
$form.Controls.Add($deleteSheetButton)

# Handle CNTRL+A to select all text in 'Assigned Work' text field
$assignedWorkTextField.Add_KeyDown(
    {
        param($textField, $e)
        # Check if Ctrl+A is pressed
        if (($e.Modifiers -band [System.Windows.Forms.Keys]::Control) -and ($e.KeyCode -eq [System.Windows.Forms.Keys]::A)) {
            $textField.SelectAll()
            $e.Handled = $true
        }
    }
)

# Handle CNTRL+A to select all text in 'Number of Agents' text field
$numOfAgentsInputField.Add_KeyDown(
    {
        param($textField, $e)
        # Check if Ctrl+A is pressed
        if (($e.Modifiers -band [System.Windows.Forms.Keys]::Control) -and ($e.KeyCode -eq [System.Windows.Forms.Keys]::A)) {
            $textField.SelectAll()
            $e.Handled = $true
        }
    }
)

# Event handler to generate excel sheet button
$generateSheetButton.Add_Click(
    {
        # Get the assignedWork from the textbox
        $assignedWork = $assignedWorkTextField.Text -split "`r?`n" | Where-Object { $_.Trim() -ne "" }

        # Get the number of agents
        $numOfAgents = $numOfAgentsInputField.Text

        # Validate inputs
        if ($assignedWork.Length -eq 0 -or $numOfAgents -le 1) {
            [System.Windows.Forms.MessageBox]::Show("Invalid input!", "Error", "OK", "Error")
            return
        }

        # Randomize the tasks if the checkbox is checked
        if ($randomizeCheckBox.Checked) {
            $assignedWork = Get-Random -InputObject $assignedWork -Count $assignedWork.Length
        }

        # Create a new Excel workbook
        $excel = New-Object -ComObject Excel.Application
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)

        # Show the Excel application
        $excel.Visible = $true

        # Wait for the Excel application to become visible
        while (-not $excel.Visible) {
            Start-Sleep -Milliseconds 100
        }

        # Populate the tasks in the Excel sheet
        for ($i = 0; $i -lt $assignedWork.Length; $i++) {
            $row = $i % $numOfAgents + 1
            $column = [math]::Floor($i / $numOfAgents) + 1
            $sheet.Cells.Item($row, $column) = $assignedWork[$i]
        }

        # Save the workbook
        $workbook.SaveAs("$env:TEMP\DelegatedAssignedWork.xlsx")

        # Set state of fileSaved value to fasle
        $script:fileSaved = $true
    }
)

# Event handler to delete excel sheet button
$deleteSheetButton.Add_Click(
    {
        # Clean up Excel objects
        if ($excel -ne $null) {
            $excel.Quit()
            if ($sheet -ne $null) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
            }
            if ($workbook -ne $null) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
            }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }

        if ($fileSaved -eq $true) {
            # Delete the Excel file
            $excelFilePath = "$env:TEMP\DelegatedAssignedWork.xlsx"
            if (Test-Path $excelFilePath) {
                Remove-Item $excelFilePath -Force
            }
        }
        $script:fileSaved = $false
    }
)

# Event handler for form closing
$form.Add_FormClosed(
    {
        # Clean up Excel objects
        if ($excel -ne $null) {
            $excel.Quit()
            if ($sheet -ne $null) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
            }
            if ($workbook -ne $null) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
            }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }

        if ($fileSaved -eq $true) {
            # Delete the Excel file
            $excelFilePath = "$env:TEMP\DelegatedAssignedWork.xlsx"
            if (Test-Path $excelFilePath) {
                Remove-Item $excelFilePath -Force
            }
        }
    }
)

# Show the form
$form.Add_Shown(
    {
        $form.Activate() 
    }
)
[void]$form.ShowDialog()