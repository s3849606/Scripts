# SCRIPT IS NOT TO BE DISTRIBUTED OR USED OUTSIDE OF SERVICE DESK TECH MENTORS

# This script is used only by Service Deks Tech Mentors to extract data from text
# Script is to be used within Powershell ISE with no administrator privileges

# v1.0.3

# FOR USER IDs
# This is used for fallbacks to easily email a list of users from a global group, format will follow firstName lastName(USERID);firstName lastName(USERID);
# This script will return all User DIs on seperate lines

# For INC/REQ/RITM/SCTASKS
# The scripot will return a list of all matches of INC/REQ/RITM/SCTASK numbers on seperate lines


# Import packages
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Creation of form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Extract User IDs/INC/REQ/RITM/SCTASK numbers"
$form.Width = 500
$form.Height = 490
$form.StartPosition = "CenterScreen"

# Create label for inputTextField
$inputTextFieldLabel = New-Object System.Windows.Forms.Label
$inputTextFieldLabel.Location = New-Object System.Drawing.Point(10, 20)
$inputTextFieldLabel.Size = New-Object System.Drawing.Size(480, 30)
$inputTextFieldLabel.Text = "Enter the raw data to extract from:"
$form.Controls.Add($inputTextFieldLabel)

# Create text field for user input
$inputTextField = New-Object System.Windows.Forms.TextBox
$inputTextField.Location = New-Object System.Drawing.Point(10, 50)
$inputTextField.Size = New-Object System.Drawing.Size(460, 70)
$inputTextField.Multiline = $true
$inputTextField.ScrollBars = "Vertical"
$form.Controls.Add($inputTextField)

# Create label for output options
$outputOptionsLabel = New-Object System.Windows.Forms.Label
$outputOptionsLabel.Location = New-Object System.Drawing.Point(162, 130)
$outputOptionsLabel.Size = New-Object System.Drawing.Size(176, 25)
$outputOptionsLabel.Text = "Select what you want to extract to:"
$form.Controls.Add($outputOptionsLabel)

# Create check box for User ID
$userIdCheckBox = New-Object System.Windows.Forms.CheckBox
$userIdCheckBox.Location = New-Object System.Drawing.Point(30, 150)
$userIdCheckBox.Size = New-Object System.Drawing.Size(65, 30)
$userIdCheckBox.Text = "User ID"
$form.Controls.Add($userIdCheckBox)

# Create check box for Name with user ID
$nameUserIDCheckBox = New-Object System.Windows.Forms.CheckBox
$nameUserIDCheckBox.Location = New-Object System.Drawing.Point(98, 150)
$nameUserIDCheckBox.Size = New-Object System.Drawing.Size(96, 30)
$nameUserIDCheckBox.Text = "Name(UserID)"
$form.Controls.Add($nameUserIDCheckBox)

# Create check box for INC
$incCheckBox = New-Object System.Windows.Forms.CheckBox
$incCheckBox.Location = New-Object System.Drawing.Point(203, 150)
$incCheckBox.Size = New-Object System.Drawing.Size(43, 30)
$incCheckBox.Text = "INC"
$form.Controls.Add($incCheckBox)

# Create check box for REQ
$reqCheckBox = New-Object System.Windows.Forms.CheckBox
$reqCheckBox.Location = New-Object System.Drawing.Point(253, 150)
$reqCheckBox.Size = New-Object System.Drawing.Size(48, 30)
$reqCheckBox.Text = "REQ"
$form.Controls.Add($reqCheckBox)

# Create check box for RITM
$ritmCheckBox = New-Object System.Windows.Forms.CheckBox
$ritmCheckBox.Location = New-Object System.Drawing.Point(324, 150)
$ritmCheckBox.Size = New-Object System.Drawing.Size(52, 30)
$ritmCheckBox.Text = "RITM"
$form.Controls.Add($ritmCheckBox)

# Create check box for SCTASK
$sctaskCheckBox = New-Object System.Windows.Forms.CheckBox
$sctaskCheckBox.Location = New-Object System.Drawing.Point(383, 150)
$sctaskCheckBox.Size = New-Object System.Drawing.Size(68, 30)
$sctaskCheckBox.Text = "SCTASK"
$form.Controls.Add($sctaskCheckBox)

# Create extract button
$extractButton = New-Object System.Windows.Forms.Button
$extractButton.Location = New-Object System.Drawing.Point(140, 195)
$extractButton.Size = New-Object System.Drawing.Size(100, 30)
$extractButton.Text = "Extract"
$form.Controls.Add($extractButton)

# Create clear button
$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Location = New-Object System.Drawing.Point(260, 195)
$clearButton.Size = New-Object System.Drawing.Size(100, 30)
$clearButton.Text = "Clear"
$form.Controls.Add($clearButton)

# Create label for output text field
$outputTextFieldLabel = New-Object System.Windows.Forms.Label
$outputTextFieldLabel.Location = New-Object System.Drawing.Point(10, 250)
$outputTextFieldLabel.Size = New-Object System.Drawing.Size(480, 30)
$outputTextFieldLabel.Text = "Extracted data (CNTRL+A/Select all and copy):"
$form.Controls.Add($outputTextFieldLabel)

# Create text field for output containing extracted data
$outputTextField = New-Object System.Windows.Forms.TextBox
$outputTextField.Location = New-Object System.Drawing.Point(10, 280)
$outputTextField.Size = New-Object System.Drawing.Size(460, 150)
$outputTextField.Multiline = $true
$outputTextField.ScrollBars = "Vertical"
$form.Controls.Add($outputTextField)

# Add a keydown event to the inputTextField
$inputTextField.Add_KeyDown(
    {
        # Check if the key pressed is "A" and the control key is also pressed
        if ($_.KeyCode -eq "A" -and $_.Control) {
            # If true, select all text in the inputTextField
            $inputTextField.SelectAll()
            # Set the event as handled to prevent the default behavior of the key press
            $_.Handled = $true
        }
    }
)

# Add a keydown event to the outputTextField
$outputTextField.Add_KeyDown(
    {
        # Check if the key pressed is "A" and the control key is also pressed
        if ($_.KeyCode -eq "A" -and $_.Control) {
            # If true, select all text in the outputTextField
            $outputTextField.SelectAll()
            # Set the event as handled to prevent the default behavior of the key press
            $_.Handled = $true
        }
    }
)

$userIdCheckBox.Add_Click(
    {
        if ($userIdCheckBox.Checked) {
            $nameUserIDCheckBox.Enabled = $false
            $nameUserIDCheckBox.Checked = $false
            $nameUserIDCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)

            $incCheckBox.Enabled = $false
            $incCheckBox.Checked = $false
            $incCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)

            $reqCheckBox.Enabled = $false
            $reqCheckBox.Checked = $false
            $reqCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)

            $ritmCheckBox.Enabled = $false
            $ritmCheckBox.Checked = $false
            $ritmCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)

            $sctaskCheckBox.Enabled = $false
            $sctaskCheckBox.Checked = $false
            $sctaskCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)
        }
        else {
            $nameUserIDCheckBox.Enabled = $true
            $nameUserIDCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
            
            $incCheckBox.Enabled = $true
            $incCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
        
            $reqCheckBox.Enabled = $true
            $reqCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
        
            $ritmCheckBox.Enabled = $true
            $ritmCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
        
            $sctaskCheckBox.Enabled = $true
            $sctaskCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
        }
    }
)

$nameUserIDCheckBox.Add_Click(
    {
        if ($nameUserIDCheckBox.Checked) {
            $userIdCheckBox.Enabled = $false
            $userIdCheckBox.Checked = $false
            $userIdCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)

            $incCheckBox.Enabled = $false
            $incCheckBox.Checked = $false
            $incCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)

            $reqCheckBox.Enabled = $false
            $reqCheckBox.Checked = $false
            $reqCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)

            $ritmCheckBox.Enabled = $false
            $ritmCheckBox.Checked = $false
            $ritmCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)

            $sctaskCheckBox.Enabled = $false
            $sctaskCheckBox.Checked = $false
            $sctaskCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)
        }
        else {
            $userIdCheckBox.Enabled = $true
            $userIdCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
            
            $incCheckBox.Enabled = $true
            $incCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
        
            $reqCheckBox.Enabled = $true
            $reqCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
        
            $ritmCheckBox.Enabled = $true
            $ritmCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
        
            $sctaskCheckBox.Enabled = $true
            $sctaskCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
        }
    }
)

$incCheckBox.Add_Click(
    {
        if ($incCheckBox.Checked -or $reqCheckBox.Checked -or $ritmCheckBox.Checked -or $sctaskCheckBox.Checked) {
            $nameUserIDCheckBox.Enabled = $false
            $nameUserIDCheckBox.Checked = $false
            $nameUserIDCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)

            $userIdCheckBox.Enabled = $false
            $userIdCheckBox.Checked = $false
            $userIdCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)
        }
        else {
            $nameUserIDCheckBox.Enabled = $true
            $nameUserIDCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)

            $userIdCheckBox.Enabled = $true
            $userIdCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
        }
    }
)

$reqCheckBox.Add_Click(
    {
        if ($incCheckBox.Checked -or $reqCheckBox.Checked -or $ritmCheckBox.Checked -or $sctaskCheckBox.Checked) {
            $nameUserIDCheckBox.Enabled = $false
            $nameUserIDCheckBox.Checked = $false
            $nameUserIDCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)

            $userIdCheckBox.Enabled = $false
            $userIdCheckBox.Checked = $false
            $userIdCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)
        }
        else {
            $nameUserIDCheckBox.Enabled = $true
            $nameUserIDCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)

            $nameUserIDCheckBox.Enabled = $true
            $nameUserIDCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)

            $userIdCheckBox.Enabled = $true
            $userIdCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
        }
    }
)

$ritmCheckBox.Add_Click(
    {
        if ($incCheckBox.Checked -or $reqCheckBox.Checked -or $ritmCheckBox.Checked -or $sctaskCheckBox.Checked) {
            $nameUserIDCheckBox.Enabled = $false
            $nameUserIDCheckBox.Checked = $false
            $nameUserIDCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)

            $userIdCheckBox.Enabled = $false
            $userIdCheckBox.Checked = $false
            $userIdCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)
        }
        else {
            $nameUserIDCheckBox.Enabled = $true
            $nameUserIDCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)

            $userIdCheckBox.Enabled = $true
            $userIdCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
        }
    }
)

$sctaskCheckBox.Add_Click(
    {
        if ($incCheckBox.Checked -or $reqCheckBox.Checked -or $ritmCheckBox.Checked -or $sctaskCheckBox.Checked) {
            $nameUserIDCheckBox.Enabled = $false
            $nameUserIDCheckBox.Checked = $false
            $nameUserIDCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)

            $userIdCheckBox.Enabled = $false
            $userIdCheckBox.Checked = $false
            $userIdCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(196, 196, 196)
        }
        else {
            $nameUserIDCheckBox.Enabled = $true
            $nameUserIDCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)

            $userIdCheckBox.Enabled = $true
            $userIdCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
        }
    }
)

# This is the event handler for when the extract button is clicked
$extractButton.Add_Click(
    {
        # If no filters were selected, error message will appear
        if (-not ($userIdCheckBox.Checked -or $nameUserIDCheckBox.Checked -or $incCheckBox.Checked -or $reqCheckBox.Checked -or $ritmCheckBox.Checked -or $sctaskCheckBox.Checked)) {
            return  [System.Windows.Forms.MessageBox]::Show("No filter was selected, please select a filter and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        
        # Initialize an empty array list to store the extracted data
        $outputData = [System.Collections.ArrayList]@()
    
        # Initialize an empty string for the regular expression pattern
        $regexPattern = ""
    
        # If the user ID checkbox is checked, add the corresponding regex pattern
        if ($userIdCheckbox.Checked) {
            $regexPattern += "\(([^)]+)\)"
        }

        # If the Name(UserID) checkbox is checked, add the corresponding regex pattern
        if ($nameUserIDCheckBox.Checked) {
            $regexPattern += "^[A-Za-z]+ [A-Za-z]+\([A-Z]+\);$"
        }
    
        # If the INC checkbox is checked, add the INC pattern
        if ($incCheckBox.Checked) {
            $regexPattern += "(INC\d+)|"
        }
    
        # If the REQ checkbox is checked, add the REQ pattern
        if ($reqCheckBox.Checked) {
            $regexPattern += "(REQ\d+)|"
        }
    
        # If the RITM checkbox is checked, add the RITM pattern
        if ($ritmCheckBox.Checked) {
            $regexPattern += "(RITM\d+)|"
        }
    
        # If the SCTASK checkbox is checked, add the SCTASK pattern
        if ($sctaskCheckBox.Checked) {
            $regexPattern += "(SCTASK\d+)|"
        }
    
        # Remove the trailing "|" character if it exists
        if ($regexPattern.EndsWith("|")) {
            $regexPattern = $regexPattern.Substring(0, $regexPattern.Length - 1)
        }
    
        # Compile the regular expression pattern
        $regex = [regex]$regexPattern
    
        # Use the regular expression to find matches in the input text field
        $matches = $regex.Matches($inputTextField.Text.ToUpper())

        # If name(USERID) is checked
        if ($nameUserIDCheckBox.Checked) {
            $nameWithUserIDInput = $inputTextField.Text -replace ";", "`n"
            
            $outputData.Add($nameWithUserIDInput) | Out-Null
        }

        # Loop through the matches and add the captured group value(s) to the output array list
        else {
            foreach ($match in $matches) {
            
                #Remove () fromm match value
                $match = $match -replace '\(|\)', ''

                # User ID found within input data
                if ($match.Length -eq 5 -and ($match.StartsWith("U") -or $match.StartsWith("A"))) {
                    $outputData.Add($match) | Out-Null
                }
            
                # INC found within input data
                if ($match.StartsWith("INC")) {
                    $outputData.Add($match) | Out-Null
                }
            
                # REQ found within input data
                if ($match.StartsWith("REQ")) {
                    $outputData.Add($match) | Out-Null
                }
            
                # RITM found within input data
                if ($match.StartsWith("RITM")) {
                    $outputData.Add($match) | Out-Null
                }
            
                # SCTASK found within input data
                if ($match.StartsWith("SCTASK")) {
                    $outputData.Add($match) | Out-Null
                }
            }
        }
    
        # Set the output text field to display the extracted data as a string, or show an error message if no data was extracted
        if ($outputData.Count -gt 0) {
            $outputTextField.Text = $outputData -join "`r`n"
        }
        else {
            return [System.Windows.Forms.MessageBox]::Show("No data could be extracted, please check the data you entered and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
)

$clearButton.Add_Click(
    {
        $inputTextField.Clear()
        $outputTextField.Clear()
    }
)

$result = $form.ShowDialog()
