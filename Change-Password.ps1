n Bishop 
# last updated 6/18/2021
#
# Changes a users password in AD


# check if the machine has the correct module, if not install all req modules
$checkADModule    = (Get-Module ActiveDirectory -ListAvailable).Name
if ($checkADModule -eq $null)
{
	Write-Host "Important: Active Directory Powershell module must be installed to continue, attempting install "
	Install-PackageProvider -Name NuGet -Force
	$RSATname = Get-WindowsCapability -name "RSAT.ActiveDirectory*" -online | select name
	Add-WindowsCapability -Online -name $RSATname.name
}

Function changePassword {
	param (	[Parameter(Mandatory=$true, Position=0)] [string] $SAM, 
			[Parameter(Mandatory=$true, Position=1)] [string] $tokenP )

	$userExists = (Get-ADUser -filter 'SamAccountName -like $SAM')
	if ($userExists -eq $null)
	{
		[System.Windows.Forms.MessageBox]::Show("User not found. Should be first initial of first name and full last name. So for Jeff Williams it would be jwilliams" , "User not found!")
	}
	else
	{
		$lGroup   = Get-AdGroup "local admin"
		$group    = Get-AdGroup "administrators"
		$inLAdmin = (Get-ADUser -filter 'SamAccountName -like $SAM' -Properties memberof | Where-Object {$lGroup.DistinguishedName -in $_.memberof}).enabled
		$inAdmin  = (Get-ADUser -filter 'SamAccountName -like $SAM' -Properties memberof | Where-Object {$group.DistinguishedName -in $_.memberof}).enabled

		# don't change password if user is admin anything
		if ($inLAdmin -or $inAdmin)
		{
			[System.Windows.Forms.MessageBox]::Show("Do not use this to change password on admin accounts! " , "not safe!")
			return
		}
		# if the password is too short, inform the user
		if ($tokenP.length -lt 15)
		{
			[System.Windows.Forms.MessageBox]::Show("Password is too short, must be at least 16 characters long" , "")		
			return
		}
		# else change the password
		else
		{
			Set-ADAccountPassword -Identity $SAM -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $tokenP -Force)
			# if the user is active display sucess text, else warn the admin that the user is inactive (change password regardless)
			if ((Get-ADUser -filter 'SamAccountName -like $SAM').enabled)
			{
				[System.Windows.Forms.MessageBox]::Show("Password changed for $($SAM)" , "success!")
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show("Warning! This user is not active! However, password was changed for $($SAM)" , "warning!")
			}
		}
	}
}

# this function that creates the GUI
Function generateForm {

	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	$form = New-Object system.Windows.Forms.Form

	$form.ClientSize = '500,300'
	$form.BackColor  = '#000000'
	$form.ForeColor  = '#ffffff'
	$form.text       = 'Do not use this to change admin passwords'

	# setup the label that describes the first name text box purpose
	$SAM          = New-Object system.Windows.Forms.Label
	$SAM.text     = "Users login name example: abishop"
	$SAM.AutoSize = $true
	$SAM.location = New-Object System.Drawing.Point(10,70)

	# setup the SAM text box
	$SAMTextBox          = New-Object System.Windows.Forms.TextBox
	$SAMTextBox.Text     = $SAMname
	$SAMTextBox.Size     = New-Object System.Drawing.Size(260,20)
	$SAMTextBox.Location = New-Object System.Drawing.Point(20,90)
	$form.Controls.Add($SAMTextBox)

	# setup the label that describes the last name text box purpose
	$tokenP          = New-Object system.Windows.Forms.Label
	$tokenP.text     = "Password:"
	$tokenP.AutoSize = $true
	$tokenP.location = New-Object System.Drawing.Point(10,120)

	# setup the token text box
	$tokenPTextBox              = New-Object System.Windows.Forms.MaskedTextBox
	$tokenPTextBox.PasswordChar = "*"
	$tokenPTextBox.Size         = New-Object System.Drawing.Size(260,20)
	$tokenPTextBox.Location     = New-Object System.Drawing.Point(20,140)
	$form.Controls.Add($tokenPTextBox)

	# setup the create user button
	$changePasswordButton           = New-Object system.Windows.Forms.Button
	$changePasswordButton.text      = "Change Password"
	$changePasswordButton.width     = 150
	$changePasswordButton.height    = 30
	$changePasswordButton.location  = New-Object System.Drawing.Point(270,250)
	$changePasswordButton.Font      = 'Microsoft Sans Serif,10'
	$changePasswordButton.ForeColor = "#000000"
	$changePasswordButton.BackColor = "#dddddd"
	$changePasswordButton.Visible   = $true

	# Compile and display the form
	$form.controls.AddRange(@($changePasswordButton,$SAM,$SAMTextBox,$tokenP,$tokenPTextBox))

	# on button click, create token and exit
	$changePasswordButton.Add_Click({	changePassword $SAMTextBox.Text $tokenPTextBox.Text  
								$form.close()})

	$tokenPTextBox.Add_KeyDown({	
							if ($_.KeyCode -eq "Enter") 
							{  	
								changePassword $SAMTextBox.Text $tokenPTextBox.Text  
								$form.close() 
							}
						})
     
	[void]$form.ShowDialog()
}

generateForm 
