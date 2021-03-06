# Alan Bishop 
# last updated 3/6/2021
#
# Asks for and stores password token encrypted on local drive under user profile. 
# Does not check password or SAM account name for accuracy.
# Token will only work on the computer it was created on.
#
# ************************************************************************************************************
# ************************************************************************************************************
# *                       NEVER copy this token to a USB drive, laptop or onto the internet                  *
# * NEVER use this to store domain admin credentials, create a service account with min priveleges instead!  *
# *                        If token is lost the users password must be changed immediately                   *
# ************************************************************************************************************
# ************************************************************************************************************
#
# usage:
# 		.\tokenpassword.ps1 				use the GUI to create tokens
# 		.\tokenpassword.ps1 $pass 			creates token using the password provided, then clears screen
# 		.\tokenpassword.ps1 $user $pass 		creates a token for $user based on the password provided, then clears the screen.
# 								note that $user must be the users SAM name (eg John Smith's SAM would be jsmith)


Function createToken {
	param (	[Parameter(Mandatory=$true, Position=0)] [string] $SAM, 
			[Parameter(Mandatory=$true, Position=1)] [string] $tokenP )


	$tokenfile = "$($SAM).enc"
	ConvertTo-SecureString -string $tokenP -asplaintext -force | ConvertFrom-SecureString | out-file $env:userprofile\$tokenfile
}

# this function that creates the GUI
Function generateForm {
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	$form = New-Object system.Windows.Forms.Form

	$form.ClientSize = '500,300'
	$form.BackColor  = '#000000'
	$form.ForeColor  = '#ffffff'
	$form.text       = 'NEVER copy token to USB drive, laptop or any online resource!'

	# setup the label that describes the first name text box purpose
	$SAM          = New-Object system.Windows.Forms.Label
	$SAM.text     = "Users login name ex: abishop (use 'allnew' for Create-NewUser script)"
	$SAM.AutoSize = $true
	$SAM.location = New-Object System.Drawing.Point(10,70)

	# setup the SAM text box
	$SAMTextBox          = New-Object System.Windows.Forms.TextBox
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
	$createTokenButton           = New-Object system.Windows.Forms.Button
	$createTokenButton.text      = "Create Token"
	$createTokenButton.width     = 150
	$createTokenButton.height    = 30
	$createTokenButton.location  = New-Object System.Drawing.Point(270,250)
	$createTokenButton.Font      = 'Microsoft Sans Serif,10'
	$createTokenButton.ForeColor = "#000000"
	$createTokenButton.BackColor = "#dddddd"
	$createTokenButton.Visible   = $true


	# Compile and display the form
	$form.controls.AddRange(@($createTokenButton,$SAM,$SAMTextBox,$tokenP,$tokenPTextBox))

	# on button click, create token and exit
	$createTokenButton.Add_Click({	createToken $SAMTextBox.Text $tokenPTextBox.Text  
								$form.close()})

	[void]$form.ShowDialog()
}

# launch GUI if no args are passed
if ($args.count -eq 0)
{
	generateForm
}
# if an argument is passed, use as password using the currently logged in SAM
elseif ($args.count -eq 1) 
{
	createToken $env:username $args[0]
}
# else assume 1st field SAM, 2nd field password, ignore other fields
else
{
	createToken $args[0] $args[1]
}

# clear screen in 
cls
echo "token file saved: $tokenfile"
