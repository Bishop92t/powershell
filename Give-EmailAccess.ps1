# created by Alan Bishop
# last updated 6/21/2021
#
# GUI presents a list of users who are unlicensed, allowing the script runner to select a user to license


# check to see where the scripts are being run from
if (Test-Path "c:\script\Give-EmailAccess.ps1")
{
	$sPath = "c:\script\"
	$ec = "c:\script\exchangeconnect.ps1"
}
elseif (Test-Path "c:\script\ps1\Give-EmailAccess.ps1")
{
	$sPath = "c:\script\ps1"
	$ec = "c:\script\ps1\exchangeconnect.ps1"
}
else
{
	Write-Output "difficulty locating PS1 path"
}

# if there is no connection to Office 365, then create
if($null -eq (get-pssession | where-object {$_.ComputerName -EQ 'outlook.office365.com'}))
{
	& $ec
	$connected = $true
	Write-Output "Attempting to connect"
}
# else set flag that connection wasnt established (so this script doesnt disconnect pre-existing)
else
{
	$connected = $false
	Write-Output "Already connected"
}


# this function creates the form the user will fill out
Function generateForm {
	param (	[Parameter(Mandatory=$true, Position=1)] $users	)

	# create an array of users
	$i = 0
	$userArray = New-Object string[] $users.count
	foreach ($user in $users)
	{
		$userArray[$i] = $user
		$i++
	}

	# start generating the form
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	$form   = New-Object system.Windows.Forms.Form

	# set the size and features of the window
	$form.ClientSize = "700,400"
	$form.BackColor  = '#000000'
	$form.ForeColor  = '#ffffff'
	$form.text       = 'Select user to license'

	# setup the descriptive label
	$userListLabel          = New-Object system.Windows.Forms.Label
	$userListLabel.text     = "Select the user to give email license"
	$userListLabel.AutoSize = $true
	$userListLabel.location = New-Object System.Drawing.Point(10,10)

	# setup button for 1st user
	$radioButton0	        = New-Object system.Windows.Forms.RadioButton
	$radioButton0.Location = New-Object System.Drawing.Size(20, 45)
	$radioButton0.Size     = New-Object System.Drawing.Size(500,30)
	$radioButton0.Font     = 'Microsoft Sans Serif,12'
	$radioButton0.text     = $userArray[0]
	$form.controls.Add($radioButton0) 

	# setup button for 2nd user
	$radioButton1	        = New-Object system.Windows.Forms.RadioButton
	$radioButton1.Location = New-Object System.Drawing.Size(20, 65)
	$radioButton1.Size     = New-Object System.Drawing.Size(500,30)
	$radioButton1.Font     = 'Microsoft Sans Serif,12'
	$radioButton1.text     = $userArray[1]
	# only add the button if there is more than 1 user
	if ($i -gt 1)
	{
		$form.controls.Add($radioButton1) 
	}	

	# setup button for 3rd user
	$radioButton2	        = New-Object system.Windows.Forms.RadioButton
	$radioButton2.Location = New-Object System.Drawing.Size(20, 85)
	$radioButton2.Size     = New-Object System.Drawing.Size(500,30)
	$radioButton2.Font     = 'Microsoft Sans Serif,12'
	$radioButton2.text     = $userArray[2]
	if ($i -gt 2)
	{
		$form.controls.Add($radioButton2) 
	}

	# setup button for 4th user
	$radioButton3	        = New-Object system.Windows.Forms.RadioButton
	$radioButton3.Location = New-Object System.Drawing.Size(20, 105)
	$radioButton3.Size     = New-Object System.Drawing.Size(500,30)
	$radioButton3.Font     = 'Microsoft Sans Serif,12'
	$radioButton3.text     = $userArray[3]
	if ($i -gt 3)
	{
		$form.controls.Add($radioButton3) 
	}	

	# setup button for 5th user
	$radioButton4	        = New-Object system.Windows.Forms.RadioButton
	$radioButton4.Location = New-Object System.Drawing.Size(20, 125)
	$radioButton4.Size     = New-Object System.Drawing.Size(500,30)
	$radioButton4.Font     = 'Microsoft Sans Serif,12'
	$radioButton4.text     = $userArray[4]
	if ($i -gt 4)
	{
		$form.controls.Add($radioButton4) 
	}

	# setup the license user button
	$licenseUserButton           = New-Object system.Windows.Forms.Button
	$licenseUserButton.text      = "License User"
	$licenseUserButton.width     = 110
	$licenseUserButton.height    = 30
	$licenseUserButton.location  = New-Object System.Drawing.Point(300,250)
	$licenseUserButton.Font      = 'Microsoft Sans Serif,10'
	$licenseUserButton.ForeColor = "#000000"
	$licenseUserButton.BackColor = "#dddddd"
	$licenseUserButton.Visible   = $true

	# Compile and display the form
	$form.controls.AddRange(@($userListLabel,$radioButton0,$licenseUserButton))

	# activate whichever user's radio button is checked
	$licenseUserButton.Add_Click({ 
								if ($radioButton0.Checked -eq $true)
								{
									licenseUser $userArray[0]
								}
								elseif ($radioButton1.Checked -eq $true)
								{
									licenseUser $userArray[1]
								}
								elseif ($radioButton2.Checked -eq $true)
								{
									licenseUser $userArray[2]
								}
								elseif ($radioButton3.Checked -eq $true)
								{
									licenseUser $userArray[3]
								}
								else
								{
									licenseUser $userArray[4]
								}
								$form.Close()
							})

	[void]$form.ShowDialog()
}


Function licenseUser {
	param (	[Parameter(Mandatory=$true, Position=1)] [string] $userUPN	)

	# email.txt : the email domain in the format  @$company.com
	$defaultDomain = Get-Content -Path 'c:\script\debloat files\email.txt'

	# set user location to US, then add O365 license
	Set-MsolUser -UserPrincipalName $userUPN -UsageLocation US

	# o365license.txt : License for the company - consists of company name and which O365 product
	$license = Get-Content -Path 'c:\script\debloat files\o365license.txt'
	Set-MsolUserLicense -UserPrincipalName $userUPN -AddLicenses $license 
} 


# list of active accounts to ignore since they don't need email
# excludedfromemail.txt : a list of SAM's, each on a new line
$excludedUsers = Get-Content -Path 'c:\script\debloat files\excludedfromemail.txt'

# setup flag for unlicensed user true/false, setup array, get a list of all unlicensed users
$noUnlicensedFound = $true
$unlicensedUsers = @()
$users = Get-MsolUser -All -UnlicensedUsersOnly

# for all unlicensed users, if they aren't in the exclusion list then show which users don't have email
foreach ($user in $users)
{
	if (!($excludedUsers -contains $user.DisplayName))
	{
		$noUnlicensedFound = $false
		$unlicensedUsers += $user.UserPrincipalName
	}
}

# if no unlicensed users found inform user, otherwise create the window for them to interact with
if ($noUnlicensedFound -eq $true)
{
	[System.Windows.Forms.MessageBox]::Show("No unlicensed users found. If you just created the account it can take several minutes to show up" , "nothing found")
}
else
{
	generateForm $unlicensedUsers
} 

# if this script connected to Exchange, than disconnect
if ($connected)
{
	"$($sPath)exchangedisconnect.ps1"
}
# else leave connections how they were
else 
{
	Write-Output "`n pre-existing connection detected, not disconnecting"
}
