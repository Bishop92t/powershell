# created by Alan Bishop
# last updated 6/21/2021
#
# Launches a GUI that creates a user in AD / Exchange 365 and assigns them to the proper groups


# global variable, flagged true if an account was created
$global:accountCreated = "false"

# declare some global variables
$global:gfirst = ""
$global:glast  = ""
$global:grole  = ""
$global:gupn   = ""

# save a log of what this script has accomplished
if (!(Test-Path c:\logs))
{
	New-Item "c:\logs" -ItemType "directory"
}
Start-Transcript -path c:\logs\backup-script.txt -force -append

# check if the machine has the correct module, if not install all req modules
$checkADModule    = (Get-Module ActiveDirectory -ListAvailable).Name
if ($checkADModule -eq $null)
{
	if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) 
	{
		if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) 
		{
			$CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
			Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
			Exit
		}
	}
	Write-Host "Important: Active Directory Powershell module must be installed to continue, attempting install "
	Install-PackageProvider -Name NuGet -Force
	$RSATname = Get-WindowsCapability -name "RSAT.ActiveDirectory*" -online | select name
	Add-WindowsCapability -Online -name $RSATname.name
}

$netDrive = "c:\script\debloat files"

# this function that creates the GUI
Function generateForm {
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	$form = New-Object system.Windows.Forms.Form

	$form.ClientSize = '500,300'
	$form.BackColor  = '#000000'
	$form.ForeColor  = '#ffffff'
	$form.text       = 'Creating new user'

	$roleTypeLabel          = New-Object system.Windows.Forms.Label
	$roleTypeLabel.text     = "Select the type of user to add"
	$roleTypeLabel.AutoSize = $true
	$roleTypeLabel.location = New-Object System.Drawing.Point(10,10)

	# setup the dropdown
	$roleType          		= New-Object system.Windows.Forms.ComboBox
	$roleType.text     		= ""
	$roleType.width    		= 170
	$roleType.dropdownheight = 200

	# add options to the dropdown
	@('Chaplain', 'CNA', 'IPU CNA', 'IPU RN', 'LPN', 'NP', 'RN', 'Social Worker', 'Other') | ForEach-Object {[void] $roleType.Items.Add($_)}

	# set the default value
	$roleType.SelectedIndex = 0
	$roleType.location      = New-Object System.Drawing.Point(20,30)
	$roleType.Font          = 'Microsoft Sans Serif,10'

	# setup the label that describes the first name text box purpose
	$firstNameLabel          = New-Object system.Windows.Forms.Label
	$firstNameLabel.text     = "Users first name:"
	$firstNameLabel.AutoSize = $true
	$firstNameLabel.location = New-Object System.Drawing.Point(10,70)

	# setup the first name text box
	$firstTextBox          = New-Object System.Windows.Forms.TextBox
	$firstTextBox.Size     = New-Object System.Drawing.Size(260,20)
	$firstTextBox.Location = New-Object System.Drawing.Point(20,90)
	$form.Controls.Add($firstTextBox)

	# setup the label that describes the last name text box purpose
	$lastNameLabel          = New-Object system.Windows.Forms.Label
	$lastNameLabel.text     = "Users last name:"
	$lastNameLabel.AutoSize = $true
	$lastNameLabel.location = New-Object System.Drawing.Point(10,120)

	# setup the last name text box
	$lastTextBox          = New-Object System.Windows.Forms.TextBox
	$lastTextBox.Size     = New-Object System.Drawing.Size(260,20)
	$lastTextBox.Location = New-Object System.Drawing.Point(20,140)
	$form.Controls.Add($lastTextBox)

	# setup the create user button
	$createUserButton           = New-Object system.Windows.Forms.Button
	$createUserButton.text      = "Create User"
	$createUserButton.width     = 90
	$createUserButton.height    = 30
	$createUserButton.location  = New-Object System.Drawing.Point(270,250)
	$createUserButton.Font      = 'Microsoft Sans Serif,10'
	$createUserButton.ForeColor = "#000000"
	$createUserButton.BackColor = "#dddddd"
	$createUserButton.Visible   = $true

	# setup the exit button
	$exitButton           = New-Object system.Windows.Forms.Button
	$exitButton.text      = "Sync and Exit"
	$exitButton.width     = 120
	$exitButton.height    = 30
	$exitButton.location  = New-Object System.Drawing.Point(370,250)
	$exitButton.Font      = 'Microsoft Sans Serif,10'
	$exitButton.ForeColor = "#000000"
	$exitButton.BackColor = "#dddddd"
	$exitButton.Visible   = $true

	# Compile and display the form
	$form.controls.AddRange(@($roleTypeLabel,$roleType,$createUserButton,$firstNameLabel,$firstTextBox,$lastNameLabel,$lastTextBox,$exitButton))

	# if user hits enter, simulate what createUserButton does
	$lastTextBox.Add_KeyDown({	
							if ($_.KeyCode -eq "Enter") 
							{  	
								createNewUser $roleType.Items[$roleType.SelectedIndex] $firstTextBox.Text $lastTextBox.Text 
							}
						})

	# creates new user on button click
	$createUserButton.Add_Click	({ 	
								createNewUser $roleType.Items[$roleType.SelectedIndex] $firstTextBox.Text $lastTextBox.Text
							})

	# syncs all held data then closes this script
	$exitButton.Add_Click	({	
							syncExit
							$form.close() 
						})

	[void]$form.ShowDialog()
}


# this function does the actual creation of an account
function createNewUser{
	param (
		[Parameter(Mandatory=$true, Position=0)] [string] $role, 
		[Parameter(Mandatory=$true, Position=1)] [string] $firstname, 
		[Parameter(Mandatory=$true, Position=2)] [string] $lastname 
	)

	# if the token isn't found run a script that will ask for it
	if (-not (Test-Path $env:userprofile\allnew.enc))
	{
		write-output "no token found"
		if (Test-Path c:\script\tokenpassword.ps1)
		{
			write-output "token found .\script"
			c:\script\tokenpassword.ps1 gui allnew "Type the default new users password (this will be changed by user later), leave username as allnew"
		}
		elseif (Test-Path c:\script\PS1\tokenpassword.ps1)
		{
			write-output "token found .\script\ps1\"
			c:\script\PS1\tokenpassword.ps1 gui allnew "Type the default new users password (this will be changed by user later), leave username as allnew"
		}
	}

	# start creating the user if the new user token has been set
	else
	{
		# the temp password is saved as a token in $env:userprofile\allnew.enc
		# you can change this password with .\tokenpassword.ps1 1 allnew $newpassword (user needs to change ASAP)
		$newpassword = Get-Content $env:userprofile\allnew.enc | ConvertTo-SecureString

		# concatenate the name, create a succesful creation flag which will be set to true if a user gets created
		$fullname = "$firstname $lastname"

		# increments SAM, name will be set when $nameExists is found to be null
		$SAMincrementer = 0
		$nameExists = " . "

		# try to look for a SAM name that isn't taken, exiting loop when a free one is found
		# eg: if jwilliams exists try jwilliams1, jwilliams2 etc 
		# this solves edge case of matching first initial and matching last name
		while ($nameExists -ne $Null)
		{
			# try a SamAccountName, first initial + last name, incrementing numbers until an opening is found
			if ($SAMincrementer -eq 0)
			{
				$sam = $firstname.Substring(0,1)+$lastname
			}
			else
			{
				$sam = $firstname.Substring(0,1)+$lastname+($SAMincrementer.ToString())
			}
			# this will return null if the name isn't taken
			$nameExists = Get-ADUser -Filter {SAMAccountName -eq $sam}
			$SAMincrementer = $SAMincrementer+1
		} 

		# if full name is taken, add the incrementer to it (the 2nd Jeff Williams becomes Jeff Williams1) 
		# this solves edge case of two Jeff Williams
		if ((Get-ADUser -Filter {DisplayName -eq $fullname}) -ne $Null)
		{
			$SAMincrementer--     #match SAM number to fullname number
			$fullname = "$fullname $SAMincrementer"
		}

		# create the UPN (email)
		# email.txt : the email domain in the format  @$company.com
		$upn = $sam+(Get-Content -Path "$($netDrive)\email.txt")

		$nonTemplateUser = $false

		# add the created user's info to a temp holding file
		$global:gfirst = $firstname
		$global:glast  = $lastname
		$global:grole  = $role
		$global:gupn   = $upn

		if ($role -eq "CNA")
		{
			# add the user to AD, enable the account, mark account as created
			New-ADUser -Name $fullname -DisplayName $fullname -GivenName $firstname -Surname $lastname -SamAccountName $sam -UserPrincipalName $upn -EmailAddress $upn -Description "CNA" -AccountPassword($newpassword)
			Enable-ADAccount -Identity $sam
			# cna.txt : just a list of groups a person in this role should belong to, each group on a new line
			$cnas = Get-Content -Path "$($netDrive)\cna.txt"
			foreach ($cna in $cnas)
			{
				Add-ADGroupMember -Identity $cna -Members $sam
			}
			$global:accountCreated = "true"
		}
		elseif ($role -eq "IPU CNA")
		{
			# add the user to AD, enable the account, mark account as created
			New-ADUser -Name $fullname -DisplayName $fullname -GivenName $firstname -Surname $lastname -SamAccountName $sam -UserPrincipalName $upn -EmailAddress $upn -Description "IPU CNA" -AccountPassword($newpassword)
			Enable-ADAccount -Identity $sam
			# ipucna.txt : just a list of groups a person in this role should belong to, each group on a new line
			$ipucnas = Get-Content -Path "$($netDrive)\ipucna.txt"
			foreach ($ipucna in $ipucnas)
			{
				Add-ADGroupMember -Identity $ipucna -Members $sam
			}
			$global:accountCreated = "true"
		}
		elseif ($role -eq "IPU RN")
		{
			# add the user to AD, enable the account, mark account as created
			New-ADUser -Name $fullname -DisplayName $fullname -GivenName $firstname -Surname $lastname -SamAccountName $sam -UserPrincipalName $upn -EmailAddress $upn -Description "IPU RN" -AccountPassword($newpassword)
			Enable-ADAccount -Identity $sam
			# ipurn.txt : just a list of groups a person in this role should belong to, each group on a new line
			$ipurns = Get-Content -Path "$($netDrive)\ipurn.txt"
			foreach ($ipurn in $ipurns)
			{
				Add-ADGroupMember -Identity $ipurn -Members $sam
			}
			# IPU RN's aren't given login access, only emails, so set password to never expire
			$global:accountCreated = "true"
		}
		elseif ($role -eq "RN")
		{
			# add the user to AD, enable the account, mark account as created
			New-ADUser -Name $fullname -DisplayName $fullname -GivenName $firstname -Surname $lastname -SamAccountName $sam -UserPrincipalName $upn -EmailAddress $upn -Description "RN" -AccountPassword($newpassword)
			Enable-ADAccount -Identity $sam
			# rn.txt : just a list of groups a person in this role should belong to, each group on a new line
			$rns = Get-Content -Path "$($netDrive)\rn.txt"
			foreach ($rn in $rns)
			{
				Add-ADGroupMember -Identity $rn -Members $sam
			}
			$global:accountCreated = "true"
		}
		elseif ($role -eq "NP")
		{
			# add the user to AD, enable the account, mark account as created
			New-ADUser -Name $fullname -DisplayName $fullname -GivenName $firstname -Surname $lastname -SamAccountName $sam -UserPrincipalName $upn -EmailAddress $upn -Description "NP" -AccountPassword($newpassword)
			Enable-ADAccount -Identity $sam
			# np.txt : just a list of groups a person in this role should belong to, each group on a new line
			$nps = Get-Content -Path "$($netDrive)\np.txt"
			foreach ($np in $nps)
			{
				Add-ADGroupMember -Identity $np -Members $sam
			}
			$global:accountCreated = "true"
		}
		elseif ($role -eq "LPN")
		{
			# add the user to AD, enable the account, mark account as created
			New-ADUser -Name $fullname -DisplayName $fullname -GivenName $firstname -Surname $lastname -SamAccountName $sam -UserPrincipalName $upn -EmailAddress $upn -Description "LPN" -AccountPassword($newpassword)
			Enable-ADAccount -Identity $sam
			# lpn.txt : just a list of groups a person in this role should belong to, each group on a new line
			$lpns = Get-Content -Path "$($netDrive)\lpn.txt"
			foreach ($lpn in $lpns)
			{
				Add-ADGroupMember -Identity $lpn -Members $sam
			}
			$global:accountCreated = "true"
		}	
		elseif ($role -eq "Chaplain")
		{
			# add the user to AD, enable the account, mark account as created
			New-ADUser -Name $fullname -DisplayName $fullname -GivenName $firstname -Surname $lastname -SamAccountName $sam -UserPrincipalName $upn -EmailAddress $upn -Description "Chaplain" -AccountPassword($newpassword)
			Enable-ADAccount -Identity $sam
			# chap.txt : just a list of groups a person in this role should belong to, each group on a new line
			$chaps = Get-Content -Path "$($netDrive)\chap.txt"
			foreach ($chap in $chaps)
			{
				Add-ADGroupMember -Identity $chap -Members $sam
			}
			$global:accountCreated = "true"
		}
		elseif ($role -eq "Social Worker")
		{
			# add the user to AD, enable the account, mark account as created
			New-ADUser -Name $fullname -DisplayName $fullname -GivenName $firstname -Surname $lastname -SamAccountName $sam -UserPrincipalName $upn -EmailAddress $upn -Description "SW" -AccountPassword($newpassword)
			Enable-ADAccount -Identity $sam
			# sw.txt : just a list of groups a person in this role should belong to, each group on a new line
			$sws = Get-Content -Path "$($netDrive)\sw.txt"
			foreach ($sw in $sws)
			{
				Add-ADGroupMember -Identity $sw -Members $sam
			}
			$global:accountCreated = "true"
		}
		else
		{
			$nonTemplateUser = $true
			# remove newline character so it doesn't mess up adding to AD
			$role = $role -Replace('\n','')
			# add to AD, enable account, then mark account as created
			New-ADUser -Name $fullname -DisplayName $fullname -GivenName $firstname -Surname $lastname -SamAccountName $sam -UserPrincipalName $upn -EmailAddress $upn -Description $role -AccountPassword($newpassword)
			Enable-ADAccount -Identity $sam
			$global:accountCreated = "true"
		}

		# if the account was created, set the groups that everyone gets
		if ($global:accountCreated -eq "true")
		{
			# moving to a complex password, never expire system
			Set-ADUser -Identity $sam -PasswordNeverExpires:$TRUE

			# everyonegroups.txt : just a list of groups everyone goes in, each group on a new line
			$egs = Get-Content -Path "$($netDrive)\everyonegroups.txt"
			foreach ($eg in $egs)
			{
				Add-ADGroupMember -Identity $eg -Members $sam
			}
			
			# grabs the users groups and compiles into a more readable format
			# to-do, make this a cleaner solution
			$groups = ''
			Get-ADPrincipalGroupMembership $sam | Select name | Sort-Object -property name | ForEach-Object {$groups += " $_ "}
			$groupsmess = $groups -replace ' @{name=', ''
			$groups = $groupsmess -replace '}', ','

			# $($first),,$($last),$($upn),($(role),,y,y,n,y,IPU All Staff,Emergency Group,,$($word)"

			# inform that user was created successfully
			if ($SAMincrementer -le 1)
			{
				[System.Windows.Forms.MessageBox]::Show("$fullname user account created for role $role `n  login will be $sam " , "Success!")
			}
			# if the default schema was not used, display warning about non-standard SAM
			else
			{
				[System.Windows.Forms.MessageBox]::Show("non-standard account name found! `n Please note the users login will be $sam `n $fullname user account created for role $role" , "Warning!")
			}
		}
	}

	$form2 = New-Object system.Windows.Forms.Form

	$form2.ClientSize = '500,300'
	$form2.BackColor  = '#000000'
	$form2.ForeColor  = '#ffffff'
	$form2.text       = 'Creating new user'
}

# takes an array of objects and dumps them into a CSV file on the users desktop
function createQliqImport
{
	# setup the path and file names to be formatted
	$path = "c:\users\$($env:username)\desktop"
	$file = "$($path)\qliqImport.csv"

	Add-Content $file ("First Name,Mid Name,Last Name,Email/Mobile,Title,Department,Full Group Access, Mobile App Login, Broadcasting, Group Messaging, Group1, Group2,Group3,Password")


	for ($i = 0; $i -lt $content.length; $i = $i + 4)
	{
		$first = $global:gfirst
		$last  = $global:glast 
		$role  = $global:grole 
		$upn   = $global:gupn 

		if ($role -eq "IPU CNA") 
		{
			$word = Get-Content "$($netDrive)\qliqcna.txt"
			$addL = "$($first),,$($last),$($upn),($(role),,y,y,n,y,IPU All Staff,Emergency Group,,$($word)"
		}
		elseif ($role -eq "IPU RN")
		{
			$word = Get-Content "$($netDrive)\qliqrn.txt"
			$addL = "$($first),,$($last),$($upn),$($role),,y,y,n,y,IPU All Staff,Emergency Group,,$($word)"
		}
		elseif ($role -eq "CNA")
		{
			$word = Get-Content "$($netDrive)\qliqcna.txt"
			$addL = "$($first),,$($last),$($upn),$($role),,y,y,n,y,All Group Daily Messaging,Emergency Group,CNAs (Field Only),$($word)"
		}
		elseif (($role -eq "RN") -or ($role -eq "LPN"))
		{
			$word = Get-Content "$($netDrive)\qliqrn.txt"
			$addL = "$($first),,$($last),$($upn),$($role),,y,y,n,y,All Group Daily Messaging,Emergency Group,Nurses (Field Only),$($word)"
		}
		elseif ($role -eq "NP")
		{
			$word = Get-Content "$($netDrive)\qliqrn.txt"
			$addL = "$($first),,$($last),$($upn),$($role),,y,y,n,y,All Group Daily Messaging,Emergency Group,NPs (HC only),$($word)"
		}
		elseif ($role -eq "Social Worker")
		{
			$word = Get-Content "$($netDrive)\qliqsw.txt"
			$addL = "$($first),,$($last),$($upn),$($role),,y,y,n,y,All Group Daily Messaging,Emergency Group,Social Services (All),$($word)"
		}
		elseif ($role -eq "Chaplain")
		{
			$word = Get-Content "$($netDrive)\qliqsw.txt"
			$addL = "$($first),,$($last),$($upn),$($role),,y,y,n,y,All Group Daily Messaging,Emergency Group,Social Services (All),$($word)"
		}
		else
		{
			$word = Get-Content "$($netDrive)\qliqad.txt"
			$addL = "$($first),,$($last),$($upn),$($role),,y,y,n,y,All Group Daily Messaging,Emergency Group,,$($word)"
		}
		Add-Content $file $addL
	}

}

# once all new accounts are added, this will sync AD,export a Qliq CSV, and exit
function syncExit 
{
	if (Test-Path "$($netDrive)\psexec.exe")
	{
		Write-Output "psexe  $(Get-Date) " >> ".\atemplog.txt"
		# psexec.txt : a simple command that forces an Exchange 365 sync with on-site AD. Not required but is faster
		$psexe = Get-Content -Path "$($netDrive)\psexec.txt"
		Start-Process -FilePath "$($netDrive)\psexec.exe" -ArgumentList $psexe
	}
	else 
	{
		Write-Output "no psexe  $(Get-Date) " >> ".\atemplog.txt"
		[System.Windows.Forms.MessageBox]::Show("PSEXEC not found, password sync may take up to 30 minutes" , "Warning")
	}

	# delete old Qliq import CSV and create new one
	$path = "c:\users\$($env:username)\desktop"
	$file = "$($path)\qliqImport.csv"
	Remove-Item -Path $file -Force
	createQliqImport
}


generateForm
