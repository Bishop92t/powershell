# created by Alan Bishop
# last modified 10/14/2020
#
# Various AD user account compiled info
# usage:
#  		.\List-Users.ps1    		displays full usage of this script



if ($args.count -eq 0)
{
Write-Output '.\List-Users.ps1 save            save most properties to users.csv'
Write-Output '.\List-Users.ps1 password        display regular users sorted by last time they reset their password'
Write-Output '.\List-Users.ps1 locked          display list of locked accounts to userslocked.csv'
Write-Output '.\List-Users.ps1 name $name      display info just about $name'
Write-Output '.\List-Users.ps1 like $name      display user info for users with similar names to $name'
Write-Output '.\List-Users.ps1 in $group       display all active users in $group'
Write-Output '.\List-Users.ps1 notin $group    display all active users not in $group'
Write-Output '.\List-Users.ps1 noexpire        display all active users whose password doesnt expire'
Write-Output '.\List-Users.ps1 last            display active users sorted by last logon (all users)'
Write-Output '.\List-Users.ps1 lastlogin       display when all active users last logged on (excluding service accounts)'
Write-Output '.\List-Users.ps1 date            display a list of all users sorted by creation date'
}
elseif ($args[0] -eq "save")
{
	Write-Host 'saving list to c:\script\users.csv'
	Get-ADUser -Filter * -Properties * | select AccountLockoutTime, BadLogonCount, CanonicalName, CN, Created, DisplayName, Enabled, LastBadPasswordAttempt, LastLogonDate, LockedOut, logonCount, Modified, modifyTimeStamp, PasswordLastSet, PasswordNeverExpires, whenChanged, whenCreated | Export-Csv C:\script\users.csv
}
elseif ($args[0] -eq "password")
{
	# list of users or service accounts to be excluded
	$excludedUsers = Get-Content '.\debloat files\nonexpiringaccounts.txt'

	# get a list of active accounts with passwords that can expire
	$users = Get-ADUser -Filter {(enabled -Eq 'True') -And (PasswordNeverExpires -Eq 'False')} -Properties * | sort-object PasswordLastSet | select Name, PasswordLastSet

	ForEach ($user in $users)
	{
		if (-not ($excludedUsers -contains $user.Name))
		{
			Write-Host $user
		}
	}
}
elseif ($args[0] -eq "locked")
{
	# showing all accounts that are locked out
	Get-ADUser -Filter * -Properties * | where AccountLockoutTime -ne $null | select AccountLockoutTime, BadLogonCount, CN, DisplayName | Sort-Object -property AccountLockoutTime 
}
elseif ($args[0] -eq "name")
{
	if ($args.count -eq 2)
	{
		Get-ADUser -Filter * -Properties * | where SamAccountName -eq $args[1] | select AccountLockoutTime, BadLogonCount, CanonicalName, CN, Created, DisplayName, Enabled, LastBadPasswordAttempt, LastLogonDate, LockedOut, logonCount, Modified, PasswordLastSet, PasswordNeverExpires, whenChanged, whenCreated
	}
	else
	{
		Write-Host "You must type something after name. Example:   .\List-Users.ps1 name jsmith"
	}
}
elseif ($args[0] -eq "like")
{
	if ($args.count -eq 2)
	{
		$name = '*' + $args[1] + '*'
		Get-ADUser -filter 'DisplayName -like $name' -Properties * | select CanonicalName, Created, DisplayName, Enabled, LastLogonDate, LockedOut, logonCount, Modified, PasswordLastSet, SamAccountName, whenCreated
	}
	else
	{
		Write-Host "You must type something after like. Example:   .\List-Users.ps1 like smith"
	}
}
elseif ($args[0] -eq "in")
{
	$group = get-adgroup $args[1]
	Get-ADUser -Properties memberof -filter 'enabled -eq $true' | Where-Object {$group.DistinguishedName -in $_.memberof} | Select Name | Sort-Object -property Name
}
elseif ($args[0] -eq "notin")
{
	$group = get-adgroup $args[1]
	Get-ADUser -filter {(enabled -Eq 'True') -And (PasswordNeverExpires -Eq 'False')} -Properties memberof | Where-Object {$group.DistinguishedName -notin $_.memberof} | Select Name | Sort-Object -property Name
}
elseif ($args[0] -eq "noexpire")
{
	Get-ADUser -filter 'enabled -eq $true' -properties Name, PasswordNeverExpires, Description | where {$_.passwordNeverExpires -eq "true" } |  Select-Object Name, Description | Sort-Object -property Name
}
elseif ($args[0] -eq "last")
{
	Get-ADUser -Filter 'enabled -eq $true' -Properties * | select Name, LastLogonDate | Sort-Object -property LastLogonDate
}
elseif ($args[0] -eq "lastlogin")
{
	# list of users or service accounts to be excluded
	$excludedUsers = Get-Content '.\debloat files\nonexpiringaccounts.txt'

	# get list of enabled users whos passwords don't expire, and setup a variable for clean printing
	$users = Get-ADUser -Filter {(enabled -Eq 'True') -And (PasswordNeverExpires -Eq 'False')} -Properties * | select Name, LastLogonDate | Sort-Object LastLogonDate
	$printableUsers = @()

	ForEach ($user in $users)
	{
		# if the user isn't an excluded service account, add it to the array $printableUsers
		if (-not ($excludedUsers -contains $user.Name))
		{
			$printableUsers = $printableUsers + $user
		}
	}

	# finally print the output
	$printableUsers | select Name, LastLogonDate
}
elseif ($args[0] -eq "date")
{
	Get-ADUser -filter * -properties * | Select DisplayName, WhenCreated | Sort-Object -Property WhenCreated
}