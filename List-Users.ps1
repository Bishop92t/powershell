# created by Alan Bishop
# last modified 2/3/2021
#
# Various AD user account compiled info
# usage:
#  		.\List-Users.ps1    		displays full usage of this script



if ($args.count -eq 0)
{
	Write-Output '.\List-Users.ps1 save            save most properties to users.csv'
	Write-Output '.\List-Users.ps1 password        display regular users sorted by last time they reset their password'
	Write-Output '.\List-Users.ps1 locked          display list of locked accounts'
	Write-Output '.\List-Users.ps1 name $name      display info just about $name'
	Write-Output '.\List-Users.ps1 like $name      display user info for users with similar names to $name'
	Write-Output '.\List-Users.ps1 in $group       display all active users in $group'
	Write-Output '.\List-Users.ps1 notin $group    display all active users not in $group'
	Write-Output '.\List-Users.ps1 noexpire        display all active users whose password doesnt expire'
	Write-Output '.\List-Users.ps1 last            display active users sorted by last logon (all users)'
	Write-Output '.\List-Users.ps1 lastlogin       display when all active users last logged on (excluding service accounts)'
	Write-Output '.\List-Users.ps1 date            display a list of all users sorted by creation date'
	Write-Output '.\List-Users.ps1 forwarding      display a list of users with server-side email forwarding rules'
	Write-Output '.\List-Users.ps1 emailrule       display a list of users with Outlook email rules (will show bad rules)'
	Write-Output '.\List-Users.ps1 emailforward    display a list of users with Outlook email forwarding rules'
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


# for viewing server-side email forwarding rules
elseif ($args[0] -eq "forwarding")
{	
	# if there is no connection to Office 365, then create
	if($null -eq (get-pssession | where-object {$_.ComputerName -EQ 'outlook.office365.com'}))
	{
		.\exchangeconnect.ps1 | Out-Null
		$connected = $true
	}
	# else set flag that connection wasnt established (so this script doesnt disconnect pre-existing)
	else
	{
		$connected = $false
	}
	# view list of all mailboxes that have a forwarding address set on the server
	Get-Mailbox -Resultsize Unlimited | where-object {($null -ne $_.ForwardingSmtpAddress) -or ($null -ne $_.ForwardingAddress)}
	# if this script connected to Exchange, then disconnect, else do nothing
	if ($connected)
	{
		.\exchangedisconnect.ps1
	}
}


# for viewing a count of all client side (Outlook) email rules
elseif ($args[0] -eq "emailrule")
{
	# to-do: drill down into rules more
	# in the interim run this for more info:    get-inboxrule -mailbox user@domain.com | format-list
	# Get-InboxRule -mailbox nbishop@hospicebr.org | select description

	Write-Output "be patient, this one takes a while to run"

	# if there is no connection to Office 365, then create
	if($null -eq (get-pssession | where-object {$_.ComputerName -EQ 'outlook.office365.com'}))
	{
		.\exchangeconnect.ps1 | Out-Null
		$connected = $true
	}
	# else set flag that connection wasnt established (so this script doesnt disconnect pre-existing)
	else
	{
		$connected = $false
	}
	$mailboxes = Get-Mailbox
	foreach ($mailbox in $mailboxes)
	{
		$numRules = (Get-InboxRule -mailbox $mailbox.UserPrincipalName).count
		if ($numRules -gt 0)
		{
			Write-Output " $($mailbox.UserPrincipalName)   $($numRules)"
		}
	}
	# if this script connected to Exchange, then disconnect, else do nothing
	if ($connected)
	{
		.\exchangedisconnect.ps1
	}
}


# for viewing client side (Outlook) forwarding email rules, differs from above by showing any client-side forwarding
elseif ($args[0] -eq "emailforward")
{
	# to-do: drill down into rules more
	# in the interim run this for more info:    get-inboxrule -mailbox user@domain.com | format-list
	# Get-InboxRule -mailbox nbishop@hospicebr.org | select description

	Write-Output "be patient, this one takes a while to run"

	# if there is no connection to Office 365, then create
	if($null -eq (get-pssession | where-object {$_.ComputerName -EQ 'outlook.office365.com'}))
	{
		.\exchangeconnect.ps1 | Out-Null
		$connected = $true
	}
	# else set flag that connection wasnt established (so this script doesnt disconnect pre-existing)
	else
	{
		$connected = $false
	}
	$mailboxes = Get-Mailbox
	foreach ($mailbox in $mailboxes)
	{
		$rules = (Get-InboxRule -mailbox $mailbox.UserPrincipalName | ? {($null -ne $_.RedirectTo) -or ($null -ne $_.ForwardTo)})
		if ($rules.count -gt 0)
		{
			Write-Output " $($mailbox.UserPrincipalName)   $($rules.Description)"
		}
	}
	# if this script connected to Exchange, then disconnect, else do nothing
	if ($connected)
	{
		.\exchangedisconnect.ps1
	}
}
