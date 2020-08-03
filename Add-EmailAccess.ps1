# created by Alan Bishop
# last updated 8/3/2020
#
# Grant a user ability to access email (and everything O365 offers)
# usage:
#		.\Add-EmailAccess.ps1  	
# 		.\Add-EmailAccess.ps1 abishop@yourcompany.com 		gives abishop@yourcompany.com access
# 		.\Add-EmailAccess.ps1 abishop						gives abishop access at default domain
#
# exchangeconnect.ps1 and exchangedisconnect.ps1 are a required dependencies


# email.txt : the email domain in the format  @$company.com
$defaultDomain = Get-Content -Path '.\debloat files\email.txt'
# email2.txt : the email domain in the format @$secondcompany.com
$domain2       = Get-Content -Path '.\debloat files\email2.txt'

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

# if a UPN is provided give it email access, otherwise it's a SAM so concatenate the default domain to create a UPN
if ($args.count -ne 0)
{
	$userUPN = $args[0]

	# if the domain isn't provided, add default
	if (($userUPN -like "*$($defaultDomain)") -or ($userUPN -like "*$($domain2)"))
	{
		#do nothing
	}
	else
	{
		$userUPN = "$($userUPN)$($defaultDomain)"
	}

	# set user location to US, then add O365 license
	Set-MsolUser -UserPrincipalName $userUPN -UsageLocation US
	# o365license.txt : License for the company - consists of company name and which O365 product
	$license = Get-Content -Path '.\debloat files\o365license.txt'
	Set-MsolUserLicense -UserPrincipalName $userUPN -AddLicenses $license
	echo "added $userUPN"
}
# else compile a list of active accounts that don't have email access, this should be blank usually
else
{
	# list of active accounts to ignore since they don't need email
	# excludedfromemail.txt : a list of SAM's, each on a new line
	$excludedUsers = Get-Content -Path '.\debloat files\excludedfromemail.txt'

	$noUnlicensedFound = $true
	$users = Get-MsolUser -All -UnlicensedUsersOnly

	echo ' '
	echo 'usage:        .\emailadd.ps1 user@yourcompany.com'
	echo ' '
	echo ' '

	# for all unlicensed users, if they aren't in the exclusion list then show which users don't have email
	foreach ($user in $users)
	{
		if (!($excludedUsers -contains $user.DisplayName))
		{
			$noUnlicensedFound = $false
			echo $user
		}
	}

	if ($noUnlicensedFound -eq $true)
	{
		echo "no unlicensed users found!"
	}
}

# if this script connected to Exchange, than disconnect
if ($connected)
{
	.\exchangedisconnect.ps1
}
# else leave connections how they were
else 
{
	echo " "
	echo "pre-existing connection detected, not disconnecting"
}
