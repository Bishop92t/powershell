# created by Alan Bishop
# last updated 8/3/2020
#
# Removes a users ability to access email (and everything O365 offers)
# Since it's a deletion we'll play it safe and require the full email address
# usage:
#		.\emailremove.ps1  	
# 		.\emailremove.ps1 abishop@yourcompany.org 		removes abishop@yourcompany.org access
#
# exchangeconnect.ps1 and exchangedisconnect.ps1 are a required dependencies



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

if ($args.count -ne 0)
{
	# remove user O365 license
	# o365license.txt : License for the company - consists of company name and which O365 product	
	$license = Get-Content -Path '.\debloat files\o365license.txt'
	Set-MsolUserLicense -UserPrincipalName $args[0] -RemoveLicenses $license
}
else
{
	echo 'no args present, listing licensed users.'
	echo 'usage:       .\emailremove.ps1 abishop@yourcompany.org'
	echo ' '
	Get-MsolUser -All | Where {$_.isLicensed -eq $true} | Select DisplayName | Sort-Object -property DisplayName
}

# if this script connected to Exchange, than disconnect
if ($connected)
{
	.\exchangedisconnect.ps1
}
# else leave connections how they were
else 
{
	echo "pre-existing connection detected, not disconnecting"
}