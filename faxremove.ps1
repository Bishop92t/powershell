# Alan Bishop 
# updated 7/16/2019
#
# removes someone to the fax distribution email group 
#
# usage:
# run without arguments to see a list of all members          ./faxremove.ps1
# run with arguments to remove someone to the list            ./faxremove.ps1 abishop



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

# if args were passed, remove user from both fax groups
if (-not ($args.count -eq 0))
{
	$email = Get-Content -Path '.\debloat files\email.txt'
	$email = $args[0] + $email
	Write-Host "removing $email from the fax distribution list"
	foreach ($faxemail in Get-Content ".\debloat files\faxemail.txt")
	{
		Remove-DistributionGroupMember -Identity $faxemail -Member $email -Confirm:$false
	}
}
# else list all current members
else
{
	$faxemail = Get-Content -Path '.\debloat files\faxemail.txt' -TotalCount 1
	Get-DistributionGroupMember -Identity $faxemail
}

# if this script connected to Exchange, than disconnect
if ($connected)
{
	echo " "
	echo "diconnecting"
	.\exchangedisconnect.ps1
}
# else leave connections how they were
else 
{
	echo " "
	echo "pre-existing connection detected, not disconnecting"
}