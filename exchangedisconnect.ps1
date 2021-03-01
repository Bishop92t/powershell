# Alan Bishop 
# Last modified 7/16/19
#
# disconnects this session from Exchange Online, intended to be a helper script but can be run solo
#
# usage:
#       check if connected only to exchange, if so then disconnect     ./exchangedisconnect.ps1
#       use to disconnect all remote sessions                          ./exchangedisconnect.ps1 all



# if ran with any arguments, disconnect all
if (-not ($args.count -eq 0))
{
	get-pssession | remove-pssession
	echo "disconnecting all"
	echo " "
}
# else check that it's safe to disconnect all
else
{
	# check if there are connections 
	if ($null -ne (get-pssession))
	{
		# if there are connections that aren't Office 365 then abort and tell user to do it manually
		if($null -ne (get-pssession | where-object {$_.ComputerName -ne 'outlook.office365.com'}))
		{
			echo 'a non-office 365 connection is detected, aborting'
			echo 'to disconnect all run    ./exchangedisconnect.ps1 all'
			echo " "
		}
		# else the only connection is to Office 365, so disconnect it
		else
		{
			get-pssession | remove-pssession
			echo 'Office 365 disconnecting'
			echo " "
		}
	}
	else
	{
		echo "no connections found"
	}
}

echo "connections that remain:"
get-pssession
