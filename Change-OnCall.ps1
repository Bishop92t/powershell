# created by Alan Bishop
# last updated 2/28/2021
#
# Each week a new person takes after hours calls, and they need to be sent eFax emails. New person gets
# added and last weeks person gets removed. Some users get the emails all the time so they are excluded.
# This is nearly 100% automated, however oncallschedule.txt must be manually updated.
#
# This script is intended to be auto-run weekly as a scheduled task
# It relies on exchangeconnect.ps1, exchangedisconnect.ps1, faxadd.ps1 and faxremove.ps1 to be in c:\script


# start a detailed log file of this script for troubleshooting, beginning with userlist before, verbose log, then userlist after
$verboseLog = "c:\logs\verbose-change-oncall.txt"
Start-Transcript -path $verboseLog -Append


# these have full time fax emails anyway, don't process
# oncallexcluded.txt : list of SamAccountName's, each on a new line
$excludedUsers = Get-Content "c:\script\debloat files\oncallexcluded.txt"


# setup in and out streams and log file
$inFile    = "c:\script\debloat files\oncallschedule.txt"
$inStream  = New-Object System.IO.StreamReader ($inFile)
$outFile   = "c:\script\debloat files\temp.txt"
$outStream = New-Object System.IO.StreamWriter ($outFile)
$logFile   = "c:\logs\change-oncall.txt"

# initialize the array list that holds the output stream
[System.Collections.ArrayList]$outArrayList = @("")  
$outArrayList.Remove("") 

# connect to the Office365 server (and flag that this script did the connection)
if($null -eq (get-pssession | where-object {$_.ComputerName -EQ 'outlook.office365.com'}))
{
	c:\script\exchangeconnect.ps1 | Out-Null
	$connected = $true
}

# once connected to Office365, pull the user list before this script runs
$userListBefore = .\faxadd.ps1


#####################################################################
# this section removes last weeks on call (and removes user from list)

# if the user is no longer on call and not in the full time list, remove their access
$userNoLongerOnCall = $inStream.ReadLine()
if (-not ($excludedUsers -match $userNoLongerOnCall))
{
	& c:\script\faxremove.ps1 $userNoLongerOnCall
	Add-Content $logFile ("Removed $userNoLongerOnCall")
}
else
{
	Add-Content $logFile ("$userNoLongerOnCall not removed, full time access")
}

#####################################################################
# this section copies the rest of the instream to an arraylist, then closes instream

# read rest of the input file to an arraylist
while (!$inStream.EndOfStream)
{
	$outArrayList.Add($inStream.ReadLine()) | Out-Null
}
$inStream.Close()
$inStream.Dispose()


#####################################################################
# this section adds the next weeks on call

# if the next oncall person doesn't always have access, give them access
$userOnCall = $outArrayList[0]
echo $userOnCall

# if line is blank there's no more file to read, email out an alert to the SA
if ($userOnCall -eq $null)
{
	# email.txt : the email domain in the format  @$company.com
	$alertEmail = "nbishop"+(Get-Content "c:\script\debloat files\email.txt")
	c:\script\Send-Email.ps1 $alertEmail "Change-OnCall.ps1 alert" "Ran out of valid emails, check oncallschedule.txt"
}
# if the user isn't on the exclusion list then add them to list and notify them
elseif (-not ($excludedUsers -match $userOnCall))
{
	& c:\script\faxadd.ps1 $userOnCall
	Add-Content $logFile ("Added $userOnCall")
	# email.txt : the email domain in the format  @$company.com
	$emailAddress = $userOnCall+(Get-Content "c:\script\debloat files\email.txt")
	$subject = "You've been added to manager on call eFax"
	$message = "You've been added to the manager on call eFax. If you are not manager on call this week please notify Alan so he can remove you. This is an automated email."
	& ".\Send-Email.ps1" $emailAddress $subject $message
}
# else the user already has fulltime access, so do nothing
else
{
	Add-Content $logFile ("$userOnCall already has full time access")
}

#####################################################################
# this section dumps the arraylist into the outstream, then copies over the original inputstream file

# dump arraylist to outstream
foreach ($line in $outArrayList)
{
	$outStream.WriteLine($line)
}

# close outstream
$outStream.Close()
$outStream.Dispose()

# for the verbose log, pull userlist before we close the connection
$userListAfter = .\faxadd.ps1

# if this script connected to Exchange, than disconnect
if ($connected)
{
	c:\script\exchangedisconnect.ps1
}

# attempt to gracefully close the files and date the log file entry
Remove-Item $inFile -Force
Rename-Item $outFile $inFile
Add-Content $logFile ("Completed "+(Get-Date -Format "MM-dd-yyyy")+" `n ")

# finish off the verbose log
Stop-Transcript | Out-Null

Add-Content $verboseLog ("User list before `n $($userListBefore)")
Add-Content $verboseLog ("User list after `n $($userListAfter)")
