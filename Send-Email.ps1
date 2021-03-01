# created by Alan Bishop
# last modified 9/10/2020
#
# Enables easy notification by email, intended to be used by other scripts
#
# usage:
# 			.\Send-Email.ps1 $recipient $subject $text
# 			.\Send-Email.ps1 jsmith@hospicebr.org "warning email" "Good morning John, this is a warning email"
#
# 			.\Send-Email.ps1 $recipient $subject $text $fileToAttachLocation
# 			.\Send-Email.ps1 jsmith@hospicebr.org "warning email" "Good morning John, this is a warning email" "C:\logs\log.txt"


$logfile = "c:\logs\send-email.txt"
Write-Output "Starting Send-Email  $(Get-Date) " >> $logfile

# for recipient, subject and body of text
if ($args.count -eq 3) 
{
	$recipient = $args[0]
	$subject   = $args[1]
	$text      = $args[2]

	# setup the email credentials
	$username    = (Get-Content -Path '.\debloat files\alertemail.txt')+(Get-Content -Path '.\debloat files\email.txt')
	$tpath       = Get-Content -Path '.\debloat files\alertemail.txt'
	$token       = ConvertTo-SecureString (Get-Content ( (Get-Content -Path '.\debloat files\alertemail2.txt')+"\$tpath.enc") )
	$credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $token

	# send out the email
	Send-MailMessage -To $recipient -From $username -SmtpServer "smtp.office365.com" -UseSsl -Subject $subject -Body $text -port 587 -Credential $credentials 

	Add-Content $logfile ("email sent to $recipient $(Get-Date) `n")
}
# for recipient, subject, text body, and a file to attach
elseif ($args.count -eq 4)
{
	$recipient = $args[0]
	$subject   = $args[1]
	$text      = $args[2]
	$fileLoc   = $args[3]

	# get the file size in MB, if too large, log and send the email without the attachement (and inform where to find the file)
	$fileSize   = ((get-item $fileLoc).length/1mb)
	$fileTooBig = $false
	if ($fileSize -gt 30)
	{
		Add-Content $logfile ("File too big, $fileSize was attempted, the limit is 30 "+(Get-Date -Format "MM-dd-yyyy")+" `n")
		$shareLoc = $fileLoc.Substring(3,5)
		Switch ($shareLoc)
		{
			"bi" {$shareLoc = "b"}
			"co" {$shareLoc = "f"}
			"ex" {$shareLoc = "g"}
			"fi" {$shareLoc = "m"}
			"it" {$shareLoc = "i"}
			"pa" {$shareLoc = "n"}
			default {$shareLoc = "unknown"}
		}
		$text       = "$text `n`n The attached file is too big to be emailed. You can find it on the $shareLoc drive"
		$fileTooBig = $true
	}

	# setup the email credentials
	$username    = (Get-Content -Path '.\debloat files\alertemail.txt')+(Get-Content -Path '.\debloat files\email.txt')
	$tpath       = Get-Content -Path '.\debloat files\alertemail.txt'
	$token       = ConvertTo-SecureString (Get-Content ( (Get-Content -Path '.\debloat files\alertemail2.txt')+"\$tpath.enc") )
	$credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $token

	# send out the email, including the attachment if it's small enough
	if ($fileTooBig)
	{
		Send-MailMessage -To $recipient -From $username -SmtpServer "smtp.office365.com" -UseSsl -Subject $subject -Body $text -port 587 -Credential $credentials 
	}
	else
	{
		Send-MailMessage -To $recipient -From $username -SmtpServer "smtp.office365.com" -UseSsl -Subject $subject -Body $text -port 587 -Credential $credentials -Attachments $fileLoc
	}

	Add-Content $logfile ("email sent to $recipient "+(Get-Date)+" `n")	
}
else
{
	Add-Content $logfile ("Incorrect number of args passed. Should be 3 or 4, received $($args.count)  "+(Get-Date -Format "MM-dd-yyyy")+" `n ")
}
