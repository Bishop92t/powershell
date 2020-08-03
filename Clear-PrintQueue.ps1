# created by Alan Bishop
# last modified 8/3/2020
#
# Clears all print jobs from the selected printers, then restarts the spooler service
# This is intended to be run at night as a scheduled task, or run on demand if necessary
# Just be aware it will wipe the queue and restart spooler on the print server
#
# Must be run from the print server itself!



# printers.txt : list of each printer to check for stuck jobs, each on a new line
$printers = Get-Content -Path '.\debloat files\printers.txt'

# log file and change counter setup
$logFile    = "c:\logs\clear-printer.txt"
$numChanges = 0

foreach ($printer in $printers)
{
	# if there aren't any print jobs just write that to the log file
	if ((Get-PrintJob $printer) -eq $null)
	{
		# do nothing
	}
	# otherwise write to log that we cleared all print jobs and mark that changes were made
	else
	{
		Add-Content $logFile ("print job CLEARED for $printer")
		Get-PrintJob $printer | Remove-PrintJob
		$numChanges++
	}
}

# if there were any changes made, restart the print spooler
if ($numChanges -gt 0)
{
	Add-Content $logFile ("print spooler service restarted")
	Restart-Service -DisplayName "Print Spooler" -force
}
else
{
	Add-Content $logFile ("no changes made")
}

# finally log the date at which this script ran
Add-Content $logFile ("Completed "+(Get-Date -Format "MM-dd-yyyy")+" `n `n ")