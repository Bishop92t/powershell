# created by Alan Bishop
# last modified 9/11/2020
#
# Watches a particular folder, if a new file appears: send it via email, then move the file to a different location
#
# Usage: 
# 	.\Watch-Directory.ps1 ap                run with the ap flag
#    .\Watch-Directory.ps1 cc 			run with the cc flag
# 	.\Watch-Directory.ps1  				if no valid flag is provided use the test flag
#
#
# Dependencies: 
# Move-ReceiptFiles.ps1  Send-Email.ps1  tpath.enc  &  several files in  .\debloat files\


# setup a file watcher object
$watcher = New-Object System.IO.FileSystemWatcher

# determine which folder to watch
if ($args[0] -eq "ap")
{
	# pathapsource.txt : the path to be watched for changes, example: c:\windows\
	$watcher.Path = Get-Content -Path '.\debloat files\pathapsource.txt'
}
elseif ($args[0] -eq "cc")
{
	# pathccsource.txt : the path to be watched for changes, example: c:\windows\
	$watcher.Path = Get-Content -Path '.\debloat files\pathccsource.txt'
}
else
{
	$watcher.Path = "d:\temp"
}


# setup the action to take when a new file is detected
$action = 
{ 
	# not being used currently, leaving in for future use
	# $newFile     = $Event.SourceEventArgs.FullPath
	
	# pull the message data to pass to Move-ReceiptFiles.ps1
	$pathSelection = $Event.MessageData.pathSelection

	Write-Output "New $pathSelection file found. $(Get-Date)" >> c:\logs\watch-directory.txt

	# call the PoSh script
	& ".\Move-ReceiptFiles.ps1" $pathSelection
}

# setup some params to pass to the action object
$actionParams = New-Object psobject -property @{pathSelection = $args[0]}

# register the file watcher event
Register-ObjectEvent -InputObject $watcher -EventName Created -Action $action -MessageData $actionParams

# to keep this session open just loop infinitely
do
{
	start-sleep -seconds 3
}
while ($true)