# Created by Alan Bishop
# Last modified 8/31/2020
#
# Schedules a task to auto run the bginfo 

$action  = New-ScheduledTaskAction -Execute "c:\bginfo\bginfo64.exe" -Argument "/iq c:\bginfo\thbr.bgi /timer:0 /nolicprompt"
$trigger = @(	$(New-ScheduledTaskTrigger -AtLogon),
			$(New-ScheduledTaskTrigger -Daily -At 11pm)  )
Register-ScheduledTask -TaskName "bginfo" -Action $action -Trigger $trigger -User "system"

