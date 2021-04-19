# Created by Nathanael Bishop
#
# Retrieves basic information of a PC.
#
# last modified: 4/19/2021

# get PS ready to display a form
Add-Type -AssemblyName System.Windows.Forms

# get the computer information
$Computer = $env:computername
$ComputerVer = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name CurrentBuild).CurrentBuild
$UBR = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name UBR).UBR
$ComputerCPU = Get-WmiObject win32_processor -ComputerName $Computer | select DeviceID,Name | FT -AutoSize  | Out-String
$ComputerRAM = Get-WmiObject Win32_PhysicalMemory -ComputerName $Computer | select DeviceLocator,Manufacturer,PartNumber,Capacity,Speed | FT -AutoSize | Out-String
$ComputerDisks = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3" -ComputerName $Computer | select DeviceID,VolumeName,Size,FreeSpace | FT -AutoSize | Out-String

# format the PC info into readable form
$text = 	"Computer Name: $($Computer) `n" +
		"Operating System: $($ComputerVer).$($UBR) `n"+
		"$($ComputerCPU) "+
		"$($ComputerRAM)"+
		"$($ComputerDisks)" 

# start creating the form
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$form = New-Object system.Windows.Forms.Form
$form.ClientSize = '800,450'
$form.BackColor  = '#000000'
$form.ForeColor  = '#ffffff'
$form.text       = "Get-PCInfo"

# setup the PC info text box
$textToDisplay          = New-Object system.Windows.Forms.Label
$textToDisplay.text     = $text
$textToDisplay.AutoSize = $true
$textToDisplay.location = New-Object System.Drawing.Point(10,10)
# best to use a unispace font like Courier or Terminal or tables won't display correctly
$textToDisplay.Font     = 'Courier New,12'

$form.controls.AddRange(@($textToDisplay))
[void]$form.ShowDialog()
