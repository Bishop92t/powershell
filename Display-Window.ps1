# created by Alan Bishop
# last modified 5/20/2021
#
# Creates a basic GUI window with a message passed in args. Can also customize the window size
#
# Usage:
# 		.\Display-Window.ps1 $message 				display a message with a default 1200x700 window
# 		.\Display-Window.ps1 $message $width $height 	display a message with a window of size $width x $height

Function generateForm
{
	param (	[Parameter(Mandatory=$true, Position=0)] $message, 
			[Parameter(Position=1)] $width, 
			[Parameter(Position=2)] $height 
		 )

	# if 3 args passed, set window width and height, otherwise use default	
	if ($width -eq $null)
	{
		$width  = 1200
		$height = 700
	}
	
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	$form = New-Object system.Windows.Forms.Form

	$form.ClientSize = "$($width),$($height)"
	$form.BackColor  = '#000022'
	$form.ForeColor  = '#ffffff'
	$form.text       = 'Windows Debloat'

	$roleTypeLabel          = New-Object system.Windows.Forms.Label
	$roleTypeLabel.text     = $message
	$roleTypeLabel.font     = 'Lucida Console,14'
	$roleTypeLabel.autosize = $true
	$roleTypeLabel.location = New-Object System.Drawing.Point(10,10)

	$form.controls.AddRange(@($roleTypeLabel,$radioButton1,$okButton,$computerNameLabel,$computerNameTextBox))
	[void]$form.ShowDialog()
}

if ($args.count -eq 1)
{
	generateForm $args[0]
}
elseif ($args.count -eq 3)
{
	generateForm $args[0] $args[1] $args[2]
}
else
{
	[System.Windows.Forms.MessageBox]::Show("Usage: `n Display-Window message `n Display-Window message width height", "Incorrect # of Args")  
}
