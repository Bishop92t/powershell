# Alan Bishop 4/22/2021
#
# PowerShell script to create Split Tunnel SSTP VPN connections
#
# Sometimes the routes are lost, if VPN not working double check routes exist with the following powershell:
# 	(Get-VpnConnection -Name $VPNConnectionName).routes


# VPN Name that will be displayed to the user
$VPNConnectionName     = ""
# the actual VPN address accessible from the internet
$VPNServerAddress      = ""
# the internal domain name
$DNSSuffix             = ""

try 
{
	# if connection not found, add it
	if (-not ((Get-VpnConnection).Name -eq $VPNConnectionName))
	{
		# Creating VPN of type $VPNType named $VPNConnectionName... 
		Add-VpnConnection -Name $VPNConnectionName -ServerAddress $VPNServerAddress -TunnelType "Sstp" -EncryptionLevel "Required" -AuthenticationMethod MsChapv2 -SplitTunneling -PassThru 

		# Adding RFC1819 routes to VPN $VPNConnectionName...
		# Any internal network IP's that need to be accessed, anything not in this list will attempt to be accessed externally 
		# example:  $DNSPrefix1 = "192.168.0.1/8"
		Add-VpnConnectionRoute -ConnectionName $VPNConnectionName -DestinationPrefix $DNSPrefix1 -PassThru

		# Setting DNS suffix and use login credentials
		Set-VPNConnection -Name "$VPNConnectionName" -DNSSuffix $DNSSuffix -UseWinlogonCredential $true -RememberCredential:$true 
	}
}
catch 
{
	Write-Host "`n An error occurred: $_ `n"
	pause
}
