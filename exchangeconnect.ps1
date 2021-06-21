# Alan Bishop 
# modified 6/21/2021
#
# connects this PS session to Exchange Online, intended to be a helper script but can be run solo
#
# the following must be installed first (script will attempt to install automatically):
#      .\required addons\msoidcli_64.msi  (MS Online Services Sign-In Assistant)
#      install-module -name AzureAD
#      install-module MSOnline
# new version requires:
#	   install-module -name ExchangeOnlineManagement


# check to see where the scripts are being run from
if (Test-Path "c:\script\Give-EmailAccess.ps1")
{
	$sPath = "c:\script\"
}
elseif (Test-Path "c:\script\ps1\Give-EmailAccess.ps1")
{
	$sPath = "c:\script\ps1"
}
else
{
	Write-Output "difficulty locating PS1 path"
}

# setup some var's to check if the required modules are installed
$checkAzureAD 				 = (Get-Module AzureAD -ListAvailable).Name
$checkMSOnline 			 = (Get-Module MSOnline -ListAvailable).Name
$checkExchangeOnlineManagement = (Get-Module ExchangeOnlineManagement -ListAvailable).Name

# if AzureAD module isn't installed, install it
if ($checkAzureAD -eq $null)
{
	Write-Host "Important: AzureAD Powershell module must be installed to continue "
	Install-Module AzureAD -Repository PSGallery -AllowClobber -Force
}
# if MSOnline module isn't installed, install it
if ($checkMSOnline -eq $null)
{
	Write-Host "Important: MSOnline Powershell module must be installed to continue "
	Install-Module MSOnline -Repository PSGallery -AllowClobber -Force
}
# if ExchangeOnlineManagement module isn't installed, install it
if ($checkExchangeOnlineManagement -eq $null)
{
	Write-Host "Important: ExchangeOnlineManagement Powershell module must be installed to continue "
	Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
}


# the token we'll be working with, based on current logged in user
$token = "$($env:userprofile)\$($env:username).enc"

# if the token is saved
if(test-path $token)
{
	Write-Output "token found $($env:username)"
	
	# load user credentials
	$SAM = $env:username
	# special cases for users without email account on elevated accounts
	if (($SAM -eq "adminnia") -or ($SAM -eq "atest"))
	{
		$name = Get-Content -Path 'c:\script\debloat files\emailna.txt'
		$name = "$($SAM)$($name)"
	}
	# if not special case, use default domain
	else
	{
		$name = Get-Content -Path 'c:\script\debloat files\email.txt'
		$name = "$($SAM)$($name)"
	}

	# pull credentials from token file
	$pass = Get-Content $token | ConvertTo-SecureString
	$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $name, $pass

	# connect to Exchange Online and Azure AD
	# considering separating these out to speed up run time, Connect-ExchangeOnline takes 7.87 seconds
	Connect-ExchangeOnline -Credential $cred
	# Connect-MsolServer takes 1.75 seconds   (  both tested using Measure-Command {command}  )
	Connect-MsolService -Credential $cred

	# old connection style, will remove once fully deprecated
	# create a new session to Office 365 Exchange
	# Import-Module MSOnline
	# $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication basic -AllowRedirection
	# import that into current PS session, waiting for it to finish before proceeding
	# Import-PSSession $Session -DisableNameChecking | Out-Null
}
# if token not saved, prompt to create one
else
{
	# handling a special case
	$userSwap = Get-Content "c:\script\debloat files\exchangeswap.txt"
	if ($env:username -eq $userSwap[0])
	{
		$username = $userSwap[1]
	}

	$tPath = "$($sPath)tokenpassword.ps1" 
	& $tPath gui $env:username 'Verify username is an administrator account for connecting to email server'
	& "$($sPath)exchangeconnect.ps1"
}


Write-Output "current active sessions: "
Get-PSSession
