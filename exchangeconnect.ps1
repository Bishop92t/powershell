# Alan Bishop 
# modified 5/13/2021
#
# connects this session to Exchange Online, intended to be a helper script but can be run solo
#
# the following must be installed first:
#      .\required addons\msoidcli_64.msi  (MS Online Services Sign-In Assistant)
#      install-module -name AzureAD
#      install-module MSOnline
# new version requires:
#	   install-module -name exchangeonlinemanagement

# the token we'll be working with, based on current logged in user
$token = "$($env:userprofile)\$($env:username).enc"

# if the token is saved
if(test-path $token)
{
	# load user credentials
	$name = Get-Content -Path '.\debloat files\email.txt'
	$name = $env:username+$name
	$pass = Get-Content $token | ConvertTo-SecureString
	$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $name, $pass

	# connect to Exchange Online and Azure AD
	Connect-ExchangeOnline -Credential $cred
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
	$userSwap = Get-Content ".\debloat files\exchangeswap.txt"
	if ($env:username -eq $userSwap[0])
	{
		$username = $userSwap[1]
	}

	# create a token then try to run this script again
	.\tokenpassword.ps1 "gui" $env:username "Verify username is an administrator account for connecting to email server"
	.\exchangeconnect.ps1
}


echo "current active sessions: "
get-pssession
