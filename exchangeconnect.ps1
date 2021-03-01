# Alan Bishop 
# modified 8/3/2020
#
# connects this session to Exchange Online, intended to be a helper script but can be run solo
#
# the following must be installed first:
#      .\required addons\msoidcli_64.msi  (MS Online Services Sign-In Assistant)
#      install-module -name AzureAD
#      install-module MSOnline
# new version requires:
#	   install-module -name exchangeonlinemanagement


# if the token is saved
if(test-path $env:userprofile\token.enc)
{
	# load user credentials
	$name = Get-Content -Path '.\debloat files\email.txt'
	$name = $env:username+$name
	$pass = Get-Content $env:userprofile\token.enc | ConvertTo-SecureString
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
# if token not saved, inform the user
else
{
	echo "first you need to run ./tokenpassword.ps1  "
	echo "to store your password encrypted in your user profile"
}


echo "current active sessions: "
get-pssession
