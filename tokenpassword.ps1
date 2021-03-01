# Alan Bishop 
# last updated 8/24/20
#
# Asks for and stores password token encrypted on local drive under user profile
#
# ********************************************************************************
# *      NEVER copy this token to a USB drive, laptop or onto the internet       *
# ********************************************************************************
#
# usage:
# 		.\tokenpassword.ps1 				      pops up a dialog asking for current admins password, creates token of that password
# 		.\tokenpassword.ps1 $pass 			  creates token using the password provided, then clears screen
# 		.\tokenpassword.ps1 $user $pass 	creates a token for $user based on the password provided, then clears the screen.
# 										                	note that $user must be the users SAM name (eg John Smith's SAM would be jsmith)

# if an argument is passed, use as password
if ($args.count -eq 1) 
{
	ConvertTo-SecureString -string $args[0] -asplaintext -force | ConvertFrom-SecureString | out-file $env:userprofile\token.enc
	$tokenfile = "$($env:username).enc"
}

# else prompt the user for their password
elseif ($args.count -eq 0)
{
	read-host -prompt "<password>" -AsSecureString | ConvertFrom-SecureString | out-file $env:userprofile\token.enc
	$tokenfile = "$($env:username).enc"
}

# else use 
elseif ($args.count -eq 2)
{
	$tokenfile = $args[0] + ".enc"
	ConvertTo-SecureString -string $args[1] -asplaintext -force | ConvertFrom-SecureString | out-file $env:userprofile\$tokenfile
}

# for arguments > 3, should be rare
else
{
$tokenfile = "incorrect number of arguments"
}

# clear screen and indicate success
cls
echo "token file saved: $tokenfile"
