# Alan Bishop 
# last updated 3/28/2021
#
#  Looks in all files in current directory and displays which files have the provided text in them


if ($args.count -eq 1)
{
	Get-ChildItem -Recurse | Select-String $args[0] -List | Select Path
}
elseif ($args.count -eq 2)
{
	Get-ChildItem -Recurse | Select-String $args[0] -List -Include $args[1] | Select Path
}
else
{
	echo "usage:"
	echo "	.\find.ps1 $text    		    Looks in all files in current dir and displays which have $text located in them"
	echo "	.\find.ps1 $text $extension	Looks in files with $extension and displays which have $text located in them"
	echo "example:"
	echo "	.\find.ps1 nbishop *.ps1 	  Looks for nbishop in all ps1 files"
}
