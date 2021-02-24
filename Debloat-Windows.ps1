# created by Alan Bishop 
#
# *********************************************************************************************************************************
#
# GUI version of debloat.ps1 
#   Installs software, renames computer, joins domain, removes Win10 bloat, add VPN/WiFi connections, uses a standard setting theme
#   Intended to be run multiple times after multiple reboots
#
# Major actions:
#   Run 1: run as temp, renames computer
#   Run 2: run as temp, domain joins
#   Run 3: run as domain admin
#   Run 4: run as end user 
#
# *********************************************************************************************************************************
#
# changelog:
#   2/24/21  bug fixes on menu trimming
#   2/14/21  removed options 2-4, added network connection safety tester, create auto restart scheduled task
#   11/13/20 fix auto login bug
#   11/2/20  UAC fix enabled
#   9/18/20  bug fixes: duplicate variable names, typo, turn off screen res change
#   8/24/20  auto connect to WiFi fixed
#   8/6/20   gearing up to break this beast into more manageable modules
#   6/16/20  switched to using the USB drive letter instead of assuming PS would figure it out. Used to work but perhaps a Windows/PS update broke it



# setup helper function to test if PS is run as administrator
function Test-Administrator
{  
    $user = [Security.Principal.WindowsIdentity]::GetCurrent();
    (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)  
}

Function generateForm
{
    # logged in as temp and no flag file : option 1 (continue with form)
    # logged in as temp and flag file : option 2 (skip form)
    # logged in not as temp, but admin rights : option 3 (skip form)
    # logged in not as temp, and not admin : option 4 (skip form)
    if ($env:username -ne "temp")
    {   
        if (Test-Administrator)
        {
            debloatWindows 3 "none"
        }
        else
        {
            debloatWindows 4 "none"
        }
        exit
    }
    elseif ($env:username -eq "temp")
    {
        if (test-path "$env:userprofile\desktop\Debloat-Windows.ps1")
        {
            debloatWindows 2 "none"
            exit
        }
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $form = New-Object system.Windows.Forms.Form

    $form.ClientSize = '1200,700'
    $form.BackColor  = '#000022'
    $form.ForeColor  = '#ffffff'
    $form.text       = 'Windows Debloat'

    $roleTypeLabel          = New-Object system.Windows.Forms.Label
    $roleTypeLabel.text     = "How many times has this script been run on this machine:"
    $roleTypeLabel.font     = 'Microsoft Sans Serif,14'
    $roleTypeLabel.autosize = $true
    $roleTypeLabel.location = New-Object System.Drawing.Point(10,10)

    # $radioButton1 is only for the very first run - renames computer
    $radioButton1          = New-Object system.Windows.Forms.RadioButton
    $radioButton1.Location = New-Object System.Drawing.Size(10, 45)
    $radioButton1.Size     = New-Object System.Drawing.Size(400,30)
    $radioButton1.Checked  = $true
    $radioButton1.Font     = 'Microsoft Sans Serif,12'
    $radioButton1.text     = "1st: logged in as temp user (run as admin)"
    $form.controls.Add($radioButton1)

    # setup the label that describes the text box purpose
    $computerNameLabel          = New-Object system.Windows.Forms.Label
    $computerNameLabel.text     = "computers name is required:"
    $computerNameLabel.Font     = 'Microsoft Sans Serif,12'
    $computerNameLabel.AutoSize = $true
    $computerNameLabel.Visible  = $false
    $computerNameLabel.location = New-Object System.Drawing.Point(30,75)

    # setup the computer name text box
    $computerNameTextBox          = New-Object System.Windows.Forms.TextBox
    $computerNameTextBox.Size     = New-Object System.Drawing.Size(260,20)
    $computerNameTextBox.Location = New-Object System.Drawing.Point(30,100)
    $computerNameTextBox.Visible  = $false
    $form.Controls.Add($computerNameTextBox)

    # $radioButton2 is for 2nd run as temp (to domain join), and any other run as Admin where software is needed to be installed
    $radioButton2          = New-Object system.Windows.Forms.RadioButton
    $radioButton2.Location = New-Object System.Drawing.Size(10,145)
    $radioButton2.Size     = New-Object System.Drawing.Size(500,50)
    $radioButton2.Checked  = $false
    $radioButton2.Font     = 'Microsoft Sans Serif,12'
    $radioButton2.text     = "2nd: logged in as temp user (run as admin)`n 3rd: logged in as domain\YOURLOGON (run as admin)"
    $form.controls.Add($radioButton2)

    # $radioButton3 is for debloat only - no domain join, no computer rename, no software installed
    $radioButton3          = New-Object system.Windows.Forms.RadioButton
    $radioButton3.Location = New-Object System.Drawing.Size(10, 200)
    $radioButton3.Size     = New-Object System.Drawing.Size(500,50)
    $radioButton3.Checked  = $false
    $radioButton3.Font     = 'Microsoft Sans Serif,12'
    $radioButton3.text     = "4th: logged in as any end user (do not run as admin) `n You can also use this to debloat existing machines"
    $form.controls.Add($radioButton3)


    # setup the create user button
    $okButton           = New-Object system.Windows.Forms.Button
    $okButton.text      = "missing info"
    $okButton.width     = 90
    $okButton.height    = 30
    $okButton.location  = New-Object System.Drawing.Point(480,300)
    $okButton.Font      = 'Microsoft Sans Serif,10'
    $okButton.ForeColor = "#000000"
    $okButton.BackColor = "#dddddd"
    $okButton.enabled   = $false
    $okButton.Visible   = $true

    $form.AcceptButton = $okButton

    # Compile and display the form
    $form.controls.AddRange(@($roleTypeLabel,$radioButton1,$okButton,$computerNameLabel,$computerNameTextBox))

    # if radioButton1 is checked, computer name is required (disable button if no text)
    $radioButton1.Add_Click({
        if ($computerNameTextBox.Text.Length -eq 0)
        {
            $okButton.Enabled = $false
            $okButton.text    = "missing info"
        }
        if (-Not $computerNameLabel.Visible)
        {
            $computerNameLabel.Visible   = $true
            $computerNameTextBox.Visible = $true
        }
    })

    # if radioButton2 is checked, computer name isn't required so hide it (and enable button)
    $radioButton2.Add_Click({
        $okButton.Enabled = $true
        $okButton.text    = "Ok"
        if ($computerNameLabel.Visible)
        {
            $computerNameLabel.Visible   = $false
            $computerNameTextBox.Visible = $false
        }
    })

    # if radioButton3 is checked, computer name isn't required so hide it (and enable button)
    $radioButton3.Add_Click({
        $okButton.Enabled = $true
        $okButton.text    = "Ok"
        if ($computerNameLabel.Visible)
        {
            $computerNameLabel.Visible   = $false
            $computerNameTextBox.Visible = $false
        }
    })

    # don't allow $okButton to be pressed if text hasn't been filled out (exception for option 2 handled above)
    $computerNameTextBox.add_TextChanged({
        if ($computerNameTextBox.Text.Length -ne 0)
        {
            $okButton.Enabled = $true
            $okButton.text    = "Ok"
        }
        else
        {
            $okButton.Enabled = $false
            $okButton.text    = "missing info"
        }
    })


    # if $okButton is clicked, run debloatWindows function
    $okButton.Add_Click({ 
        if ($radioButton1.Checked -eq $true)
        {
            debloatWindows 1 $computerNameTextBox.Text
        }
        elseif ($radioButton2.Checked -eq $true)
        {
            debloatWindows 2 "none"
        }
        else
        {
            debloatWindows 3 "none"
        }
        $form.Close()
    })

    [void]$form.ShowDialog()
}



function debloatWindows
{
    param (
        [Parameter(Mandatory=$true, Position=0)] [int] $runOption, 
        [Parameter(Mandatory=$true, Position=1)] [string] $computerName
    )


    # start a detailed log file of this script for troubleshooting
    try 
    {
        Stop-Transcript | out-null
    }
    catch [System.InvalidOperationException]{}

    Start-Transcript -path c:\logs\verbose.txt

    #
    # $runOption=1 set computer name, install software
    # $runOption=2 install software, joins domain if run as temp, deletes temp if run as domain admin 
    # $runOption=3 don't install software, just debloats
    #

    # list of apps that are going to be removed
    $apps = @("07AF453C.IndexCards", "2414FC7A.Viber", "2FE3CB00.PicsArt-PhotoStudio", "41038Axilesoft.ACGMediaPlayer", "46928bounde.EclipseManager", "4DF9E0F8.Netflix", "64885BlueEdge.OneCalendar", "6Wunderkinder.Wunderlist", "7458BE2C.WorldofTanksBlitz", "7EE7776C.LinkedInforWindows", "8075Queenloft.BlendCollagePhotoEditor", "828B5831.HiddenCityMysteryofShadows", "89006A2E.AutodeskSketchBook", "9E2F88E3.Twitter", "A278AB0D.DisneyMagicKingdoms", "A278AB0D.DragonManiaLegends", "A278AB0D.MarchofEmpires", "ActiproSoftwareLLC.562882FEEB491", "AdobeSystemsIncorporated.AdobePhotoshopExpress", "CAF9E577.Plex", "ClearChannelRadioDigital.iHeartRadio", "D52A8D61.FarmVille2CountryEscape", "D5EA27B7.Duolingo-LearnLanguagesforFree", "DB6EA5DB.CyberLinkMediaSuiteEssentials", "DolbyLaboratories.DolbyAccess", "Drawboard.DrawboardPDF", "Facebook.317180B0BB486", "Facebook.Facebook", "Facebook.InstagramBeta", "flaregamesGmbH.RoyalRevolt2", "Flipboard.Flipboard", "GAMELOFTSA.Asphalt8Airborne", "KeeperSecurityInc.Keeper", "king.com.BubbleWitch3Saga", "king.com.CandyCrushSodaSaga", "Microsoft.3DBuilder", "Microsoft.Advertising.Xaml", "Microsoft.Advertising.Xaml", "Microsoft.AgeCastles","Microsoft.AppConnector","Microsoft.BingFinance","Microsoft.BingFoodAndDrink", "Microsoft.BingHealthAndFitness", "Microsoft.BingNews", "Microsoft.BingSports", "Microsoft.BingWeather", "Microsoft.CommsPhone", "Microsoft.ConnectivityStore", "Microsoft.FreshPaint", "Microsoft.Getstarted", "Microsoft.Messaging", "Microsoft.Microsoft3DViewer", "Microsoft.MicrosoftMahjong", "Microsoft.MicrosoftOfficeHub", "Microsoft.MicrosoftPowerBIForWindows", "Microsoft.MicrosoftSolitaireCollection", "Microsoft.MicrosoftSudoku", "Microsoft.MinecraftUWP", "Microsoft.NetworkSpeedTest", "Microsoft.Office.OneNote", "Microsoft.Office.Sway", "Microsoft.OneConnect", "Microsoft.People", "Microsoft.SkypeApp", "Microsoft.Windows.FeatureOnDemand.InsiderHub", "Microsoft.WindowsAlarms", "Microsoft.WindowsFeedbackHub", "Microsoft.WindowsMaps", "Microsoft.WindowsPhone", "Microsoft.WindowsSoundRecorder", "Microsoft.ZuneMusic", "Microsoft.ZuneVideo", "PandoraMediaInc.29680B314EFC2", "Playtika.CaesarsSlotsFreeCasino", "Psykosoft.Psykopaint", "ShazamEntertainmentLtd.Shazam", "TheNewYorkTimes.NYTCrossword", "TuneIn.TuneInRadio", "XINGAG.XING", "microsoft.windowscommunicationsapps")

    # log file setup
    if(-not (Test-Path 'c:\logs\'))
    {
        New-Item -Path "c:\" -Name "logs" -ItemType "directory"
    }
    $logFile = "c:\logs\debloatwindows.txt"
    $date = Get-Date -Format "MM-dd-yyyy"
    Add-Content $logFile ("`n `n Starting on $date `n")

    # determines where the script is being run from, some pc's seem to struggle with this (PS ver issue?)
    $drives = Get-WmiObject Win32_Volume -Filter "DriveType='2'" | select -expand driveletter
    foreach ($drive in $drives)
    {   
        if (test-path "$($drive)\Debloat-Windows.bat")
        {
            $installdrive = $drive
        }
    }

    # close the option window, start progress box, and set it to center on page
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    # this is out of scope, shouldn't be necessary anymore
    # $form.Dispose()
    $form2  = New-Object system.Windows.Forms.Form
    $center = [System.Windows.Forms.FormStartPosition]::CenterScreen;

    # setup progress box
    $form2.ClientSize = '1200,700'
    $form2.BackColor  = '#000022'
    $form2.ForeColor  = '#ffffff'
    $form2.text       = 'Windows Debloat'
    $form2.StartPosition = $center;

    # start progress box text
    $roleTypeLabel          = New-Object system.Windows.Forms.Label
    $roleTypeLabel.text     = "Give each install 10+ minutes"
    $roleTypeLabel.font     = 'Microsoft Sans Serif,16'
    $roleTypeLabel.autosize = $true
    $roleTypeLabel.location = New-Object System.Drawing.Point(10,10)    
    
    # display progress box, enable refresh when text changes
    $form2.controls.AddRange(@($roleTypeLabel))
    $roleTypeLabel.add_textchanged{$form2.refresh()}
    [void]$form2.Show()

    # pull the network location and connect to WIFI
    # network.txt : a shared network folder, ex:  \\server\directory\
    $netpath = Get-Content -Path "$($installdrive)\debloat files\network.txt"
    netsh wlan add profile filename="$($installdrive)\debloat files\Wi-Fi-THBR.xml"

    # if the computer is a surface, change res to 1280x800 (necessary for Netsmart)
    # this has stopped working, need to find a solution that still works
    # if ($computerName -match "surface")
    # {        Invoke-Expresion -Command "$($installdrive)\debloat files\Set-DisplayResolution.ps1"    }

    # if this isn't the 3rd run option and script is run as admin, install software
    if (($runOption -ne 3) -and (test-administrator))
    {
        # if Office 2016 isn't installed, uninstall Office 365, this takes a minute in background so it's done first
        if (-not (Test-Path 'c:\Program Files (x86)\Microsoft Office\Office16\'))
        {
            $timeStarted = get-date -format HH:mm
            Add-Content $logFile("removing Office 365 at $timeStarted")
            $roleTypeLabel.text += "`n removing Office 365 at $timeStarted"
            Start-Process "$($installdrive)\debloat files\OffScrub365C2Run_working.vbs" -Wait
        }

        # if Bitdefender isn't installed, install it (requires admin). 
        if (-not (Test-Path 'c:\Program Files\Bitdefender\'))
        {
            $timeStarted = get-date -format HH:mm
            Add-Content $logFile("Bitdefender install started at $timeStarted")
            $roleTypeLabel.text += "`n Bitdefender install started at $timeStarted"
            Start-Process -FilePath "$($installdrive)\install files\epskit_x64.exe" -ArgumentList '/bdparams /silent' -Wait
        }
        else
        {
            Add-Content $logFile("Bitdefender already installed, skipping")
            $roleTypeLabel.text += "`n Bitdefender already installed, skipping"
        }

        # if Netsmart isn't installed, install it. 
        if ((-not (Test-Path 'c:\Program Files (x86)\Allscripts Homecare\')) -and (-not (Test-Path 'c:\Program Files (x86)\Netsmart Homecare\')))
        {
            # install old Visual C in order to install NS 19.1
            $timeStarted = get-date -format HH:mm
            Add-Content $logFile("Visual C 2005 install started at $timeStarted")
            $roleTypeLabel.text += "`n Visual C 2005 install started at $timeStarted"
            Start-Process "$($installdrive)\install files\2005.exe" /Q -Wait
            
            # install new Visual C for NS 19.2 and on
            $timeStarted = get-date -format HH:mm
            Add-Content $logFile("Visual C 2015 install started at $timeStarted")
            $roleTypeLabel.text += "`n Visual C 2015 install started at $timeStarted"
            Start-Process "$($installdrive)\install files\2015.exe" /Q -Wait

            # install SQL Server Compact 4.0
            $timeStarted = get-date -format HH:mm
            Add-Content $logFile("SQL Compact 4.0 install started at $timeStarted")
            $roleTypeLabel.text += "`n SQL Compact 4.0 install started at $timeStarted"
            Start-Process "$($installdrive)\install files\SSCERuntime_x64-ENU.msi" /QN -Wait

            # install Crystal Reports for NS 19.2 and on
            $timeStarted = get-date -format HH:mm
            Add-Content $logFile("Crystal Reports install started at $timeStarted")
            $roleTypeLabel.text += "`n Crystal Reports install started at $timeStarted"
            Start-Process "$($installdrive)\install files\Crystal Reports32bit_13_0_24.msi" /Q -Wait

            Add-Content $logFile("Netsmart install started at $timeStarted")
            $roleTypeLabel.text += "`n Netsmart install started at $timeStarted"
            Start-Process "$($installdrive)\install files\NHCClientSetup.exe" "/S /v/qn" -Wait
        }
        # else attempt to delete shortcuts on users desktop, we want the shortcut on the global desktop instead
        else
        {
            Add-Content $logFile("Netsmart already installed, skipping")
            $roleTypeLabel.text += "`n Netsmart already installed, skipping"
            del $env:userprofile\desktop\Netsmart*
            del $env:userprofile\desktop\Allscript*
        }
        # if Allscripts is installed and there's no global shortcut, add it
        if ((-not (Test-Path c:\users\public\desktop\Allscript*)) -and (Test-Path 'c:\Program Files (x86)\Allscripts Homecare\'))
        {
            copy "$($installdrive)\debloat files\Allscripts Homecare.lnk" c:\users\public\desktop\
        }
        # if Netsmart is installed and there's no global shortcut, add it
        if ((-not (Test-Path c:\users\public\desktop\Netsmart*)) -and (Test-Path 'c:\Program Files (x86)\Netsmart Homecare\'))
        {
            copy "$($installdrive)\debloat files\Netsmart Homecare Client.lnk" c:\users\public\desktop\
        }

        # if Adobe Reader isn't installed, install it
        if (-not (Test-Path 'C:\Program Files (x86)\Adobe\Acrobat Reader DC\'))
        {
            $timeStarted = get-date -format HH:mm
            Add-Content $logFile("Adobe Reader install started at $timeStarted")
            $roleTypeLabel.text += "`n Adobe Reader install started at $timeStarted"
            Start-Process "$($installdrive)\Adobe Reader\AcroRead.msi" /qn -Wait
        }
        # else try to delete global desktop shortcut
        else
        {
            del 'c:\users\public\desktop\acrobat*'
            Add-Content $logFile("Adobe Reader already installed, skipping")
            $roleTypeLabel.text += "`n Adobe Reader already installed, skipping"
        }

        # if Chrome isn't installed, install it
        if (-not (Test-Path 'C:\Program Files (x86)\Google\Chrome'))
        {
            $timeStarted = get-date -format HH:mm
            Add-Content $logFile("Chrome install started at $timeStarted")
            $roleTypeLabel.text += "`n Chrome install started at $timeStarted"
            Start-Process "$($installdrive)\chrome\GoogleChromeStandaloneEnterprise64.msi" /qn -Wait
        }
        else
        {
            Add-Content $logFile("Chrome already installed, skipping")
            $roleTypeLabel.text += "`n Chrome already installed"
        }
        # delete the MS Edge shortcut
        del $env:userprofile\desktop\microsoft*

        # if Office 2016 isn't installed, install it (and activate it), Copy shortcut regardless (if 2016 is installed)
        if (-not (Test-Path 'c:\Program Files (x86)\Microsoft Office\Office16\'))
        {
            $timeStarted = get-date -format HH:mm
            Add-Content $logFile("Office install started at $timeStarted")
            $roleTypeLabel.text += "`n Office install started at $timeStarted"
            Start-Process "$($installdrive)\office 2016\setup.exe" '/adminfile 0setup.msp' -Wait
            # officekey : product key for Office 2016
            $officeKey = Get-Content -Path "$($installdrive)\debloat files\officekey.txt"
            cscript 'C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS' /inpkey:$officeKey
        }
        else
        {
            Add-Content $logFile("Office already installed, skipping")
            $roleTypeLabel.text += "`n Office already installed, skipping"
        }
        if (-not (Test-Path $env:userprofile\desktop\Outlook.lnk) -and (Test-Path 'c:\program files (x86)\Microsoft office\office16'))
        {
            copy "$($installdrive)\debloat files\Outlook 2016.lnk" $env:userprofile\desktop
        }
    }

    # since Qliq is garbage software, it has to be installed on every user
    # (notice it will install all QliqConnect versions on USB drive, so keep only the latest version on USB)
    # ((you can rename the old Qliq to anything as long as it doesn't have "QliqConnect" anywhere in the name))
    if (-not (Test-Path $env:userprofile\AppData\Roaming\Qliqsoft))
    {
        $timeStarted = get-date -format HH:mm
        Add-Content $logFile("Qliq install started at $timeStarted")
        $roleTypeLabel.text += "`n Qliq install started at $timeStarted"
        Start-Process "$($installdrive)\install files\QliqConnect*" "/m /qn" -Wait
    }
    else
    {
        Add-Content $logFile("Qliq already installed, skipping")
        $roleTypeLabel.text += "`n Qliq already installed, skipping"
    }
    

    ##################################################
    # this portion runs only if PS is running elevated
    ################################################## 
    if (test-administrator)
    {
        Add-Content $logFile("running debloat stuff as administrator (just ignore registry key errors)")
        $roleTypeLabel.text += "`n running admin elevated debloat"

        # applying VPN L2TP fix - deprecated since we're using SSTP now
        # REG ADD HKLM\SYSTEM\CurrentControlSet\Services\PolicyAgent /v AssumeUDPEncapsulationContextOnSendRule /t REG_DWORD /d 2 /f

        # cleaning up app bloat
        foreach ($app in $apps) 
        {
            $package = Get-AppxPackage -Name $app -AllUsers
            if ($package -ne $null) {
                $package | Remove-AppxPackage -ErrorAction SilentlyContinue
                Get-AppXProvisionedPackage -Online | Where-Object DisplayName -EQ $app | Remove-AppxProvisionedPackage -Online
                $appPath = "$Env:LOCALAPPDATA\Packages\$app*"
                Remove-Item $appPath -Recurse -Force -ErrorAction 0
            }
        }

        # importing start menu tiles, please note this only affects new logins not this account
        Import-StartLayout -LayoutPath "$($installdrive)\debloat files\THBRlayout.xml" -MountPath $env:SystemDrive\

        # setting default apps
        dism /online /import-defaultappassociations:"$($installdrive)\debloat files\THBRdefaultapps.xml"

        # remove OneDrive
        REG ADD "HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" /v SystemIsPinnedToNameSpaceTree /t REG_DWORD /d 0 /f



        # this last section forces UAC back on, so setup the area we'll be working in
        $Key = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" 

        # 2 and 1 are always notify, 5 and 1 is the default, 5 and 0 does not dim desktop
        $ConsentPromptBehaviorAdmin_Value = 5  
        $PromptOnSecureDesktop_Value = 1 

        # if UAC registry key not found, create it
        if ((Test-Path -Path $key) -Eq $false) 
        { 
            New-Item -ItemType Directory -Path $key | Out-Null 
        }

        # set the registry values
        Set-ItemProperty -Path $key -Name "ConsentPromptBehaviorAdmin" -Value $ConsentPromptBehaviorAdmin_Value -Type "Dword"  
        Set-ItemProperty -Path $key -Name "PromptOnSecureDesktop" -Value $PromptOnSecureDesktop_Value -Type "Dword"  
        Set-ItemProperty -Path $key -Name "EnableLUA" -Value 1 -Type "Dword"

    }

    ######################################################
    # this portion runs only if PS is running non-elevated
    ######################################################
    else
    {
        # copy staff laptop documents over to this computer and unhide the shortcut that will be used
        Add-Content $logFile("Copying staff laptop docs")
        $roleTypeLabel.text += "`n Copying staff laptop docs"

        robocopy $netpath"staff laptop documents" 'c:\staff laptop documents' /mir /nfl /ndl /ns /nc /njh
        ATTRIB -H 'c:\staff laptop documents\staff laptop documents - shortcut.lnk'

        # remove the un-needed batch file, and move shortcut from SLD to desktop
        remove-item 'c:\staff laptop documents\copy.bat'
        # if the shortcut already exists, delete from SLD
        $enviro = $env:userprofile
        if (test-path ("$($enviro)\desktop\staff laptop documents - shortcut.lnk"))
        {
            remove-item 'c:\staff laptop documents\staff laptop documents - shortcut.lnk'
        }
        # else just move the shortcut to desktop
        else
        {            
            move-item -path 'c:\staff laptop documents\staff laptop documents - shortcut.lnk' -destination $env:userprofile\desktop
        }

        # pull the users role from AD
        $dom = $env:userdomain
        $usr = $env:username
        $userRole = ([adsi]"WinNT://$dom/$usr,user").description

        # if user is a social worker, copy SW shortcut
        if ($userRole -like 'SW*')
        {
            copy-item "$($installdrive)\debloat files\SW POC Update Clinical Note Template.txt - Shortcut.lnk" $env:userprofile\desktop
        }

        # if user is a RN or LPN, copy RN shortcut
        if (($userRole -like 'RN*') -or ($userRole -like 'LPN*') -or ($userRole -like 'IPU RN'))
        {
            copy-item "$($installdrive)\debloat files\RN POC Update Clinical Note Template.txt - Shortcut.lnk" $env:userprofile\desktop
        }

        # if user is a chaplain, copy chaplain shortcut
        if ($userRole -like 'Chaplain*')
        {
            copy-item "$($installdrive)\debloat files\Chap POC Update Clinical Note Template.txt - Shortcut.lnk" $env:userprofile\desktop
        }
    }


    # read vpn info from files
    # vpnsstp.txt : VPN name, VPN server address, and domain name all on separate lines  
    $vpn     = Get-Content -Path "$($installdrive)\debloat files\vpnsstp.txt"
    $VPNConnectionName = $vpn[0]

    # the file must contain the name, address, and psk all on separate lines
    if ($vpn.count -ne 3)
    {
        Add-Content $logFile("vpn file is damaged")
        $roleTypeLabel.text += "`n VPN file is damaged!!"
    }
    # creating VPN connection if it doesn't exist
    elseif (-not ((get-vpnconnection).Name -eq $VPNConnectionName))
    {
        $VPNConnectionName = $vpn[0]
        $VPNServerAddress  = $vpn[1]
        $DNSSuffix         = $vpn[2]
        
        Add-VpnConnection -Name $VPNConnectionName -ServerAddress $VPNServerAddress -TunnelType "Sstp" -EncryptionLevel "Required" -AuthenticationMethod MsChapv2 -SplitTunneling -PassThru

        # Adding RFC1819 routes to VPN $VPNConnectionName...
        Add-VpnConnectionRoute -ConnectionName $VPNConnectionName -DestinationPrefix "10.0.0.0/8" -PassThru
        Add-VpnConnectionRoute -ConnectionName $VPNConnectionName -DestinationPrefix "172.16.0.0/12" -PassThru
        Add-VpnConnectionRoute -ConnectionName $VPNConnectionName -DestinationPrefix "192.168.0.0/16" -PassThru

        # Setting Connectin DNSSuffix to $DNSSuffix
        Set-VPNConnection -Name "$VPNConnectionName" -DNSSuffix $DNSSuffix -UseWinlogonCredential $True

        Add-Content $logFile("creating VPN and WiFi connections")
        $roleTypeLabel.text += "`n creating VPN and WiFi connections"
    }

    # setting Cortana to icon only
    REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Search /v SearchboxTaskbarMode /t REG_DWORD /d 1 /f

    # cleaning up taskbar
    REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\People /v PeopleBand /t REG_DWORD /d 0 /f
    REG ADD HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\WindowsInkWorkspace /v AllowWindowsInkWorkspace /t REG_DWORD /d 0 /f
    REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced /v ShowTaskViewButton /t REG_DWORD /d 0 /f
    REG ADD HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Windows\Explorer /v NoPinningStoreToTaskbar /t REG_DWORD /d 1 /f
    REG ADD HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Explorer /v NoPinningStoreToTaskbar /t REG_DWORD /d 1 /f

    # turn off background apps, and fix desktop icon bug (VPN)
    REG ADD HKCU\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications /v GlobalUserDisabled /t REG_DWORD /d 1 /f
    REG ADD HKCU\Software\Microsoft\Windows\CurrentVersion\Search /v BackgroundAppGlobalToggle /t REG_DWORD /d 0 /f
    REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes" /v ThemeChangesDesktopIcons /t REG_DWORD /d 1 /f

    # fix a bug with temp account not disappearing or trying to auto logon
    REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v AutoAdminLogon /t REG_DWORD /d 0 /f
    Remove-Item "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultUserName" -ErrorAction SilentlyContinue

    # remove 3d objects, music, videos
    Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}" -ErrorAction SilentlyContinue
    Remove-Item -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}" -ErrorAction SilentlyContinue
    Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}" -ErrorAction SilentlyContinue
    Remove-Item -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}" -ErrorAction SilentlyContinue
    Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}" -ErrorAction SilentlyContinue
    Remove-Item -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}" -ErrorAction SilentlyContinue

    # turn off tablet mode
    REG ADD HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\ImmersiveShell /v TabletMode /t REG_DWORD /d 0 /f

    # setting power parameters
    powercfg -change -monitor-timeout-ac 30
    powercfg -change -monitor-timeout-dc 10
    powercfg -change -standby-timeout-ac 0
    powercfg -change -standby-timeout-dc 30
    powercfg -change -hibernate-timeout-ac 0
    powercfg -change -hibernate-timeout-dc 120
    powercfg -change -disk-timeout-ac 0

    # set the time zone to be Central time
    Set-TimeZone -Name "Central Standard Time"

    # create a scheduled task to auto reboot computer
    if ($runOption -eq 3)
    {
        # if the auto restart scheduled task isn't found, then create it
        $gst = Get-ScheduledTask -taskname "auto restart" 2>$Null
        if ($null -eq $gst)
        {
            $taskAction  = New-ScheduledTaskAction -execute "c:\Windows\System32\shutdown.exe" -Argument "-f -r"
            $taskTrigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 4am
            Register-ScheduledTask -Action $taskAction -Trigger $taskTrigger -TaskName "auto restart"
            Set-ScheduledTask -TaskName "auto restart" -User "NT Authority\system" -Settings $(New-ScheduledTaskSettingsSet -StartWhenAvailable -WakeToRun)
        }
    }

    # if logged in as 'temp' account & this is first run, mark first run done and install Win 10 key
    # else it must be 2nd run so join domain and reboot PC
    if ($env:username -eq 'temp')
    {
        # setting up to determine if machine is connected to ethernet
        $notConnected = $true
        $ethers = get-netadapter -name "ethernet*"
        foreach ($ether in $ethers)
        {
            if ($ether.status -eq "Up")
            {
                $notConnected = $false
            }
        }
        # if we can't find WiFi or connected ethernet warn user
        if (($null -eq (get-netadapter -name "wi-fi*")) -and $notConnected) 
        { 
            [System.Windows.Forms.MessageBox]::Show("No Wi-Fi found. Machine must be connected to network to proceed safely. Ignore if machine is plugged into ethernet", "Warning")  
        }

        # if first run, activate Windows, rename machine, reboot
        if ($runOption -eq 1)
        {
            # this file on the desktop marks it as having been run once
            copy "$($installdrive)\debloat files\Debloat-Windows.ps1" $env:userprofile\desktop\Debloat-Windows.ps1   
            # installs Win10 key only on first script run (pulled from a file so it's not GitHub visible)
            Add-Content $logFile("Installing Windows 10 key, renaming computer to $computerName, rebooting after")
            $roleTypeLabel.text += "`n Installing Windows 10 key, renaming computer to $computerName"
            Start-Sleep 3
            # wait for win10 key to install, can't do it silently so placing it right in front of a reboot command
            # key.txt : Windows 10 product key
            $key = "/ipk $(get-content "$($installdrive)\debloat files\key.txt")"
            Start-Process slmgr $key -Wait
            rename-computer -newname $computerName -force -restart
        }
        # else it must be 2nd run, so join domain, disable temp account, and reboot
        else
        {
            Add-Content $logFile("2nd run, domain joining")
            $roleTypeLabel.text += "`n joining domain"
            # pull domain name from a file and add this machine to the network
            # domain.txt : domain name, ex: yourdomain.com  (or yourdomain.local if old domain)
            $domain = Get-Content -Path "$($installdrive)\debloat files\domain.txt"
            Start-Sleep 3
            # disabling temp account before
            Add-Content $logFile("disabling temp account")
            $roleTypeLabel.text += "`n disabling temp account"
            Remove-LocalUser temp
            add-computer -domainname $domain -credential $domain"\Administrator" -force -restart
        }
    }

    Stop-Transcript

    # wait 1 second and close display box
    $roleTypeLabel.text += "`n `n Complete!"
    Start-Sleep 5


}


generateForm
