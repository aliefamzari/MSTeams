<#
    .SYNOPSIS
        Orsted SD Script Tool
        MS Teams Deployment for Single PC
    
    .DESCRIPTION
        Reinstall MS Teams.
        Remove MS Teams residual when uninstall MS Teams.
        Download MS Teams using bootstrap.exe, MSIX package, or the classic MS Teams.
        Delete browser cache.
        Register Teams addin for bootstrap and msix install.
    
    .PARAMETER DeploymentType
        Choose between MSIX, bootstrap or Classic teams to be install.
        Default installation without paramater is MSIX.
        MSIX:   
            The script use the public URL which only available in x86.
            If Microsoft Store is disabled, can only be deploy using Add-AppxPackage CMDlet.
            The package by default does not include MS Teams Addin for Outlook
        Bootstrap:  
            Install x64 version
        Classic:    
            Install the Classic version of MS Teams. 

    

    .NOTES
        ===================================================================
        Created with: Visual Studio Code
        Git Control: GitHub - Migrated over Azure DevOps
        Repository: https://github.com/aliefamzari/MSTeams
        Author: Alif Amzari, ALMAZ
        Encoding: UTF-8 with BOM
        Previous Author: XSIOL, TOBKO
        Known issues:
            1.  Teams Addin for Outlook   
                The current New MS Teams deployment package (Bootstrap and MSIX),
                do not include MS Teams Addin for Outlook. When Classic Teams
                removed, the folder contains DLL for the Addins get removed too.
                To overcome this, without downloading both classic and New Teams
                during the deployment, the scrip will take a backup of current
                $env:LOCALAPPDATA\Microsoft\TeamsMeetingAddin. If there folder
                is missing or empty, the script will stop and ask to manually
                install Teams Classic. 
            2. Bootstrap installation need elevation
                The Classic teams executables and MSIX package pass UAC or EPM check.
                Possible cause is additional digital signature embedded into the executables. 

        ===================================================================
#>

function MSTeamsReinstallFull {
    param (
        [ValidateSet("MSIX","BootStrap","Classic")]
        [string]$DeploymentType,
        [switch]$byPassSSL,
        [ValidateSet("TeamsAddinFix")]
        [string]$Options,
        [ValidateSet("all","teams")]
        [string]$cacheType,
        [switch]$BackupTeamsAddin
    )
    $ErrorActionPreference = "SilentlyContinue"
    $ProgressPreference = "SilentlyContinue"
    $InstallerDir = "$ENV:USERPROFILE\Downloads"
    $TeamsMeetingAddinDir = "$env:LOCALAPPDATA\Microsoft\TeamsMeetingAddin"
    $TeamsMeetingAddinDir2 = "C:\Program Files (x86)\Microsoft\TeamsMeetingAddin"
    $backupPath = "$env:LOCALAPPDATA\Microsoft"
    $backupDestination = Join-Path $backupPath "TeamsMeetingAddinBackup"

    #Region Function
    function ShowBannerReinstall {
        Write-Host "Are you sure you wish to completely reinstall MS Teams?"
        if ($DeploymentType -ne "Classic") {
            Write-Host "The default deployment will removed the Classic Teams in favor of New Teams."
        }
        if ($cacheType -eq "all") {
            Write-Host "This will delete cache of Internet Explorer, MS Edge, Chrome & Firefox" -ForegroundColor Red
            Write-Host "Please ensure you save and close your Browsers" -ForegroundColor Red
        }
        Write-Host "Please ensure you save and close your Outlook" -ForegroundColor Red
        # Display the current deployment type in green with a newline
        Write-Host -NoNewline "The current deployment is "
        Write-Host -NoNewline -ForegroundColor Green "$DeploymentType"
        Write-Host
    }
    function KillApp {
        #Stops Microsoft Teams
        Write-Host "Stopping Teams Process" -ForegroundColor Yellow

        try{
            Get-Process -ProcessName *Teams*  | Stop-Process -Force
            Start-Sleep -Seconds 3
            Write-Host "Teams Process Sucessfully Stopped" -ForegroundColor Green
        }
        
        catch{
            Write-Output $_
        }

        #Stop Outlook
        $outlook = get-process -name "Outlook"
        if ($outlook) {
            Write-Host "Stopping MS Outlook Process" -ForegroundColor Yellow
            $outlook.CloseMainWindow() |Out-Null
            Start-Sleep 3
            if (!$outlook.HasExited) {
                $outlook | Stop-Process -Force
            }   
            Write-Host "MS Outlook Succesfully Stopped" -ForegroundColor Green
        }
    }
    function MSIXInstall {
        Write-Host "Installing Teams" -ForegroundColor Magenta
        try {
            Add-AppxPackage -Path $InstallerLocation
        }
        catch {
            Write-Host "AppX error!" -ForegroundColor Red
            #Read-Host "Press enter to exit"
            Break  
        }
    }
    function BootStrapInstall {
        Write-Host "Installing Teams" -ForegroundColor Magenta
        try {
            Unblock-File -Path $InstallerLocation
            $proc = Start-Process -FilePath $InstallerLocation -ArgumentList "-p" -PassThru
        }
        catch {
            Write-Host "Elevation or EPM error. Script stop" -ForegroundColor Red
            # Read-Host "Press enter to exit"
            Break
            # Stop-Process -Id $PID
            # Exit
        }
        # Add-AppxPackage -Path $InstallerLocation
        $proc.WaitForExit()
    }
    function ClassicInstall {
        Write-Host "Installing Teams" -ForegroundColor Magenta
        try {
            # $proc = Start-Process -FilePath $InstallerLocation -ArgumentList "-s" -PassThru
            $proc = Start-Process -FilePath $InstallerLocation -PassThru
        }
        catch {
            Write-Host "Elevation or EPM error. Script stop" -ForegroundColor Red
            # #Read-Host "Press enter to exit"
            Break  
        }
        # finally {
        #     $jsonContent = '{"enableProcessIntegrityLevel": false}'
        #     $filePath = Join-Path $env:APPDATA\Microsoft\Teams hooks.json
        #     $jsonContent | Set-Content -Path $filePath -Encoding UTF8
        #     Write-Host "JSON file created at: $filePath"
        # }
        $proc.WaitForExit()
    }
    # function StartApp {
    #     if ($DeploymentType -in "MSIX", "BootStrap") {
    #         $AppPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps\ms-teams.exe"
    #     } else {
    #         $AppPath = "$env:APPDATA\Microsoft\Teams\Teams.exe"
    #     }
    #     $ExeExist = Test-Path $AppPath
    #     if ($ExeExist) {
    #         Write-Host "Starting MS Teams" -ForegroundColor Green
    #         Start-Process $AppPath
    #     }
    #     else {
    #         Write-Host "Installation Failed. App Could not start" -ForegroundColor Red
    #         Write-Host "Please re-install again using different deployment method" -ForegroundColor Red
    #     }
    # }
    function StartApp {
        if ($DeploymentType -in "MSIX", "BootStrap") {
            $AppPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps\ms-teams.exe"
        } else {
            $AppPath = "$env:APPDATA\Microsoft\Teams\Teams.exe"
        }
        Switch ($DeploymentType){
            'MSIX' {
                $ExeExist = Test-Path $AppPath
                if ($ExeExist) {
                    Write-Host "Starting MS Teams" -ForegroundColor Green
                    Start-Process $AppPath
                }
                else {
                    Write-Host "Installation Failed. App Could not start" -ForegroundColor Red
                    Write-Host "Please re-install again using different deployment method" -ForegroundColor Red
                }
            }
            'BootStrap'{
                while ($true) {
                    $process = Get-Process -Name TeamsBootStrapper -ErrorAction SilentlyContinue
                    if (-not $process) {
                        # Process has ended, execute your script or command here
                        $ExeExist = Test-Path $AppPath
                        if ($ExeExist) {
                            Write-Host "Starting MS Teams" -ForegroundColor Green
                            Start-Process $AppPath
                        }
                        else {
                            Write-Host "Installation Failed. App Could not start" -ForegroundColor Red
                            Write-Host "Please re-install again using different deployment method" -ForegroundColor Red
                        }
                        break  # Exit the loop
                    }
                    # Wait for a moment before checking again
                    Start-Sleep -Seconds 5
                }

            }
            'Classic'{
                $ExeExist = Test-Path $AppPath
                if ($ExeExist) {
                    Write-Host "Starting MS Teams" -ForegroundColor Green
                    Start-Process $AppPath
            }
            }
        }
    }
    function TeamsAddinFix {
        param (
            [switch]$Resiliency
        )
        if (Test-path $TeamsMeetingAddinDir ){
            $TruePath = $TeamsMeetingAddinDir
        }
        if (Test-Path $TeamsMeetingAddinDir2){
            $TruePath = $TeamsMeetingAddinDir2
        }
        
        if (-not (Test-Path $TeamsMeetingAddinDir) -and (Test-path $TeamsMeetingAddinDir2)){
            Write-Host "Warning!! Teams Meeting DLL does not exist" -ForegroundColor Red
            Write-Host "Please reinstall MS Teams" -ForegroundColor Red
            Return
        }
        $items = Get-ChildItem $TruePath
        if ($items.Count -eq 0) {
            Write-Host "Warning!! $TruePath is empty" -ForegroundColor Red
            Write-Host "Please reinstall MS Teams" -ForegroundColor Red
            # Read-Host "Press enter to exit"
            return
        } 
                # Register TeamsAddin DLL   
        $LattestDLLversion = (Get-ChildItem -Path $TruePath -Directory |Sort-Object CreationTime -Descending| Select-Object -First 1)
        $LattestDLLversion = ($LattestDLLversion).FullName
        $teamsdotdead = "$LattestDLLversion\.dead"
        $teamsdll = "$LattestDLLversion\x64\Microsoft.Teams.AddinLoader.dll"
                
        Write-Host "Removing .dead file if exist" -ForegroundColor Yellow
        if (Test-Path -Path $teamsdotdead) {
            Remove-Item -Path $teamsdotdead
            Write-host ".dead file found and removed" -ForegroundColor Green
        }
        else {
            Write-host "No .dead file exist"  -ForegroundColor Green
        }
        # write-host "Deregistring Microsoft.Teams.AddinLoader.dll" -ForegroundColor Yellow
        # start-sleep 5
        regsvr32.exe /U "$teamsdll" /s
        write-host "Registring Microsoft.Teams.AddinLoader.dll" -ForegroundColor Green
        regsvr32.exe /n /i:user "$teamsdll" /s
        # Write-Host "Done"
        
        # Check if Microsoft Teams add-ins for Outlook are enabled
        $TeamsMeetingAddinRegPath = "HKCU:\SOFTWARE\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect" 
        $FastConnectReg = Get-Item -Path $TeamsMeetingAddinRegPath -ErrorAction SilentlyContinue
        
        if ($null -eq $FastConnectReg) {
            Write-Host "Microsoft Teams add-ins for Outlook is not enable" -ForegroundColor Yellow
            Write-Host "Enabling Teams Addin in Outlook" -ForegroundColor Yellow
            New-Item -Path $TeamsMeetingAddinRegPath
            New-ItemProperty -Path $TeamsMeetingAddinRegPath -Name "Description" -Value "Microsoft Teams Meeting Add-in for Microsoft Office"
            New-ItemProperty -Path $TeamsMeetingAddinRegPath -Name "FriendlyName" -Value "Microsoft Teams Meeting Add-in for Microsoft Office"
            New-ItemProperty -Path $TeamsMeetingAddinRegPath -Name "LoadBehavior" -PropertyType DWord -Value 3
            Write-Host "Teams Addins Enabled" -ForegroundColor Green
        } 
        else {
            $CurLoadBehavior = $FastConnectReg.GetValue("LoadBehavior")
            if ($CurLoadBehavior -eq 3) {
                Write-Host "Microsoft Teams add-ins LoadBehavior is already set to 3." -ForegroundColor Yellow
            } else {
                Write-Host "Microsoft Teams add-ins LoadBehavior is $CurLoadBehavior" -ForegroundColor Yellow
                Set-ItemProperty -path $TeamsMeetingAddinRegPath -Name LoadBehavior -Value 3
                $newloadbehavior = $FastConnectReg.GetValue("LoadBehavior")
                Write-Host "Microsoft Teams add-ins LoadBehavior has been set to $newloadbehavior." -ForegroundColor Green
            }
        }
        
        #Region KB0016283 - Add registry entry (if not exist) - KB from Dina Rantzau https://onewebshop.service-now.com/kb_view.do?sysparm_article=KB0016283
if ($Resiliency) {
        Write-Host "Applying KB0016283" -ForegroundColor Yellow
        # ResiliencyTeamsAddinConnect1
        $ResiliencyTeamsAddinPath1 = "HKCU:\software\Policies\Microsoft\office\16.0\outlook\resiliency\addinlist\"
        $ResiliencyTeamsAddinConnect1 = Get-ItemPropertyValue -Path "HKCU:\software\Policies\Microsoft\office\16.0\outlook\resiliency\addinlist\" -Name TeamsAddin.Connect -ErrorAction SilentlyContinue
        $ResiliencyTeamsAddinFastConnect1 = Get-ItemPropertyValue -Path "HKCU:\software\Policies\Microsoft\office\16.0\outlook\resiliency\addinlist\" -Name TeamsAddin.FastConnect -ErrorAction SilentlyContinue
    
        if ($null -eq $ResiliencyTeamsAddinConnect1) {
            Write-Host "Create ResiliencyTeamsAddinConnect1"
            New-ItemProperty -Path $ResiliencyTeamsAddinPath1 -Name "TeamsAddin.Connect" -Value "1" -PropertyType "String"
        }
        else {
            Write-host "ResiliencyTeamsAddinConnect1 already exist"
        }
        if ($null -eq $ResiliencyTeamsAddinFastConnect1) {
            Write-Host "Create ResiliencyTeamsAddinFastConnect1"
            New-ItemProperty -Path $ResiliencyTeamsAddinPath1 -Name "TeamsAddin.FastConnect" -Value "1" -PropertyType "String"
    
        }
        else {
            Write-host "ResiliencyTeamsAddinFastConnect1 already exist"
        }
    
        #ResiliencyTeamsAddinConnect2
        $ResiliencyTeamsAddinPath2 = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList\"
        $ResiliencyTeamsAddinConnect2 = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList\" -Name TeamsAddin.Connect
        $ResiliencyTeamsAddinFastConnect2 = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList\" -Name TeamsAddin.FastConnect
    
        if ($null -eq $ResiliencyTeamsAddinConnect2) {
            Write-Host "Create ResiliencyTeamsAddinConnect2"
            New-ItemProperty -Path $ResiliencyTeamsAddinPath2 -Name "TeamsAddin.Connect" -Value "1" -PropertyType DWord
        }
        else {
            Write-host "ResiliencyTeamsAddinConnect2 already exist"
        }
        if ($null -eq $ResiliencyTeamsAddinFastConnect2) {
            Write-Host "Create ResiliencyTeamsAddinFastConnect2"
            New-ItemProperty -Path $ResiliencyTeamsAddinPath2 -Name "TeamsAddin.FastConnect" -Value "1" -PropertyType DWord
        }
        else {
            Write-host "ResiliencyTeamsAddinFastConnect2 already exist"
        }
}
        #endregion
        Write-Host "Done Registring MS TeamsAddin!" -ForegroundColor Green
    }
    function BackupTeamsAddin {
            $DirExist = Test-Path $TeamsMeetingAddinDir
            if (!$DirExist) {
                Write-Host "Error!`n$TeamsMeetingAddinDir does not exist. `nPlease reinstall MS Teams Classic" -ForegroundColor Red
                # Read-Host "Press enter to exit"
                Break
            }
        
            $sourceItems = Get-ChildItem $TeamsMeetingAddinDir
            $DirEmpty = $sourceItems.Count -eq 0
            if ($DirEmpty) {
                Write-Host "Error!`n$TeamsMeetingAddinDir is empty. `nPlease reinstall MS Teams Classic." -ForegroundColor Red
                # Read-Host "Press enter to exit"
                Break
            }
        
            if (-not (Test-Path $backupPath)) {
                Write-Host "Error!`nBackup directory $backupPath does not exist." -ForegroundColor Red
                # Read-Host "Press enter to exit"
                Break
            }
        
            $backupDestination = Join-Path $backupPath "TeamsMeetingAddinBackup"
        
            # Check if the destination folder already exists
            if (Test-Path $backupDestination) {
                Write-Host "$backupDestination already exists. Proceed to remove it."  -ForegroundColor Red
                Remove-item -Path $backupDestination -Recurse
                #Read-Host "Press enter to exit"
                # Break
            }
            Copy-Item -Path $TeamsMeetingAddinDir -Destination $backupDestination -Recurse
            Write-Host "Backup of $TeamsMeetingAddinDir created in $backupDestination" -ForegroundColor Green
    }
    function ClearCache {
        param(
            [string]$cacheType = "all"
        )
    
        function CleanCache($path) {
            if (Test-Path $path) {
                Get-ChildItem -Path $path | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            }
        }
    
        function StopProcess($processName) {
            Get-Process -Name $processName -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        }
    
        Write-Host "Clearing Teams Disk Cache" -ForegroundColor Yellow
        if ($cacheType -eq "teams" -or $cacheType -eq "all") {
            "application cache\cache", "blob_storage", "databases", "cache", "gpucache", "Indexeddb", "Local Storage", "tmp" | ForEach-Object { CleanCache "$env:APPDATA\Microsoft\teams\$_" }  #Classic Teams
            Write-Host "Classic Teams Cache Cleaned" -ForegroundColor Green
            "MSTeams" |ForEach-Object {CleanCache "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\$_"} #New Teams
            Write-Host "New Teams Cache Cleaned" -ForegroundColor Green
        }
    
        if ($cacheType -eq "all") {
            Write-Host "Stopping Chrome Process" -ForegroundColor Yellow
            StopProcess "Chrome"
            Start-Sleep -Seconds 3
            Write-Host "Chrome Process Successfully Stopped" -ForegroundColor Green
            
            Write-Host "Clearing Chrome Cache" -ForegroundColor Yellow
            "Cache", "Cookies", "Web Data" | ForEach-Object { CleanCache "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\$_" }
            Write-Host "Chrome Cache Cleaned" -ForegroundColor Green
            
            Write-Host "Stopping IE & Edge Process" -ForegroundColor Yellow
            "MicrosoftEdge", "MSEdge", "IExplore" | ForEach-Object { StopProcess $_ }
            Write-Host "Internet Explorer and Edge Processes Successfully Stopped" -ForegroundColor Green
            
            Write-Host "Clearing IE & Edge Cache" -ForegroundColor Yellow
            RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 8
            RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 2
            CleanCache "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache"
            Write-Host "IE and Edge Cache Cleaned" -ForegroundColor Green
            
            Write-Host "Stopping Firefox Process" -ForegroundColor Yellow
            StopProcess "Firefox"
            Start-Sleep -Seconds 3
            Write-Host "Firefox Process Successfully Stopped" -ForegroundColor Green
            
            Write-Host "Clearing Firefox Cache" -ForegroundColor Yellow
            CleanCache "$env:LOCALAPPDATA\Mozilla\Firefox\Profiles"
            Write-Host "Firefox Cache Cleaned" -ForegroundColor Green
        }
        Write-Host "Cleanup Complete..." -ForegroundColor Green
    
    }
    function TeamsUninstall {
        $NewMSTeams = Get-AppxPackage -Name MSTeams
        try {
        if ($NewMSTeams){
        Write-Host "Uninstalling New Teams process" -ForegroundColor Yellow
        Remove-AppxPackage $NewMSTeams
        }
        }
        catch {
        Write-Error -ErrorRecord $_
        }
        $TeamsPath = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Microsoft', 'Teams')
        $TeamsUpdateExePath = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Microsoft', 'Teams', 'Update.exe')
        try
        {
        if (Test-Path -Path $TeamsUpdateExePath) {
        Write-Host "Uninstalling Classic Teams process" -ForegroundColor Yellow
        # Uninstall app
        $proc = Start-Process -FilePath $TeamsUpdateExePath -ArgumentList "-uninstall -s" -PassThru
        $proc.WaitForExit()
        }
        if (Test-Path -Path $TeamsPath) {
        Write-Host "Deleting Teams directory" -ForegroundColor Yellow
        Remove-Item -Path $TeamsPath -Recurse
        }
        }
        catch
        {
        Write-Error -ErrorRecord $_
        }
    }
    function RestoreTeamsAddinBackup {
        Write-Host "Restoring TeamsAddinDLL backup" -ForegroundColor Green
        if (Test-Path $TeamsMeetingAddinDir) {
            Remove-Item $TeamsMeetingAddinDir -Recurse -Force
            Rename-Item -path $backupDestination -NewName $TeamsMeetingAddinDir
        }
        Rename-Item -path $backupDestination -NewName $TeamsMeetingAddinDir
    }
    function DownloadTeams {
        function CurlDownload {
            if ($byPassSSL) {
                curl.exe -fSLo $InstallerLocation $DownloadSource --ssl-no-revoke  # 10 second download (with progress bar)
            }
            else {
                curl.exe -fSLo $InstallerLocation $DownloadSource --ssl-no-revoke # temp solution is to force --ssl-no-revoke
            }
        }
        switch ($DeploymentType) {
            MSIX {
                $DownloadSource = "https://go.microsoft.com/fwlink/?linkid=2196106" #MSTeams-x64.msix,64bit,no elevation
                $ConverToStaticUrl = curl.exe -fks -X GET -w "%{redirect_url}" $DownloadSource -o NUL
                $StaticUrlFilename = $ConverToStaticUrl.Split('/')[-1]
                $Script:InstallerLocation = "$InstallerDir\$StaticUrlFilename"
                If([System.IO.File]::Exists($InstallerLocation) -eq $false){
                    Write-Host "Downloading Teams, please wait." -ForegroundColor Magenta
                    CurlDownload
                }
                Else{
                    Write-Host "Installer file already present in Downloads folder. Removing old installer." -ForegroundColor Yellow
                    Remove-item -path $InstallerLocation
                    Write-Host "Downloading Teams, please wait." -ForegroundColor Magenta
                    CurlDownload
                }
            }
            BootStrap {
                $DownloadSource = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409" #teamsbootstrapper.exe,64bit,require elevation
                $ConverToStaticUrl = curl.exe -fks -X GET -w "%{redirect_url}" $DownloadSource -o NUL
                $StaticUrlFilename = $ConverToStaticUrl.Split('/')[-1]
                $Script:InstallerLocation = "$InstallerDir\$StaticUrlFilename"
                If([System.IO.File]::Exists($InstallerLocation) -eq $false){
                    Write-Host "Downloading Teams, please wait." -ForegroundColor Magenta
                    CurlDownload
                }
                Else{
                    Write-Host "Installer file already present in Downloads folder. Removing old installer." -ForegroundColor Yellow
                    Remove-item -path $InstallerLocation
                    Write-Host "Downloading Teams, please wait." -ForegroundColor Magenta
                    CurlDownload
                }
            }
            Classic {
                # $DownloadSource = "https://teams.microsoft.com/downloads/desktopurl?env=production&plat=windows&arch=x64&managedInstaller=true&download=true"
                $DownloadSource = "https://go.microsoft.com/fwlink/?linkid=2187327"
                # $ConverToStaticUrl = curl.exe -fks -X GET -w "%{redirect_url}" $DownloadSource -o NUL
                $ConverToStaticUrl = curl.exe -fkLs -w "%{url_effective},%{filename_effective}" $DownloadSource -OJ
                $StaticUrlFilename = $ConverToStaticUrl.Split(',')[-1]
                # $StaticUrlFilename = $ConverToStaticUrl.Split('/')[-1]
                $Script:InstallerLocation = "$InstallerDir\$StaticUrlFilename"
                If([System.IO.File]::Exists($InstallerLocation) -eq $false){
                    Write-Host "Downloading Teams, please wait." -ForegroundColor Magenta
                    CurlDownload
                }
                Else{
                    Write-Host "Installer file already present in Downloads folder. Removing old installer." -ForegroundColor Yellow
                    Remove-item -path $InstallerLocation
                    Write-Host "Downloading Teams, please wait." -ForegroundColor Magenta
                    CurlDownload
                }
            }
        }
    }

    function UninstallTeamsAddins {
        $Program = Get-WmiObject -Class Win32_Product | Where-Object { $_.IdentifyingNumber -match "{A7AB73A3-CB10-4AA5-9D38-6AEFFBDE4C91}"}
        if ($Program) {
            Write-Host "Uninstalling Teams Addins for Outlook" -ForegroundColor Yellow
            $Program.uninstall() |Out-Null
        }
    }
    #EndRegion Function
    
    Switch ($DeploymentType) {
        'BootStrap' {
            do {
                ShowBannerReinstall
                $choice = Read-Host "Do you want to continue? (yes/no)"
            } while ($choice -notin "yes", "no", "y", "n", "Y", "N")
            if ($choice -in "yes", "y", "Y") {
                KillApp
                # BackupTeamsAddin
                DownloadTeams
                TeamsUninstall
                UninstallTeamsAddins
                ClearCache -cacheType $cacheType
                BootStrapInstall
                # RestoreTeamsAddinBackup
                # TeamsAddinFix
                StartApp
            }
        }
        'MSIX' {
            do {
                ShowBannerReinstall
                $choice = Read-Host "Do you want to continue? (yes/no)"
            } while ($choice -notin "yes", "no", "y", "n", "Y", "N")
            if ($choice -in "yes", "y", "Y") {
                KillApp
                # BackupTeamsAddin
                DownloadTeams
                TeamsUninstall
                UninstallTeamsAddins
                ClearCache -cacheType $cacheType
                MSIXInstall
                # RestoreTeamsAddinBackup
                # TeamsAddinFix
                StartApp
            }
        }
        'Classic' {
            do {
                ShowBannerReinstall
                $choice = Read-Host "Do you want to continue? (yes/no)"
            } while ($choice -notin "yes", "no", "y", "n", "Y", "N")
            if ($choice -in "yes", "y", "Y") {
            KillApp
            DownloadTeams
            TeamsUninstall
            ClearCache -cacheType teams
            ClassicInstall
            StartApp
            }
        }
    }

    Switch ($Options) {
        TeamsAddinFix {
            KillApp
            TeamsAddinFix
        }
    }
    Read-Host "Press Enter to return"
    break
}

function ShowServiceMenu {
                $MainTitle = "Ørsted SD Script Tool"
            $MainMenuTitle =  "[MS Teams Options]"
            $Menu1 = "MS Teams Re-Deploy"
            # $Menu1SubMenuTitle = "[MS Teams Options]"
            #     $Menu1Option1 = "MS Teams Re-Deploy (Default)"
            #     $Menu1Option2 = "MS Teams Re-Deploy (Classic)"
            #     $Menu1Option3 = "MS Teams Re-Deploy (BootStrap)"
            #     $Menu1Option4 = "Fix MS Teams Addins Missing in Outlook"


            # $Menu2 = "MS Teams Re-Deploy (Classic)"
            # $Menu3 = "MS Teams Re-Deploy (BootStrap)"
            $Menu2 = "Fix MS Teams Add-ins Missing in Outlook"
            # $Menu2SubMenuTitle = "[Menu2SubMenuTitle]"
            #     $Menu2Option1 = "Menu2Option1"
            #     $Menu2Option2 = "Menu2Option2"
    $quit = $false
    # $pswho = $env:USERNAME
    # Write-Host "Enter your admin account for Active Directory. This will be use as the credentials to perform password reset." -ForegroundColor Cyan
    # $AdmCredential = Get-AdmCred
    
    while (-not $quit) {
        Clear-Host
        Write-Host -foregroundcolor White "`n`t`t $MainTitle`n"
        # Write-Host -foregroundcolor White "Welcome $pswho"
        Write-Host -foregroundcolor Cyan `n"$MainMenuTitle" 
        Write-Host -foregroundcolor White -NoNewline "`n["; Write-Host -foregroundcolor Cyan -NoNewline "1"; Write-Host -foregroundcolor White -NoNewline "]"; `
        Write-Host -foregroundcolor White " $Menu1"
        Write-Host -foregroundcolor White -NoNewline "`n["; Write-Host -foregroundcolor Cyan -NoNewline "2"; Write-Host -foregroundcolor White -NoNewline "]"; `
        Write-Host -foregroundcolor White " $Menu2"
        # Write-Host -foregroundcolor White -NoNewline "`n["; Write-Host -foregroundcolor Cyan -NoNewline "3"; Write-Host -foregroundcolor White -NoNewline "]"; `
        # Write-Host -foregroundcolor White " $Menu3"
        # Write-Host -foregroundcolor White -NoNewline "`n["; Write-Host -foregroundcolor Cyan -NoNewline "4"; Write-Host -foregroundcolor White -NoNewline "]"; `
        # Write-Host -foregroundcolor White " $Menu4"
        Write-Host -foregroundcolor White -NoNewline "`n["; Write-Host -foregroundcolor Cyan -NoNewline "Q"; Write-Host -foregroundcolor White -NoNewline "]"; `
        Write-Host -foregroundcolor White " Quit"
        Write-Host
        $choice = Read-Host "Enter Selection [1]-[2] or press Q to quit"
    
        switch ($choice) {
            # '1' {
            #     $Menu1SubMenu = $true
            #     while ($Menu1SubMenu) {
            #         Clear-Host
            #         Write-Host -foregroundcolor White "`n`t`t $MainTitle`n"
            #         Write-Host -foregroundcolor Cyan `n"$Menu1SubMenuTitle" 
            #         Write-Host -foregroundcolor White -NoNewline "`n["; Write-Host -foregroundcolor Cyan -NoNewline "1"; Write-Host -foregroundcolor White -NoNewline "]"; `
            #         Write-Host -foregroundcolor White " $Menu1Option1"
            #         Write-Host -foregroundcolor White -NoNewline "`n["; Write-Host -foregroundcolor Cyan -NoNewline "2"; Write-Host -foregroundcolor White -NoNewline "]"; `
            #         Write-Host -foregroundcolor White " $Menu1Option2"
            #         Write-Host -foregroundcolor White -NoNewline "`n["; Write-Host -foregroundcolor Cyan -NoNewline "3"; Write-Host -foregroundcolor White -NoNewline "]"; `
            #         Write-Host -foregroundcolor White " $Menu1Option3"
            #         Write-Host -foregroundcolor White -NoNewline "`n["; Write-Host -foregroundcolor Cyan -NoNewline "4"; Write-Host -foregroundcolor White -NoNewline "]"; `
            #         Write-Host -foregroundcolor White " $Menu1Option4"
            #         Write-Host -foregroundcolor White -NoNewline "`n["; Write-Host -foregroundcolor Cyan -NoNewline "Q"; Write-Host -foregroundcolor White -NoNewline "]"; `
            #         Write-Host -foregroundcolor White " Back to Main Menu"
            #         Write-Host
            #         $Menu1SubmenuChoice = Read-Host "Enter Select 1-2 or press Q to go back to Main Menu"
    
            #         switch ($Menu1SubmenuChoice) {
            #             '1' {
            #                 Clear-Host
            #                 Write-Host -foregroundcolor White "`n`t`t $MainTitle`n"
            #                 # Write-Host "Menu1Choice1"
            #                 MSTeamsReinstallFull -DeploymentType MSIX
            #                 $prompt = Read-Host "Type Q to go back to $Menu1 "
            #                 if ($prompt -eq 'Q') {
            #                     continue
            #                 }
            #             }
            #             '2' {
            #                 Clear-Host
            #                 Write-Host -foregroundcolor White "`n`t`t $MainTitle`n"
            #                 MSTeamsReinstallFull -DeploymentType Classic
            #                 $prompt = Read-Host "Type Q to go back to $Menu1 "
            #                 if ($prompt -eq 'Q') {
            #                     continue
            #                 }
            #             }
            #             '3' {
            #                 Clear-Host
            #                 Write-Host -foregroundcolor White "`n`t`t $MainTitle`n"
            #                 MSTeamsReinstallFull -DeploymentType BootStrap
            #                 $prompt = Read-Host "Type Q to go back to $Menu1 "
            #             }
            #             '4' {
            #                 Clear-Host
            #                 Write-Host -foregroundcolor White "`n`t`t $MainTitle`n"
            #                 MSTeamsReinstallFull -Options TeamsAddinFix
            #                 $prompt = Read-Host "Type Q to go back to $Menu1 "
            #             }
            #             'Q' {
            #                 $Menu1SubMenu = $false
            #                 $quit = $false
            #             }
            #         }
            #     }
            # }
            # '2' {
            #     $Menu2SubMenu = $true
            #     while ($Menu2SubMenu) {
            #         Clear-Host
            #         Write-Host -foregroundcolor White "`n`t`t $MainTitle`n"
            #         Write-Host -foregroundcolor Cyan "$Menu2SubMenuTitle" 
            #         Write-Host -foregroundcolor White -NoNewline "`n["; Write-Host -foregroundcolor Cyan -NoNewline "1"; Write-Host -foregroundcolor White -NoNewline "]"; `
            #         Write-Host -foregroundcolor White " $Menu2Option1"
            #         Write-Host -foregroundcolor White -NoNewline "`n["; Write-Host -foregroundcolor Cyan -NoNewline "2"; Write-Host -foregroundcolor White -NoNewline "]"; `
            #         Write-Host -foregroundcolor White " $Menu2Option2"
            #         Write-Host -foregroundcolor White -NoNewline "`n["; Write-Host -foregroundcolor Cyan -NoNewline "Q"; Write-Host -foregroundcolor White -NoNewline "]"; `
            #         Write-Host -foregroundcolor White " Back to Main Menu"
            #         Write-Host
            #         $Menu2SubmenuChoice = Read-Host "Enter Select 1-2 or press Q to go back to Main Menu"

            #         switch ($Menu2SubmenuChoice) {
            #             '1' {
            #                 Clear-Host
            #                 Write-Host -foregroundcolor White "`n`t`t $MainTitle`n"
            #                 Write-Host "Menu2Choice1"
            #                 $prompt = Read-Host "Type Q to go back to $Menu2 "
            #                 if ($prompt -eq 'Q') {
            #                     continue
            #                 }
            #             }
            #             '2' {
            #                 Clear-Host
            #                 Write-Host -foregroundcolor White "`n`t`t $MainTitle`n"
            #                 Write-Host "Menu2Choice2"
            #                 $prompt = Read-Host "Type Q to go back to $Menu2 "
            #                 if ($prompt -eq 'Q') {
            #                     continue
            #                 }
            #             }
            #             # '3' {
            #             #     $Menu2SubMenu = $false
            #             # }
            #             'Q' {
            #                 $Menu2SubMenu = $false
            #                 $quit = $false
            #             }
            #         }
            #     }
            # }
            '1' {
                Clear-Host
                Write-Host -foregroundcolor White "`n`t`t $MainTitle`n"
                # Write-Host "Menu1Choice1"
                MSTeamsReinstallFull -DeploymentType MSIX -cacheType teams
                $prompt = Read-Host "Type Q to go back to $Menu1 "
                if ($prompt -eq 'Q') {
                    continue
                }
            }
            # '2' {
            #     Clear-Host
            #     Write-Host -foregroundcolor White "`n`t`t $MainTitle`n"
            #     MSTeamsReinstallFull -DeploymentType Classic -cacheType teams
            #     $prompt = Read-Host "Type Q to go back to $Menu1 "
            #     if ($prompt -eq 'Q') {
            #         continue
            #     }
            # }
            # '3' {
            #     Clear-Host
            #     Write-Host -foregroundcolor White "`n`t`t $MainTitle`n"
            #     MSTeamsReinstallFull -DeploymentType BootStrap -cacheType teams
            #     $prompt = Read-Host "Type Q to go back to $Menu1 "
            # }
            '2' {
                Clear-Host
                Write-Host -foregroundcolor White "`n`t`t $MainTitle`n"
                MSTeamsReinstallFull -Options TeamsAddinFix
                $prompt = Read-Host "Type Q to go back to $Menu1 "
            }
            'Q' {
                $quit = $true
            }
        }
    }
    
    Write-Host "Goodbye!"
}
ShowServiceMenu