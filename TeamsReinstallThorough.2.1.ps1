<#
    .SYNOPSIS
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
        Git Control: Azure DevOps
        Project URL: https://dev.azure.com/ALMAZ0773/Teams%20Reinstall
        Repository: https://dev.azure.com/ALMAZ0773/_git/Teams%20Reinstall
        Author: Alif Amzari, ALMAZ
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
param (
    [ValidateSet("MSIX","BootStrap","Classic")]
    [string]$DeploymentType = "MSIX",
    [switch]$byPassSSL
)
$ErrorActionPreference = "SilentlyContinue"

Write-Host "Are you sure you wish to completely reinstall MS Teams?"
Write-Host "The default deployment will removed the Classic Teams in favor of New Teams."
Write-Host "The default deployment type is MSIX."
Write-Host "To install MS Teams Classic, use -DeploymentType Classic" 
Write-Host "To install using BootStrap package, use -DeploymentType BootStrap"
Write-Host "This will delete cache of Internet Explorer, MS Edge, Chrome & Firefox"
Write-Host "Please ensure you save and close your Outlook and browsers." -ForegroundColor Red
# Display the current deployment type in green with a newline
Write-Host -NoNewline "The current deployment is "
Write-Host -NoNewline -ForegroundColor Green "$DeploymentType"
Write-Host

$InstallerDir = "$ENV:USERPROFILE\Downloads"
$TeamsMeetingAddinDir = "$env:LOCALAPPDATA\Microsoft\TeamsMeetingAddin"
$backupPath = "$env:LOCALAPPDATA\Microsoft"
$backupDestination = Join-Path $backupPath "TeamsMeetingAddinBackup"

do {
    $choice = Read-Host "Do you want to continue? (yes/no)"
} while ($choice -notin "yes", "no", "y", "n", "Y", "N")

if ($choice -in "yes", "y", "Y") {
    #Region Function
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
        $proc.WaitForExit()
    }
    function StartApp {
        if ($DeploymentType -in "MSIX", "BootStrap") {
            $AppPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps\ms-teams.exe"
        } else {
            $AppPath = "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe"
        }
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
    function TeamsAddin {
        if (-not (Test-Path $TeamsMeetingAddinDir)) {
            Write-Host "Warning!! $TeamsMeetingAddinDir does not exist" -ForegroundColor Red
            Write-Host "Please reinstall MS Teams Classic" -ForegroundColor Red
            #Read-Host "Press enter to exit"
            Exit
        } 
        else {
            $items = Get-ChildItem $TeamsMeetingAddinDir
            if ($items.Count -eq 0) {
                Write-Host "Warning!! $TeamsMeetingAddinDir is empty" -ForegroundColor Red
                Write-Host "Please reinstall MS Teams Classic" -ForegroundColor Red
                #Read-Host "Press enter to exit"
                Exit
            }
        }
        # Register TeamsAddin DLL   
        $LattestDLLversion = (Get-ChildItem -Path $TeamsMeetingAddinDir -Directory |Sort-Object CreationTime -Descending| Select-Object -First 1)
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
        
        # KB0016283 - Add registry entry (if not exist) - KB from Dina Rantzau https://onewebshop.service-now.com/kb_view.do?sysparm_article=KB0016283
        Write-Host "Applying KB0016283" -ForegroundColor Yellow
        # ResiliencyTeamsAddinConnect1
        $ResiliencyTeamsAddinPath1 = 'HKCU:\software\Policies\Microsoft\office\16.0\outlook\resiliency\addinlist\'
        $ResiliencyTeamsAddinConnect1 = Get-ItemPropertyValue -Path "HKCU:\software\Policies\Microsoft\office\16.0\outlook\resiliency\addinlist\" -Name TeamsAddin.Connect
        $ResiliencyTeamsAddinFastConnect1 = Get-ItemPropertyValue -Path "HKCU:\software\Policies\Microsoft\office\16.0\outlook\resiliency\addinlist\" -Name TeamsAddin.FastConnect

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
        $ResiliencyTeamsAddinPath2 = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList\'
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
        Write-Host "Done Registring MS TeamsAddin!" -ForegroundColor Green
    }
    function BackupTeamsAddin {
            $DirExist = Test-Path $TeamsMeetingAddinDir
            if (!$DirExist) {
                Write-Host "Error!`n$TeamsMeetingAddinDir does not exist. `nPlease reinstall MS Teams Classic
                `n To install MS Teams Classic, use the script in powershell and execute with parameter -DeploymentType Classic" -ForegroundColor Red
                # Read-Host "Press enter to exit"
                Break
            }
        
            $sourceItems = Get-ChildItem $TeamsMeetingAddinDir
            $DirEmpty = $sourceItems.Count -eq 0
            if ($DirEmpty) {
                Write-Host "Error!`n$TeamsMeetingAddinDir is empty. `nPlease reinstall MS Teams Classic.
                `n To install MS Teams Classic, use the script in powershell and execute with parameter -DeploymentType Classic" -ForegroundColor Red
                # Read-Host "Press enter to exit"
                Break
            }
        
            if (-not (Test-Path $backupPath)) {
                Write-Host "Error!`nBackup directory $backupPath does not exist."
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
        function CleanCache($path) {
            if (Test-Path $path) {
                Get-ChildItem -Path $path | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            }
        }
        
        function StopProcess($processName) {
            Get-Process -Name $processName -ErrorAction SilentlyContinue| Stop-Process -Force -ErrorAction SilentlyContinue
        }
        
        Write-Host "Clearing Teams Disk Cache" -ForegroundColor Yellow
        "application cache\cache", "blob_storage", "databases", "cache", "gpucache", "Indexeddb", "Local Storage", "tmp" | ForEach-Object { CleanCache "$env:APPDATA\Microsoft\teams\$_" }
        Write-Host "Teams Disk Cache Cleaned" -ForegroundColor Green
        
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
                curl.exe -fSLo $InstallerLocation $DownloadSource
            }
        }
        switch ($DeploymentType) {
            MSIX {
                $DownloadSource = "https://go.microsoft.com/fwlink/?linkid=2196060&clcid=0x409" #MSTeams-x86.msix,32bit,no elevation
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
    #EndRegion Function

    Switch ($DeploymentType) {
        'BootStrap' {
            KillApp
            BackupTeamsAddin
            DownloadTeams
            TeamsUninstall
            ClearCache
            BootStrapInstall
            RestoreTeamsAddinBackup
            TeamsAddin
            StartApp
        }
        'MSIX' {
            KillApp
            BackupTeamsAddin
            DownloadTeams
            TeamsUninstall
            ClearCache
            MSIXInstall
            RestoreTeamsAddinBackup
            TeamsAddin
            StartApp
        }
        'Classic' {
            KillApp
            DownloadTeams
            TeamsUninstall
            ClearCache
            ClassicInstall
            StartApp
        }
    }
    Read-Host "Press Enter to exit"
    break
}
else 
{
    Write-Host "You chose 'no'."
}
