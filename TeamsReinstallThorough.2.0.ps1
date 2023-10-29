<#
    .SYNOPSIS
        MS Teams Deployment for Single PC
    
    .DESCRIPTION
        Reinstall MS Teams.
        Remove MS Teams residual when uninstall MS Teams.
        Download MS Teams using bootstrap.exe, MSIX package, or the classic MS Teams.
        Delete browser cache.
    
    .PARAMETER DeploymentType
        Choose between MSIX, bootstrap or Classic teams to be install.
        Default installation without paramater is bootstrap.
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
    [string]$DeploymentType = "BootStrap"
)
$ErrorActionPreference = "SilentlyContinue"
$ClassicInstall = $DeploymentType -eq "Classic"


$challenge = Read-Host "Are you sure you wish to completely reinstall MS Teams?
`nMS Teams will be removed from this PC and will be replace with New MS Teams.
`nThe default deployment type is BootStrap.
`nTo install MS Teams Classic, use -DeploymentType Classic
`nTo install using MSIX package, use -DeploymentType MSIX
`nThis will also close Internet Explorer, Chrome, Firefox & Edge (Y/N)"
$challenge = $challenge.ToUpper()

# Check if user wrote YES/NO
if ($challenge.Length -gt 1){
    if ($challenge -eq "NO"){
        $challenge = "N"
    }
    
    elseif ($challenge -eq "YES"){
        $challenge = "Y"
    }

    else{
    }
}

if ($challenge -eq "N"){
    Stop-Process -Id $PID
}

elseif ($challenge -eq "Y"){
    

    #Region Kill Process ===============================================================
    #Stops Microsoft Teams
    Write-Host "Deplyoment type is $DeploymentType" -ForegroundColor Green
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

    

    #EndRegion Kill Process ============================================================

    #Region Backup TeamsMeetingAddinDLL ================================================

    if (!$ClassicInstall) {
        $sourcePath = "$env:LOCALAPPDATA\Microsoft\TeamsMeetingAddin"
        $backupPath = "$env:LOCALAPPDATA\Microsoft"
        $DirExist = Test-Path $sourcePath
        if (!$DirExist) {
            Write-Host "Error!`n$sourcePath does not exist. `nPlease reinstall MS Teams Classic
            `n To install MS Teams Classic, use the script in powershell and execute with parameter -DeploymentType Classic" -ForegroundColor Red
            Read-Host "Press enter to exit"
            return
        }
    
        $sourceItems = Get-ChildItem $sourcePath
        $DirEmpty = $sourceItems.Count -eq 0
        if ($DirEmpty) {
            Write-Host "Error!`n$sourcePath is empty. `nPlease reinstall MS Teams Classic.
            `n To install MS Teams Classic, use the script in powershell and execute with parameter -DeploymentType Classic" -ForegroundColor Red
            Read-Host "Press enter to exit"
            return
        }
    
        if (-not (Test-Path $backupPath)) {
            Write-Host "Error!`nBackup directory $backupPath does not exist."
            Read-Host "Press enter to exit"
            return
        }
    
        $backupDestination = Join-Path $backupPath "TeamsMeetingAddinBackup"
    
        # Check if the destination folder already exists
        if (Test-Path $backupDestination) {
            Write-Host "$backupDestination already exists. Proceed to remove it."  -ForegroundColor Red
            Remove-item -Path $backupDestination -Recurse
            # Read-Host "Press enter to exit"
            # return
        }
    
        Copy-Item -Path $sourcePath -Destination $backupDestination -Recurse
        Write-Host "Backup of $sourcePath created in $backupDestination" -ForegroundColor Green
    }
   
    #EndRegion Backup TeamsMeetingAddinDLL =============================================
    
    #Region MS Teams Uninstall =========================================================
    #Classic MS Teams Uninstall
    $TeamsPath = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Microsoft', 'Teams')
    $TeamsUpdateExePath = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Microsoft', 'Teams', 'Update.exe')
    try
    {
        if (Test-Path -Path $TeamsUpdateExePath) {
            Write-Host "Uninstalling Teams process"
            # Uninstall app
            $proc = Start-Process -FilePath $TeamsUpdateExePath -ArgumentList "-uninstall -s" -PassThru
            $proc.WaitForExit()
        }
        if (Test-Path -Path $TeamsPath) {
            Write-Host "Deleting Teams directory"
            Remove-Item -Path $TeamsPath -Recurse
        }
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
    #New MSTeams uninstall
    $NewMSTeams = Get-AppxPackage -Name MSTeams
    try {
        if ($NewMSTeams){
            Remove-AppxPackage $NewMSTeams
        }
    }
    catch {
        Write-Error -ErrorRecord $_
    }
    #EndRegion MS Teams Uninstall ======================================================

    #Region ClearCache =================================================================
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
    #Endregion cleacache ===============================================================


    #Region DeploymentType =============================================================
    $InstallerDir = "$ENV:USERPROFILE\Downloads"
    switch ($DeploymentTYpe) {
        MSIX {
            $DownloadSource = "https://go.microsoft.com/fwlink/?linkid=2196060&clcid=0x409&culture=en-us&country=us" #MSTeams-x86.msix,32bit,no elevation
            $InstallerLocation = "$InstallerDir\MSTeams-x86.msix"
            If([System.IO.File]::Exists($InstallerLocation) -eq $false){
                Write-Host "Downloading Teams, please wait." -ForegroundColor Red
                curl.exe -fSLo $InstallerLocation $DownloadSource # 10 second download (with progress bar)
            }
            Else{
                Write-Host "Installer file already present in Downloads folder. Skipping download." -ForegroundColor Yellow
            }
            Write-Host "Installing Teams" -ForegroundColor Magenta
            try {
                Add-AppxPackage -Path $InstallerLocation
            }
            catch {
                Write-Host "AppX error!" -ForegroundColor Red
                Read-Host "Press enter to exit"
                return  
            }
        }
        BootStrap {
            $DownloadSource = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409" #teamsbootstrapper.exe,64bit,require elevation
            $InstallerLocation = "$InstallerDir\teamsbootstrapper.exe"
            If([System.IO.File]::Exists($InstallerLocation) -eq $false){
                Write-Host "Downloading Teams, please wait." -ForegroundColor Red
                curl.exe -fSLo $InstallerLocation $DownloadSource # 10 second download (with progress bar)
            }
            Else{
                Write-Host "Installer file already present in Downloads folder. Skipping download." -ForegroundColor Yellow
            }
        
            Write-Host "Installing Teams" -ForegroundColor Magenta
            try {
                Unblock-File -Path $InstallerLocation
                $proc = Start-Process -FilePath $InstallerLocation -ArgumentList "-p" -PassThru
            }
            catch {
                Write-Host "Elevation or EPM error. Script stop" -ForegroundColor Red
                Read-Host "Press enter to exit"
                return  
            }
            # Add-AppxPackage -Path $InstallerLocation
            $proc.WaitForExit()
        }
        Classic {
            $DownloadSource = "https://go.microsoft.com/fwlink/?linkid=2187327"
            $InstallerLocation = "$InstallerDir\TeamsSetup_c_w_.exe"
            If([System.IO.File]::Exists($InstallerLocation) -eq $false){
                Write-Host "Downloading Teams, please wait." -ForegroundColor Red
                curl.exe -fSLo $InstallerLocation $DownloadSource # 10 second download (with progress bar)
                # $ProgressPreference = 'SilentlyContinue'
                # Invoke-WebRequest $DownloadSource -OutFile $InstallerLocation # 6 minutes 11 seconds download (with progress bar, 11 second no progress bar)
                # $ProgressPreference = 'Continue'
                # $wc = New-Object Net.Webclient
                # $wc.DownloadFile($DownloadSource,$InstallerLocation) # 11 second download (no progress bar)
                # Unblock-File -Path $InstallerLocation
            }
            Else{
                Write-Host "Installer file already present in Downloads folder. Skipping download." -ForegroundColor Yellow
            }
        
            Write-Host "Installing Teams" -ForegroundColor Magenta
            try {
                $proc = Start-Process -FilePath $InstallerLocation -ArgumentList "-s" -PassThru
            }
            catch {
                Write-Host "Elevation or EPM error. Script stop" -ForegroundColor Red
                Read-Host "Press enter to exit"
                return  
            }
            $proc.WaitForExit()
        }
    }
    #EndRegion DeploymentType ==========================================================

    #Region Restore TeamsAddin backup ==================================================
    if (!$ClassicInstall){
        Write-Host "Restoring TeamsAddinDLL backup" -ForegroundColor Green
        if (Test-Path $sourcePath) {
            Remove-Item $sourcePath -Recurse -Force
        }
        Rename-Item -path $backupDestination -NewName $sourcePath
    }

    #EndRegion Restore TeamsAddin backup ===============================================
    
    #Region Registring Teams Addin ====================================================
    if (!$ClassicInstall) {
        $TeamsMeetingAddinPath= "$env:LOCALAPPDATA\Microsoft\TeamsMeetingAddin"

        if (-not (Test-Path $TeamsMeetingAddinPath)) {
            Write-Host "Warning!! $TeamsMeetingAddinPath does not exist" -ForegroundColor Red
            Write-Host "Please reinstall MS Teams Classic" -ForegroundColor Red
            Read-Host "Press Enter to exit"
            Exit
        } 
        else {
            $items = Get-ChildItem $TeamsMeetingAddinPath
            if ($items.Count -eq 0) {
                Write-Host "Warning!! $TeamsMeetingAddinPath is empty" -ForegroundColor Red
                Write-Host "Please reinstall MS Teams Classic" -ForegroundColor Red
                Read-Host "Press Enter to exit"
                Exit
            }
        }
        # Register TeamsAddin DLL   
        $LattestDLLversion = (Get-ChildItem -Path $TeamsMeetingAddinPath-Directory |Sort-Object CreationTime -Descending| Select-Object -First 1)
        $LattestDLLversion = ($LattestDLLversion).FullName
        $teamsdotdead = "$LattestDLLversion\.dead"
        $teamsdll = "$LattestDLLversion\x64\Microsoft.Teams.AddinLoader.dll"
                
        Write-Host "Removing .dead file if exist"
        if (Test-Path -Path $teamsdotdead) {
            Remove-Item -Path $teamsdotdead
            Write-host ".dead file found and removed"
        }
        else {
            Write-host "No .dead file exist"
        }
        write-host "Deregistring Microsoft.Teams.AddinLoader.dll" -ForegroundColor Yellow
        start-sleep 5
        regsvr32.exe /U "$teamsdll" /s
        write-host "Registring Microsoft.Teams.AddinLoader.dll" -ForegroundColor Green
        regsvr32.exe /n /i:user "$teamsdll" /s
        Write-Host "Done"
        
        # Check if Microsoft Teams add-ins for Outlook are enabled
        $TeamsMeetingAddinRegPath = "HKCU:\SOFTWARE\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect" 
        $FastConnectReg = Get-Item -Path $TeamsMeetingAddinRegPath -ErrorAction SilentlyContinue
        
        if ($null -eq $FastConnectReg) {
            Write-Host "Microsoft Teams add-ins for Outlook is not enable"
            Write-Host "Enabling Teams Addin in Outlook"
            New-Item -Path $TeamsMeetingAddinRegPath
            New-ItemProperty -Path $TeamsMeetingAddinRegPath -Name "Description" -Value "Microsoft Teams Meeting Add-in for Microsoft Office"
            New-ItemProperty -Path $TeamsMeetingAddinRegPath -Name "FriendlyName" -Value "Microsoft Teams Meeting Add-in for Microsoft Office"
            New-ItemProperty -Path $TeamsMeetingAddinRegPath -Name "LoadBehavior" -PropertyType DWord -Value 3
            Write-Host "Teams Addins Enabled" -ForegroundColor Green
        } 
        else {
            $CurLoadBehavior = $FastConnectReg.GetValue("LoadBehavior")
            if ($CurLoadBehavior -eq 3) {
                Write-Host "Microsoft Teams add-ins LoadBehavior is already set to 3."
            } else {
                Write-Host "Microsoft Teams add-ins LoadBehavior is $CurLoadBehavior"
                Set-ItemProperty -path $TeamsMeetingAddinRegPath -Name LoadBehavior -Value 3
                $newloadbehavior = $FastConnectReg.GetValue("LoadBehavior")
                Write-Host "Microsoft Teams add-ins LoadBehavior has been set to $newloadbehavior."
            }
        }
        
        # KB0016283 - Add registry entry (if not exist) - KB from Dina Rantzau https://onewebshop.service-now.com/kb_view.do?sysparm_article=KB0016283
        Write-Host "Applying KB0016283"
        # RTAC1
        $RTAC1P = 'HKCU:\software\Policies\Microsoft\office\16.0\outlook\resiliency\addinlist\'
        $RTAC1 = Get-ItemPropertyValue -Path "HKCU:\software\Policies\Microsoft\office\16.0\outlook\resiliency\addinlist\" -Name TeamsAddin.Connect
        
        if ($null -eq $RTAC1) {
            Write-Host "Create RTAC1"
            New-ItemProperty -Path $RTAC1P -Name "TeamsAddin.FastConnect" -Value "1" -PropertyType "String"
        }
        else {
            Write-host "RTAC1 already exist"
        }
        #RTAC2
        $RTAC2P = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList\'
        $RTAC2 = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList\" -Name TeamsAddin.Connect
        
        if ($null -eq $RTAC2) {
            Write-Host "Create RTAC2"
            New-ItemProperty -Path $RTAC2P -Name "TeamsAddin.FastConnect" -Value "1" -PropertyType DWord
        }
        else {
            Write-host "RTAC2 already exist"
        }
        Write-Host "Done Registring MS TeamsAddin!"
    }
    #EndRegion Registring Teams Addin =================================================

    Write-Host "Starting MS Teams.." -ForegroundColor Green

    Switch -regex ($DeploymentType) {
        "BootStrap|MSIX" {
            Start-Process -FilePath $env:LOCALAPPDATA\Microsoft\WindowsApps\ms-teams.exe
        }
        "Classic" {
            Start-Process -FilePath $env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe -PassThru
        }
    }
    Write-Host "Done!" -ForegroundColor Green
    Start-Sleep 3
    Read-Host "Press Enter to exit.."
    Stop-Process -Id $PID
}

else{
    Write-Host "Not a valid input, stopping script"
    Start-Sleep -s 6
    Stop-Process -Id $PID
}