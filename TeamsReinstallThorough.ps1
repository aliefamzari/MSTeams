##### Script based on the great works of XSIOL. #####
##### Formatting, commenting, modifications and merging of the scripts done by TOBKO. #####

##################################
##################################
#####      Initiation        #####
#### By XSIOL & TOBKO & ALMAZ ####
##################################
##################################

$ErrorActionPreference = "SilentlyContinue"


$challenge = Read-Host "Are you sure you wish to completely reinstall Microsoft Teams? 
This will also close Internet Explorer, Chrome, Firefox & Edge (Y/N)"
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
    
    ##################################
    ##################################
    ##### Teams Uninstall Script #####
    #####        By XSIOL        #####
    ##################################
    ##################################
    
    # Stops Microsoft Teams
    Write-Host "Stopping Teams Process" -ForegroundColor Yellow

    try{
        Get-Process -ProcessName Teams  | Stop-Process -Force
        Start-Sleep -Seconds 3
        Write-Host "Teams Process Sucessfully Stopped" -ForegroundColor Green
    }
    
    catch{
        echo $_
    }

    # Starts uninstall
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
    
    ##############################
    ##############################
    ##### Cache Clear Script #####
    #####  By XSIOL & TOBKO  #####
    ##############################
    ##############################
    
    Write-Host "Clearing Teams Disk Cache" -ForegroundColor Yellow
    
    try{
        Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\application cache\cache"  | Remove-Item -Confirm:$false 
        Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\blob_storage"  | Remove-Item -Confirm:$false 
        Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\databases"  | Remove-Item -Confirm:$false 
        Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\cache"  | Remove-Item -Confirm:$false 
        Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\gpucache"  | Remove-Item -Confirm:$false 
        Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\Indexeddb"  | Remove-Item -Confirm:$false 
        Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\Local Storage"  | Remove-Item -Confirm:$false 
        Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\tmp"  | Remove-Item -Confirm:$false 
        Write-Host "Teams Disk Cache Cleaned" -ForegroundColor Green
    }
    
    catch{
        echo $_
    }
    
    Write-Host "Stopping Chrome Process" -ForegroundColor Yellow

    try{
        Get-Process -ProcessName Chrome  | Stop-Process -Force 
        Start-Sleep -Seconds 3
        Write-Host "Chrome Process Sucessfully Stopped" -ForegroundColor Green
    }
    
    catch{
        echo $_
    }

    Write-Host "Clearing Chrome Cache" -ForegroundColor Yellow
    
    try{
        Get-ChildItem -Path $env:LOCALAPPDATA"\Google\Chrome\User Data\Default\Cache"  | Remove-Item -Confirm:$false 
        Get-ChildItem -Path $env:LOCALAPPDATA"\Google\Chrome\User Data\Default\Cookies" -File  | Remove-Item -Confirm:$false 
        Get-ChildItem -Path $env:LOCALAPPDATA"\Google\Chrome\User Data\Default\Web Data" -File  | Remove-Item -Confirm:$false 
        Write-Host "Chrome Cleaned" -ForegroundColor Green
    }
    
    catch{
        echo $_
    }
    
    Write-Host "Stopping IE & Edge Process" -ForegroundColor Yellow
    
    try{
        Get-Process -ProcessName MicrosoftEdge  | Stop-Process -Force 
        Get-Process -ProcessName MSEdge  | Stop-Process -Force 
        Get-Process -ProcessName IExplore  | Stop-Process -Force 
        Write-Host "Internet Explorer and Edge Processes Sucessfully Stopped" -ForegroundColor Green
    }
    
    catch{
        echo $_
    }

    Write-Host "Clearing IE & Edge Cache" -ForegroundColor Yellow
    
    try{
        RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 8
        RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 2
        Get-ChildItem -Path $env:LOCALAPPDATA"\Microsoft\Edge\User Data\Default\Cache"  | Remove-Item -Confirm:$false 
        Get-ChildItem -Path $env:LOCALAPPDATA"\Google\Chrome\User Data\Default\Cookies" -File  | Remove-Item -Confirm:$false 
        Get-ChildItem -Path $env:LOCALAPPDATA"\Google\Chrome\User Data\Default\Web Data" -File  | Remove-Item -Confirm:$false 
        Start-Sleep 3
        Write-Host "IE and Edge Cleaned" -ForegroundColor Green
    }
    
    catch{
        echo $_
    }

    Write-Host "Stopping Firefox Process" -ForegroundColor Yellow

    try{
        Get-Process -ProcessName Firefox  | Stop-Process -Force 
        Start-Sleep -Seconds 3
        Write-Host "Firefox Process Sucessfully Stopped" -ForegroundColor Green
    }
    
    catch{
        echo $_
    }

    Write-Host "Clearing Firefox Cache" -ForegroundColor Yellow
    
    try{
        Get-ChildItem -Path $env:LOCALAPPDATA"\Mozilla\Firefox\Profiles\"  | Remove-Item -Confirm:$false -Recurse 

        Write-Host "Firefox Cleaned" -ForegroundColor Green
    }

    catch{
        echo $_
    }

    Write-Host "Cleanup Complete..." -ForegroundColor Green


    ##################################
    ##################################
    ##### Teams Reinstall Script #####
    #####        By TOBKO        ##### 
    ##################################
    ##################################
    
    # Make a Teams install that automatically downloads the installer
    $ExeFolder = "$ENV:USERPROFILE\Downloads"
    $DownloadSource = "https://go.microsoft.com/fwlink/p/?LinkID=869426&clcid=0x409&culture=en-us&country=US&lm=deeplink&lmsrc=groupChatMarketingPageWeb&cmpid=directDownloadWin64"
    $ExeDestination = "$ExeFolder\Teams_windows_x64.exe"

    If([System.IO.File]::Exists($ExeDestination) -eq $false){
        Write-Host "Downloading Teams, please wait." -ForegroundColor Red
        Invoke-WebRequest $DownloadSource -OutFile $ExeDestination
    }
    Else{
        Write-Host "Installer file already present in Downloads folder. Skipping download." -ForegroundColor Red
    }

    Write-Host "Installing Teams" -ForegroundColor Magenta
    $proc = Start-Process -FilePath $ExeDestination -ArgumentList "-s" -PassThru
    $proc.WaitForExit()

    Start-Sleep 5

    Write-Host "Checking install" -ForegroundColor Magenta
    $proc = Start-Process -FilePath $ExeDestination -ArgumentList "-s" -PassThru
    $proc.WaitForExit()

    Start-Process -FilePath $env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe
    
    Stop-Process -Id $PID
}

else{
    Write-Host "Not a valid input, stopping script"
    Start-Sleep -s 6
    Stop-Process -Id $PID
}