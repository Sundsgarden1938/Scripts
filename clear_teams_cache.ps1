#Set Variables
$TeamsCacheFolder = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe"
$CacheSubfolders = @(get-childitem -Path $TeamsCacheFolder)
$TeamsRunning = (Get-Process -Name "ms-teams" -ErrorAction Ignore).ProcessName
$TeamsApp = "$env:LOCALAPPDATA\Microsoft\WindowsApps\ms-teams.exe"
$logfolder = "C:\Software\MSTeams"
$logfile = "$logfolder\ClearCache.log" 

#Check if log file exists, clear content if it does
If(test-path $logfile)
{
Clear-Content $logfile
}
#Check if log folder exists, create it if not
If(!(test-path $logfolder))
{
New-Item -ItemType Directory -Force -Path $logfolder
}   
#Check if log file exists, create it if not
If(!(test-path $logfile))
{
New-Item $logfile
}

Start-Transcript -Path $logfile

#Stop Teams process if it is running
If($null -ne $TeamsRunning)
{
Stop-Process -Name $TeamsRunning -Force
}

#Wait until Teams process is no longer running
while($null -ne (Get-Process -Name "ms-teams" -ErrorAction Ignore))
{
sleep 1
}

#Remove the contents of the cache folder
foreach($Subfolder in $CacheSubfolders)
{
Remove-Item -Path "$TeamsCacheFolder\$Subfolder" -Force -Recurse
}

#Stop logging and add final comment with current date and time
Stop-Transcript
$datetime = (Get-Date)
Add-Content $logfile "$datetime - MS Teams Cache Clear script has completed"

#Wait 3 seconds
Sleep 3
#Restart Teams
Start-Process $TeamsApp