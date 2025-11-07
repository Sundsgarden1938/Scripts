$appName = "Microsoft Teams classic"
# A function to search for Teams installation
function Get-MSTeams {
    $path = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", 
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", 
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    $entry = Get-ChildItem -Path $path | Get-ItemProperty | where {$_.Displayname -match $appName}
    return $entry
}

# Assign the result to a variable, so you can use .InstallLocation method further down and trigger uninstallation
$app = Get-MSTeams

# Check if the function has returned result (Teams is installed)
if ($app){

# If yes, prompt for action
    $prompt = Read-Host "Microsoft Teams Classic installation FOUND. Uninstall? [Y/N]"
    
    if ($prompt -eq "Y"){
        Set-Location -Path $app.InstallLocation
        Start-Process Update.exe -ArgumentList "--uninstall", "-s"
        Write-Output "`nUninstalling..."
        Start-Sleep 5

# Keep checking if Teams is still installed, until it's not
        while (Get-MSTeams){
            Start-Sleep 2
            Write-Output "Uninstalling..."
        }
        Write-Output "`nMicrosoft Teams Classic successfully uninstalled!`n"
    }
    else {break}
}
else {
    Write-Output "Microsoft Teams Classic installation NOT found!"
}

# Set location to home directory
Set-Location $HOME
