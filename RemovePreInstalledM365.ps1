#$M365Packages = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "*Microsoft 365*"} 
$M365Packages = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "*Microsoft 365 - *"} 
if ($M365Packages.Count -gt 1) 
{ 
    write-host "Starting Uninstallation..." -ForegroundColor Cyan 
    foreach ($M365Package in $M365Packages) 
    { 
        write-host "Removing $($M365Package.DisplayName)" -ForegroundColor Yellow 
        $UninstallString = $M365Package.UninstallString 
        $UninstallEXE = ($UninstallString -split '"')[1] 
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False" 
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait 
    } 
}
else
{
    write-host "None found." -ForegroundColor Green 
}
