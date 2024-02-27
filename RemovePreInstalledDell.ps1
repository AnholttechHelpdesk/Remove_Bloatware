############################################################################################################
#                                       Grab all Uninstall Strings                                         #
#                                                                                                          #
############################################################################################################
write-host "Checking 32-bit System Registry"
##Search for 32-bit versions and list them
$allstring = @()
$path1 =  "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
#Loop Through the apps if name has Adobe and NOT reader
$32apps = Get-ChildItem -Path $path1 | Get-ItemProperty | Select-Object -Property DisplayName, UninstallString

foreach ($32app in $32apps) 
{
    #Get uninstall string
    $string1 =  $32app.uninstallstring
    #Check if it's an MSI install
    if ($string1 -match "^msiexec*") 
    {
        #MSI install, replace the I with an X and make it quiet
        $string2 = $string1 + " /quiet /norestart"
        $string2 = $string2 -replace "/I", "/X "
        #Create custom object with name and string
        $allstring += New-Object -TypeName PSObject -Property @{
            Name = $32app.DisplayName
            String = $string2
        }
    }
    else 
    {
        #Exe installer, run straight path
        $string2 = $string1
        $allstring += New-Object -TypeName PSObject -Property @{
            Name = $32app.DisplayName
            String = $string2
        }
    }
}
write-host "32-bit check complete"
write-host "Checking 64-bit System registry"
##Search for 64-bit versions and list them

$path2 =  "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
#Loop Through the apps if name has Adobe and NOT reader
$64apps = Get-ChildItem -Path $path2 | Get-ItemProperty | Select-Object -Property DisplayName, UninstallString

foreach ($64app in $64apps) 
{
    #Get uninstall string
    $string1 =  $64app.uninstallstring
    #Check if it's an MSI install
    if ($string1 -match "^msiexec*") 
    {
        #MSI install, replace the I with an X and make it quiet
        $string2 = $string1 + " /quiet /norestart"
        $string2 = $string2 -replace "/I", "/X "
        #Uninstall with string2 params
        $allstring += New-Object -TypeName PSObject -Property @{
            Name = $64app.DisplayName
            String = $string2
        }
    }
    else 
    {
        #Exe installer, run straight path
        $string2 = $string1
        $allstring += New-Object -TypeName PSObject -Property @{
            Name = $64app.DisplayName
            String = $string2
        }
    }
}

write-host "64-bit checks complete"

############################################################################################################
#                                        Remove Manufacturer Bloat                                         #
#                                                                                                          #
############################################################################################################
##Check Manufacturer
write-host "Detecting Manufacturer"
$details = Get-CimInstance -ClassName Win32_ComputerSystem
$manufacturer = $details.Manufacturer


if ($manufacturer -like "*Dell*") 
{
    Write-Host "Dell detected"
    #Remove Dell bloat
    
    ##Dell

    $UninstallPrograms = @(
        "Dell Power Manager"
        "Dell SupportAssist OS Recovery"
        "Dell SupportAssist"
        "DellInc.PartnerPromo"
        "DellInc.DellPowerManager"
        "DellInc.DellDigitalDelivery"
            "DellInc.DellSupportAssistforPCs"
            "DellInc.PartnerPromo"
            "Dell Command | Power Manager"
            "Dell Digital Delivery Service"
        "Dell Digital Delivery"
            "Dell Peripheral Manager"
            "Dell Power Manager Service"
        "Dell SupportAssist Remediation"
        "SupportAssist Recovery Assistant"
            "Dell SupportAssist OS Recovery Plugin for Dell Update"
            "Dell SupportAssistAgent"
            "Dell Update - SupportAssist Update Plugin"
            "Dell Core Services"
            "Dell Pair"
            "Dell Display Manager 2.0"
            "Dell Display Manager 2.1"
            "Dell Display Manager 2.2"
            "Dell SupportAssist Remediation"
            "Dell Update - SupportAssist Update Plugin"
            "DellInc.PartnerPromo"
    )

    $WhitelistedApps = @(
        "WavesAudio.MaxxAudioProforDell2019"
        "Dell - Extension*"
        "Dell, Inc. - Firmware*"
    )

    $InstalledPackages = Get-AppxPackage -AllUsers | Where-Object {(($_.Name -in $UninstallPrograms) -or ($_.Name -like "*Dell*")) -and ($_.Name -NotMatch $WhitelistedApps)}

    $ProvisionedPackages = Get-AppxProvisionedPackage -Online | Where-Object {(($_.Name -in $UninstallPrograms) -or ($_.Name -like "*Dell*")) -and ($_.Name -NotMatch $WhitelistedApps)}

    $InstalledPrograms = $allstring | Where-Object {(($_.Name -in $UninstallPrograms) -or ($_.Name -like "*Dell*")) -and ($_.Name -NotMatch $WhitelistedApps)}
    # Remove provisioned packages first
    ForEach ($ProvPackage in $ProvisionedPackages) 
    {
        Write-Host -Object "Attempting to remove provisioned package: [$($ProvPackage.DisplayName)]..."

        Try 
        {
            $Null = Remove-AppxProvisionedPackage -PackageName $ProvPackage.PackageName -Online -ErrorAction Stop
            Write-Host -Object "Successfully removed provisioned package: [$($ProvPackage.DisplayName)]"
        }
        Catch {Write-Warning -Message "Failed to remove provisioned package: [$($ProvPackage.DisplayName)]"}
    }

    # Remove appx packages
    ForEach ($AppxPackage in $InstalledPackages) 
    {
        Write-Host -Object "Attempting to remove Appx package: [$($AppxPackage.Name)]..."

        Try 
        {
            $Null = Remove-AppxPackage -Package $AppxPackage.PackageFullName -AllUsers -ErrorAction Stop
            Write-Host -Object "Successfully removed Appx package: [$($AppxPackage.Name)]"
        }
        Catch {Write-Warning -Message "Failed to remove Appx package: [$($AppxPackage.Name)]"}
    }

    # Remove any bundled packages
    ForEach ($AppxPackage in $InstalledPackages) 
    {
        Write-Host -Object "Attempting to remove Appx package: [$($AppxPackage.Name)]..."

        Try 
        {
            $null = Get-AppxPackage -AllUsers -PackageTypeFilter Main, Bundle, Resource -Name $AppxPackage.Name | Remove-AppxPackage -AllUsers
            Write-Host -Object "Successfully removed Appx package: [$($AppxPackage.Name)]"
        }
        Catch {Write-Warning -Message "Failed to remove Appx package: [$($AppxPackage.Name)]"}
    }

    # Remove installed programs
    $InstalledPrograms | ForEach-Object {
        Write-Host -Object "Attempting to uninstall: [$($_.Name)]..."
        $uninstallcommand = $_.String

        Try 
        {
            if ($uninstallcommand -match "^msiexec*") 
            {
                #Remove msiexec as we need to split for the uninstall
                $uninstallcommand = $uninstallcommand -replace "msiexec.exe", ""
                $uninstallcommand = $uninstallcommand + " /quiet /norestart"
                $uninstallcommand = $uninstallcommand -replace "/I", "/X "   
                #Uninstall with string2 params
                Start-Process 'msiexec.exe' -ArgumentList $uninstallcommand -NoNewWindow -Wait
            }
            else 
            {
                #Exe installer, run straight path
                $string2 = $uninstallcommand
                start-process $string2
            }
            #$A = Start-Process -FilePath $uninstallcommand -Wait -passthru -NoNewWindow;$a.ExitCode        
            #$Null = $_ | Uninstall-Package -AllVersions -Force -ErrorAction Stop
            Write-Host -Object "Successfully uninstalled: [$($_.Name)]"
        }
        Catch {Write-Warning -Message "Failed to uninstall: [$($_.Name)]"}
    }

    ##Belt and braces, remove via CIM too
    foreach ($program in $UninstallPrograms) 
    {
        Get-CimInstance -Classname Win32_Product | Where-Object Name -Match $program | Invoke-CimMethod -MethodName UnInstall
    }

    ##Manual Removals

    ##Dell Optimizer
    $dellSA = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | Where-Object { $_.DisplayName -like "Dell*Optimizer*Core" } | Select-Object -Property UninstallString
 
    ForEach ($sa in $dellSA) 
    {
        If ($sa.UninstallString) 
        {
            cmd.exe /c $sa.UninstallString /SILENT /norestart
        }
    }

    ##Dell Dell SupportAssist OS Recovery Plugin for Dell Update
    $dellSA = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | Where-Object { $_.DisplayName -match "SupportAssist" } | Select-Object -Property UninstallString
 
    ForEach ($sa in $dellSA) 
    {
        If ($sa.UninstallString) 
        {
            cmd.exe /c $sa.UninstallString /quiet /norestart
        }
    }

    ##Dell Dell SupportAssist Remediation
    $uninstallcommand = "/X {C4543FDB-3BC0-4585-B1C5-258FB7C2EA71} /qn"
    Start-Process 'msiexec.exe' -ArgumentList $uninstallcommand -NoNewWindow -Wait
}
