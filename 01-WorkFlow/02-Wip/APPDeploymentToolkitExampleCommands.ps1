###to Create a new MST

        $transfromProperties = @{
            'ALLUSERS'='1';
            'AgreeToLicense'='ReallySuppress';
            'RebootYesNo'='No';
            'ROOTDRIVE'='C:';
        }

        New-MsiTransform C:\00_Project\7zip\7z1900-x64.msi -TransformProperties $transfromProperties

		## <Perform Installation tasks here>

        Execute-MSI -Path 7z1900-x64.msi -Transform 7z1900-x64.mst -Action Install -AddParameters 'ARPNOREMOVE=0'
           
		## <Perform Pre-Installation tasks here>
        $InstallValue = Get-RegistryKey -Key "HKLM:\software\DWPinstalls\" -Value "7Zip"

		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
        $7zipinstallname = "7Zip 1900"
		## <Perform Post-Installation tasks here>
        Set-PinnedApplication -Action PintoTaskbar -FilePath "C:\HypervCopy\2_7zip\Orca.exe - Shortcut.lnk"
        Set-PinnedApplication -Action PintoTaskbar -FilePath "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Orca.exe - Shortcut.lnk"
        Remove-Folder "$env:SystemDrive\hypervcopy\dave"
        Remove-File -Path "$env:SystemDrive\hypervcopy\badgers.txt"
        Remove-File -path "$env:SystemDrive\hypervcopy\badgers\" -
        Set-RegistryKey "HKEY_LOCAL_MACHINE\SOFTWARE\DWPInstalls" -Name $7zipinstallname -Value '1'
        New-Shortcut -Path "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\myshortcut.lnk" -TargetPath "$env:ProgramFiles\Internet Explorer\iexplore.exe" -WorkingDirectory "$env:ProgramFiles\Internet Explorer\" -Arguments "www.google.co.uk" -IconLocation "$Env:windir\system32\shell32.dll" -IconIndex "10"
        New-Folder -path "$env:SystemDrive\hypervcopy\dancing queen"
        Remove-MSIApplications -Name "Orca" -Exact
        Test-RegistryValue -Key "HKEY_LOCAL_MACHINE\SOFTWARE\DWPInstalls" #gives true/false value
        Update-Desktop
        Get-WindowTitle -WindowTitle "powershell"
        Get-UniversalDate
        Get-UserProfiles
        [string[]]$ProfilePaths = Get-UserProfiles | Select-Object -ExpandProperty 'ProfilePath' > c:\text.txt



        ## <Perform Pre-Uninstallation tasks here>

        Set-PinnedApplication -Action UnpinfromStartMenu -FilePath "$env:ProgramFiles\7-zip\7zFM.exe"
        Set-PinnedApplication -Action UnpinfromTaskbar -FilePath "$env:ProgramFiles\7-zip\7zFM.exe"
        Remove-RegistryKey -key "HKEY_LOCAL_MACHINE\SOFTWARE\DWPInstalls" -name $7zipinstallname
        