#==========================================================================
#
# APPNAME   :   <OutlookHolidayFile_1.0_x64_R01>
# AUTHOR    :   <David Hislop>
# DATE      :   <15/03/2020>
#
# COMMENT   :   <Script to Copy new Outlook.hol file to office folder
#                Machine based deployment, requires adding to task sequence>
#
#==========================================================================

# Get the script parameters if there are any
param
(
[ValidateSet('Install','Uninstall')]
[String]$InstallType = 'Install'
)


############################
# Functions                #
############################

function Write-Log
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Alias("LogContent")]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [Alias('LogPath')]
        [string]$Path=$DWPLogPath,
        
        [Parameter(Mandatory=$false,
                   Position=1)]
        [ValidateSet("Error","Warn","Info")]
        [string]$Level="Info",
        
        [Parameter(Mandatory=$false)]
        [switch]$NoClobber
    )
    Begin{}
    Process
    {
        
        # If the file already exists and NoClobber was specified, do not write to the log.
        if ((Test-Path $Path) -AND $NoClobber) {
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name."
            Return
            }

        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path.
        elseif (!(Test-Path $Path)) {
            Write-Verbose "Creating $Path."
            $NewLogFile = New-Item $Path -Force -ItemType File
            }

        $FormattedDate = Get-Date -Format "dd-MM-yyyy HH:mm:ss"

        switch ($Level) {
            'Error' {$LevelText = 'ERROR:'}
            'Warn' {$LevelText  = 'WARNING:'}
            'Info' {$LevelText  = 'INFO:'}
            }

        "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append
    }
    End{}
}

############################
# Standard Variables       #
############################
$ScriptRoot = $PSScriptRoot
$DWPVendor = "Microsoft"
$DWPAppName = "OutlookHolidayFiles"
$DWPVersion = "1.0.0"
$DWPArch = "X64"
$DWPRelease = "R01"
$DWPFullName = $DWPVendor + "_" + $DWPAppName + "_" + $DWPVersion + "_" + $DWPArch +"_" + $DWPRelease
$DWPLogFolder = $env:SystemDrive + "\DWPSource\MSILogs\"
$LogDate = Get-Date -Format "dd-MM-yyyy_HH-mm-ss"
$DWPLogPath = "$DWPLogFolder" + "$DWPFullName" + "_" + "$InstallType" + "_" + "$LogDate.log" 
$errorall = $null
############################
# Additonal Variables      #
############################
$OfficePath = $env:ProgramW6432 + "\Microsoft Office\root\Office16\1033\"
$FileName = "outlook"
$HolFile = $OfficePath + $filename + ".hol"
$NewHolFile = $ScriptRoot + "\Outlook.hol"

############################
# Install Section          #
############################
if ($installtype -eq "install") {

    Write-Log "Start of $($installtype)ation $DWPFullName" 

    if (Test-path -path $HolFile) {
        try {
            Move-Item -Path $HolFile -Destination ($OfficePath + $FileName + ".hol.bk") -Force -ErrorAction Stop -ErrorVariable x
            Write-Log "Moving of `"$HolFile`"  to `"$OfficePath`Outlook.hol.bk`" was Successful "
        }
        catch {
            $errorall += $x
            Write-Log "Moving of`"$HolFile`" to `"$OfficePath`Outlook.hol.bk`" was Unsuccessful" -Level Error
            Write-Log "PowerShell Error[$($_.exception.message)]"  -Level Error  
        }
    }

    if (!(Test-path -path $HolFile)) {
        try {
            Copy-Item -Path $NewHolFile -Destination $HolFile -Force -ErrorAction Stop -ErrorVariable x
            Write-Log "Copying of `"$NewHolFile`"  to `"$HolFile`" was Successful "
        }
        catch {
            $errorall += $x
            Write-Log "Copying of `"$HolFile`" to `"$HolFile`" was Unsuccessful" -Level Error
            Write-Log "PowerShell Error[$($_.exception.message)]"  -Level Error  
        }
    }

}

############################
# Uninstall Section        #
############################

if ($installtype -eq "Uninstall") {

    Write-Log "Start of $($installtype)ation $DWPFullName" 

    if (Test-path -path $HolFile) {
        try {
            Remove-Item -Path $HolFile -Force -ErrorAction Stop -ErrorVariable x
            Write-Log "Deleting of `"$HolFile`" was Successful "
        }
        catch {
            $errorall += $x
            Write-Log "Deleting of`"$HolFile`" was Unsuccessful" -Level Error
            Write-Log "PowerShell Erro[$($_.exception.message)]" -Level Error  
        }
    }

}

############################
# Exit Script              #
############################

if ($errorall.count -gt 0) {
    Write-Log "$($installtype)ation of $DWPFullName was Unsuccessful with $($errorall.count) Errors" Error
    #[System.Environment]::Exit(1603)
    $LASTEXITCODE = 1
}else{
    Write-Log "$($installtype)ation of $DWPFullName was Successful"
    #[System.Environment]::Exit(0)
    $LASTEXITCODE = 0
}
Write-Host $LASTEXITCODE
Write-Output $LASTEXITCODE
#notepad.exe $DWPLogPath

