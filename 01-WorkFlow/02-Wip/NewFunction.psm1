Function get-sysinfo {
    param (
        [String[]]$computername
    )
    foreach ($Comp in $computername) {
        $OS = Get-WmiObject -Class win32_operatingsystem -ComputerName $computername
        $CS = Get-WmiObject -Class win32_Computersystem -ComputerName $computername
        $Bios = Get-WmiObject -Class win32_BIOS -ComputerName $computername
    }


    $props= [ordered]@{
        'ComputerName' = $comp;
        'OSVersion' = $os.version;
        'SPVersion' = $os.servicepackmajorversion;
        'MFGR' = $cs.manufacturer;
        'Model' = $cs.Model
        'RAM' = $cs.totalphysicalmemory;
        'BIOSSerial' = $bios.serialnumber
    }
    $obj = New-Object -TypeName psobject -Property $props
    Write-Output $obj
}