function Get-LatestSecurityLog{
    param (
        
        [string]$computername
    )
    get-eventlog -logname security -newest 50 -computername $computername
}

function Get-OSInfo {
    param (
        [String]$computername
    )
    get-ciminstance -className Win32_Bios -computername $computername
}

$computername = env:computername
Get-OSInfo
Get-LatestSecurityLog -computername $computername

get-eventlog -logname security -newest 50 -computername localhost

$env:computername

get-wind