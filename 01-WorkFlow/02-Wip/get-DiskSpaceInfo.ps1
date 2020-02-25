



function get-DiskSpaceInfo {
<#
.SYNOPSIS
Retrives basic disk ca[acioty infomration from one or more computers
.DESCRIPTION
See the Synopsius. this isnt complex.
.PARAMETER computername
one or more computer names
For Example:
get-diskspaceinfo - computername
.EXAMPLE 
Get0DiskSpaceInfo-ComputerName ONe, Two
This example retrieves disk spave info from computers one and two



#>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,
                   Position=1,
                   HelpMessage='Computername',
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias('Hostname')]
        # [ValidateCount(1,3)] # only allows between 1-3 computer names
        # [ValidatePattern]
        # [ValidateLength]
        [String[]]$computername,

        [Parameter(Position=2,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Floppy','Local','Optical')]
        [String]$DriveType =  'local', #$ExampleDriveTypePreference,

        [String]$ErrorLogFile = $ExampleErrorLogFile
        
    )
    Begin{
        Remove-Item $ErrorLogFile -ErrorAction SilentlyContinue
    }
    Process{
        foreach ($computer in $computername) {
            $params = @{ 'ComputerName' = $computer;
                         'Class' = 'Win32_LogicalDisk' }
            switch ($DriveType){
                'Local' {$params.add('Filter','DriveType=3')}
                'Floppy' {$params.add('Filter','DriveType=2')}
                'Optical' {$params.add('Filter','DriveType=5')}
                
            }
            Try {
                Get-WmiObject @Params -ErrorAction Stop -ErrorVariable myerr  |
                Select-Object @{n='Drive';e={$_.DeviceID}},
                              @{n='Size';e={"{0:N2}" -f ($_.Size /1GB)}},
                              @{n='FreeSpace';e={"{0:N2}" -f ($_.FreeSpace / 1Gb)}},
                              @{n='FreePercent';e={"{0:N2}" -f ($_.FreeSpace / $_.Size *100)}},
                              PSComputerName
            } Catch {
                $computer | Out-File $ErrorLogFile -Append
                Write-Verbose "Failed to connect to $computer; Error is $myerr"
            } 
        }              
    }
    End{}
}
