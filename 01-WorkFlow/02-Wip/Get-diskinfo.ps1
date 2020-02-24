

param (
        [string]$Computername = (Read-Host "Computer Name to Query"),
        [int]$Drivetype = 3
)


Get-WmiObject -class win32_logicalDisk -filter "DriveType=$Drivetype" -ComputerName $computername |
Select-Object @{n='computername';e={$_.__Server}},
              @{n='Driver';e={$_.DeviceID}},
              @{n='Freespace(GB)';e={$_.Freespace / 1GB -as [int]}},
              @{n='Size(gb)';e={$_.Size / 1GB -as [int]}}