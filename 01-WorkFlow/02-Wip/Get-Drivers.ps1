Get-WindowsDriver -All -Online #gets all drivers including system drivers

Get-WindowsDriver -online # online is if its local machine #removal of all only runs third party drivers

get-help Get-WindowsDriver
Get-Command -Noun "WINDOWSDR*"

Get-WindowsDriver -online 


Get-Service # services

$hp = Get-WindowsDriver -Online | Where-Object {$_.ProviderName -EQ "hp"-and $_.ClassName -eq "Printer"} 
$bd = Get-ChildItem -path $hp.OriginalFileName -Directory
Export-WindowsDriver -online -Destination "C:\aPPvTEST" #export all windows drivers

$bd.CreationTimeUtc
#recurse Windows driver folder for drivers modified or installed in last day.
$path2 = "C:\Windows\System32\DriverStore\FileRepository\"
$badger = Get-ChildItem $path2 | Where-Object {$_.LastAccessTime -ge $(get-date).adddays(-1)}
$badger.LastWriteTime
get-

$dave = Get-WindowsDriver -Online | Where-Object {$_.ProviderName -EQ "hp"-and $_.ClassName -eq "Printer"}
gci $dave.OriginalFileName -Directory
Remove-WindowsDriver -Driver $dave.Driver

$drivers = Get-WindowsDriver -Online | Where-Object {$_.ProviderName -EQ "hp"-and $_.ClassName -eq "Printer"} 
{
    pnputil -f -d $driver.Driver
}

$path2 = "C:\Windows\System32\DriverStore\FileRepository\"
Get-ChildItem $path2 | Where-Object {$_.LastAccessTime -ge $(get-date).adddays(-1)} | Write-Output
{
    pnputil -f -d $driver.Driver
}

$drivers = 


$1 = 
Get-WindowsDriver -Online | ForEach-Object -Process {$_.Driver } | Where-Object $_.LastAccessTime -ge $(get-date).AddDays(-1)

foreach ($2 in $1) {

    Write-Output $1.Driver
}

$Driversall = Get-WindowsDriver -Online #| Where-Object {$_.providername -eq "Hewlett-Packard" -or $_.ProviderName -eq "HP"}

#Remove Drivers that have been installed in the last day, this checks the driver then the folder where the inf is created for the last modified date.
foreach ($Drive in $Driversall) {

    #$DriverName = $drive.Driver
    #$Locaiton = $drive.OriginalFileName
    $GCI = Get-ChildItem $drive.OriginalFileName -Directory 
    
    if ($gci.LastAccessTime -ge $(get-date).adddays(-1)) {

        pnputil -f -d $drive.Driver
        
    }
}
