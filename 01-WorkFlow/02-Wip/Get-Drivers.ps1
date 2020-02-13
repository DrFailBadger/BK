Get-WindowsDriver -All -Online #gets all drivers including system drivers

Get-WindowsDriver -online # online is if its local machine #removal of all only runs third party drivers

get-help Get-WindowsDriver
Get-Command -Noun "WINDOWSDR*"

Get-WindowsDriver -online 


Get-Service # services
Get-OdbcDsn # gets all odbc
Get-OdbcDsn -Name "Mypayroll" | Select-Object -ExpandProperty Attribute

Add-OdbcDsn -Name "Mypayroll" -DsnType "System" -DriverName "SQL Server" -SetPropertyValue "Database=Dave"
Set-OdbcDsn -Name "Mypayroll" -DsnType "System" -DriverName "SQL Server" -SetPropertyValue "Database=David"

get-help Set-OdbcDsn -online


Update-Help

$dave =Get-OdbcDsn -Name "Mypayroll"  
$dave.Attribute.Values

$hp = Get-WindowsDriver -Online | Where-Object {$_.ProviderName -EQ "hp"-and $_.ClassName -eq "Printer"} 
$bd = Get-ChildItem -path $hp.OriginalFileName -Directory
Export-WindowsDriver -online -Destination "C:\aPPvTEST" 

$bd.CreationTimeUtc
#recurse Windows driver folder for drivers modified or installed in last day.
$path2 = "C:\Windows\System32\DriverStore\FileRepository\"
$badger = Get-ChildItem $path2 | Where-Object {$_.LastAccessTime -ge $(get-date).adddays(-1)}
$badger.LastWriteTime
Get-Dat

Get-ChildItem $path2 | Select-Object -property *