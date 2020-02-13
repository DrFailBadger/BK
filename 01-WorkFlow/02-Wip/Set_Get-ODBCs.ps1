


Get-OdbcDsn # gets all odbc
Get-OdbcDsn -Name "Mypayroll" | Select-Object -ExpandProperty Attribute

Add-OdbcDsn -Name "Mypayroll" -DsnType "System" -DriverName "SQL Server" -SetPropertyValue "Database=Dave"
Set-OdbcDsn -Name "Mypayroll" -DsnType "System" -DriverName "SQL Server" -SetPropertyValue "Database=David"

get-help Set-OdbcDsn -online
$dave =Get-OdbcDsn -Name "Mypayroll"  
$dave.Attribute.Values

Get-OdbcDriver

Set-OdbcDriver 