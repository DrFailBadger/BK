Get-Process
Get-Process | Where-Object { $_.name -eq "notepad" } | get-mem    
Get-Process | Where-Object { $_.name -eq "notepad" } | Stop-Process

(Get-Process | Where-Object { $_.name -eq "notepad" }).Kill()

Get-Process | Sort-Object -Property Name | Where-Object { $_.Name -eq "Notepad" }

Get-Process -name notepad | Sort-Object -Property Id

Get-Process -name notepad | Sort-Object -Property Id | Stop-Service

ipconfig | Get-Member

$var1 = "world"
Write-Output $var1
$procs = Get-Process
$procs[0]
$procs[0] | Get-Member
$procs.GetType()
$procs.GetType().FullName
Get-Process | Format-List
Get-Process | Select-Object -Property name, @{name = 'procid'; expression = { $_.Id } }

Get-Process -name notepad | Stop-Process 
get-help Stop-Process -Full

Get-Process | Where-Object { $_.Handles -gt 1000 }
Get-Process | Where-Object handles  -gt 1000 | Sort-Object -Property Handles | Format-Table name, Handles -autosize

#Get-Process | Out-GridView -passthru | stop-process

get-process -name w* | clip 

Get-Process > C:\aPPvTEST\proce.txt
C:\aPPvTEST\proce.txt
Remove-Item C:\aPPvTEST\proce.txt
Get-Alias del
Get-Process | Out-File .\test.txt
Get-Content .\test.txt
Get-Alias cat
cat .\test.txt
Remove-Item .\test.txt

Get-Process | Export-Csv proc.Export-Csv
Get-Content .\proc.csv
$procs = Import-Csv .\proc.csv
$procs
$procs | Get-Member

Get-Process | Export-Clixml .\proc.xml
$procs = Import-Clixml .\proc.xml
$procs | Get-Member
Get-Process | Measure-Object WS -Sum -Maximum -Minimum -Average
Get-Process | Sort-Object -Property WS -Descending | Select-Object -First 5

Get-EventLog 
Get-WinEvent -LogName security  -MaxEvents 5 -Oldest
$applog = Get-WinEvent -LogName application
$applog  | Select-Object -First 10 | Where-Object -Property ProviderName -eq "MsiInstaller"
$applog  |  Where-Object {$_.ProviderName -eq "MsiInstaller" -and ($_.leveldisplayname -eq "Error" -or $_.leveldisplayname -eq "Warning")} | Sort-Object -Property TimeCreated |Format-Table -Property TimeCreated, leveldisplayname, Message
$applog  |  Where-Object -Property ProviderName -eq "MsiInstaller" 
$now = Get-Date
$now.AddDays(-1)
#Get Event log - Application - MSInstaller events - Errors or Warning for last 100 days - change date value for less days
$applog  |  Where-Object -Property ProviderName -eq "MsiInstaller" 
$applog  |  Where-Object {$_.ProviderName -eq "MsiInstaller" -and $_.TimeCreated -ge $now.AddDays(-100)-and ($_.leveldisplayname -eq "Error" -or $_.leveldisplayname -eq "Warning")}

Get-NetAdapter | Where-Object {$_.Name -like "Ethernet*"} | Enable-NetAdapter
$procs1 = Get-Process
$procs2 = Get-Process

notepad

Compare-Object $procs $procs2
Compare-Object -ReferenceObject $procs2 -DifferenceObject $procs1 -Property name
get-date
$path1 = "C:\Git\BK"

#Folders and files create in the last hour
Get-ChildItem $path1 -Recurse | Where-Object {$_.LastWriteTime -ge $(Get-Date).AddHours(-1) }


$path = Get-ChildItem HKCU:\Software
cd HKCU:dir

dir HKCU:\Software | Get-ItemProperty | gm

$path1 = Get-ChildItem HKCU:\Software
Compare-Object -ReferenceObject $path -DifferenceObject $path1 -Property name
