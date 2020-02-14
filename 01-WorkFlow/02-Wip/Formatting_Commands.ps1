#Format-Wide
Get-Process | Format-Wide
Get-Process | Format-Wide -Property Id -Autosize
Get-Process | Format-Wide -Property Id -Column 6

get-process | Format-List -Property Id.cpu,vm,ws,pm

get-process | Format-List -Property *

get-process | Format-List -Property
Get-Process -Name "Notepad" | Format-List * #shows all properties
get-process | Format-Table -Property name,id,vm,cpu,ws -AutoSize -Wrap

get-process | ft -Property *
get-service | sort-object Status | ft -GroupBy status -Property name,status,DisplayName
Get-Process | FT -Property name,id,@{n='vm(MB)';e={$_.vm / 1MB}},@{n='PM(MB)';e={$_.PM / 1MB};'formatstring'='N2';'align'='right';width=10} -AutoSize


##format --------------->>>>


get-service | Where-Object -FilterScript {$_.status -eq 'running' -and $_.name -like 's*'}


get-service | Where-Object  {$_.status -eq 'running' -and  $_.name -like 's*'} | fw

get-service | ?  {$_.status -eq 'running' -and  $_.name -like 's*'} | fw

Get-Process | Where-Object {$_.Responding -eq $true} | FW -Property
Get-Process | Where-Object {$_.Responding } | FW -Property name

Get-Process | Where-Object {-not $_.Responding } | FW -Property name

get-service | Where-Object {$PSItem.status -eq 'running'}


get-service | Where-Object status -eq running #single comparision not multiple
Get-service | Where-Object status -eq running | where name -like 's*'

get-service | where {$PSItem.Status -eq 'Running' -and $PSItem.Name -like 's*'}

get-service | where {$PSItem.Status -eq 'Running' -and $PSItem.Name -like 's*'}

Get-Service -Name s* | where Status -eq running

$x ='hello'
$x -is [string] # checks if value or variable is a string
$x -is [int] #checks if value is interga
$x -as [int]

56.776776 -as [int] #changes to what it would be
$x ='5555555'
$y = $x -as [int]
$y

$x = 'Powershell'
$x -like '*shell'

$x -contains "shell"
$x = 1,2,3,4,5,6,'one','two','three','four','five','six'

$x -contains 'one'
$x -contains '7'
$x -notcontains 7
7 -in $x
8 -notin  $x

$1 = 'powershell'
$1 -replace 'l','x'
$2 = $1 -replace 'l','x'
$xc = 1,2,3,4,5,6,7,8,9
$xc += "one","Two","three" # add to variable / array
$xc = $xc + 'four', 'five','six' # same as above

$list = $xc -join ','

$list[0]
1
$list[1]
,
 
$a = $list -split ","
$v = 10
10
$v += 10
20
$v *= 10
200
$s = 'helllo'
$s += ' there' #concats
Hello there

$v++ # adds one
202
$v-- # minus one
# bit in a byte
# 1 2 4 8 32 64 128
# 1 1 1 0 0   0  0
# 7


(5 -gt 1) -and (5 -lt 10)

>> # append
out-file -Append
2> # error to output

1..100 # creats range of items
1..49 | get-random

1..100 | ForEach-Object {}
"{0} {1:N2} {3} {2:N4}" -f "hello",4.567765,6.42313123,"there" #fomrat -f operator











