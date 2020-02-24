#### Learning Powershell, prompting for input, functions etc.






write-host hello -ForegroundColor Red -BackgroundColor Green


Read-Host  


$computername = Read-Host "Enter the computer name for the query"
"Dc"
$computername # come back as dc if  this was entered.

$VerbosePreference = 'continue'
Write-Verbose 'hi'

$VerbosePreference = 'silentlycontinue'
Write-Verbose 'ssdsd'

#####Creating Parameters to script


Get-WmiObject -class win32_logicalDisk -filter "DriveType=3" -ComputerName localhost |
Select-Object @{n='computername';e={$_.__Server}},
              @{n='Driver';e={$_.DeviceID}},
              @{n='Freespace(GB)';e={$_.Freespace / 1GB -as [int]}},
              @{n='Size(gb)';e={$_.Size / 1GB -as [int]}}


Get-WmiObject win32_logicalDisk -filter "DriveType=3" | fl *


########IF Statements

$procs =Get-Process

If( $procs.Count -gt 1000){
    Write-Host " You have alot of processes"
}
elseif ($procs.Count -lt 5) {
    Write-host "Very Few processes"
}
elseif ($procs[0].name -like 'a*') {
    Write-host "the first proc starts with a"
}
else {
    Write-Host "less than 100 processes"
}
#as many else if as you want

$disk = Get-WmiObject -class win32_logicalDisk|
ForEach-Object {
    $disk =$_
    if($disk.DriveType -eq 2){
        Write-Host "Drive is floppy"
    }elseif($disk.DriveType -eq 3){
        Write-Host "Drive is Fixed"
    }else{
    Write-Host " Dunno what drive is"
    }
}

if($disk.DriveType -eq 2){
    Write-Host "Drive is floppy"
}elseif($disk.DriveType -eq 3){
    Write-Host "Drive is Fixed"
}else{
Write-Host " Dunno what drive is"
}

$disk = Get-WmiObject -class win32_logicalDisk
switch ($disk.DriveType) {
    2 { Write-Host 'Floppy' }
    3 { Write-Host 'Fixed' }
    4 { Write-Host"optiical" }
    Default {Write-Host 'Don`t know which drive'}
}

$Name = Read-Host " enter Server Name"
switch -wildcard ($name) {
    "*DC*"    { Write-Host "is a domain controller"  }
    "*FS*"    { Write-Host "is a File Server"}
    "*NYC*"   { Write-Host "New york" }
    "*LON*"   {  Write-Host "London"  }
    Default {}
}

#Enter NYC-dc-01 - as this keeps going both are evaluated
#is a domain controller
#New yor

$Name = Read-Host " enter Server Name"
switch -regex ($name) {
    "DC"    { Write-Host "is a domain controller"  }
    "FS"    { Write-Host "is a File Server"}
    "^NYC"   { Write-Host "New york" }
    "^LON"   {  Write-Host "London"  } #lon at begining
    Default {}
}

#does the same but as an if statement not a switch
if ($Name -match "DC") {Write-Host "Domain Controller"}
if ($Name -match "FS") {Write-Host "Filer Server"}
if ($Name -match "^NYC") {Write-Host "New York"}
if ($Name -match "^LON") {Write-Host "London"}

####Scipting loops

##foreach

foreach ($item in $collection) {
    
}
notepad

Get-Process -name "Notepad" |
ForEach-Object {$_.Kill()}


$Procs = Get-Process -name "Notepad" 
foreach ($Proc in $PRocs) {
    $proc.Kill()
    
}

$Array1 = "one", "Two", "Three", "four", "Five"
$array2 = "fred", "barney", "wilma", "Betty", "pebbles"
$string = "one day we saw two eating four"

for([int]$x = 0 ; $x -lt $array1.count ; $x++){
    Echo "Loop $x"
    $string = $string -replace $array1[$x],$array2[$x]
}
$string

#######while

$Existing = 'Server1','Server3','Server4', 'Server7','Server2'

$Candidate = 0

do {
    $Candidate++
    $possiblename ="Server$candidate"
}while ($existing.contains($possiblename)) 
$possiblename
#########


$this = 5
$that = 5
while ($this -eq $that) {
    write-host 'hello'
    $that++
}


do {
    
} until ($this -eq $that)


###while loops
$i = 0

do {$I;$I +=11
} while ($i -lt 99)



$i = 0 

do {
    
} while (condition)





#####Do loopps --- always perform a condition once and check condition last
$problemservice = "Service name"
$computername = " Computer name"


##continually try to start a service while it's stopped
Do{
    start-service -name $problemservice -ComputerName $computername
} while ((Get-serice -name $problemservice).status -eq 'stoppped')

#while loops -- check condition first and run action last.....

while ((get-service -name $problemservice).status -eq 'stopped') {
    start-service -name $problemservice -ComputerName $computername    
}

#### do/untiol loop.....
##This attempts to start the service untilk its running. opposite of  above

do {
    start-service -name $problemservice -ComputerName $computername

} until ((get-service -name $problemservice -ComputerName $computername)status -eq 'Running')


######Functions
$computername = 'localhost'
Function Get-ComputerSystemInfo 
    param (
        [String[]]($computername)
          )
    Get-WmiObject -Class win32_computersystem -computername $computername | 
        Select-Object -prop   Name, Manufacturer, Model
}
Get-ComputerSystemInfo -computername Localhost

notepad computers.txt

filter Get-ComputerSystemInfo{
Get-WmiObject -Class win32_computersystem -computername $computername | 
Select-Object -prop   Name, Manufacturer, Model
}


function Verb-Noun {
    
    param (
        [String[]]$computername
    )
    
    begin {
        
    }
    
    process {
        if ($_ -ne $null) {
            $computername = $_
        }
        foreach ($Computer in $computername) {
            Get-WmiObject -Class win32_computersystem -computername $computername | 
            Select-Object -prop   Name, Manufacturer, Model
        }
        
    }
    
    end {
    }
}


####
function Get-LatestSecurityLog{
    param (
        
        [string]$computername
    )
    get-eventlog -logname security -newest 50 -computername $computername
}