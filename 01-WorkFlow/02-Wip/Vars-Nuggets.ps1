
get-service

$var2 = "hello" #anything in quotes will be executed
$var3 = "Sometext $Var2" #anything in quotes will be executed
$var3
$ServiceList = Get-Service
$ServiceList

$ServiceList[0]
$Servar = $ServiceList[0].name
$Servar1 = $($ServiceList[0].name) # subexpression anyting after the $(will be exectued as code)


$Servar
$Servar1

$a = "this is the new
string value
over here"
$a

$message = @" 

    var
    this 
    is a 
    string
"@ #here string, block of text etc make sure the "@ in on the last line. Large String
$message

$5 = 5 #interga
$5 = "5" #string
[int]$5d= 5 #intrgra 
[string]$sting = string # string
[float]$f = 5.4444 #float decimal
[bool]$badger = $true or $false # boolean true or false

