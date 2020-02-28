##testingvars
#$VMUserName = "Packaging User"
#$VMPassword = "P4ckag!ng"
#$vmxpath = "E:\Windows 7 x64\Windows 7 x64\Windows 7 x64.vmx"
#Get-RunningVMs
#red 255, 192, 192
#green 192, 255, 192


[void] [System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”)
$registryPath = "Registry::HKEY_CURRENT_USER\Software\Atos\UAMAssistant\Settings"
$information = [System.Drawing.SystemIcons]::Information
$Exclamation = [System.Drawing.SystemIcons]::Exclamation
$Sharepointinstance = (Get-ItemProperty -Path $registryPath -Name Sharepointinstance -ErrorAction SilentlyContinue).Sharepointinstance
IF($Sharepointinstance -eq "Acceptance"){
$strSharePointSiteURL = "https://uam-acc.ms.myatos.net/"
$strSharePointServer = "uam-acc.ms.myatos.net"
$trackerurl = "https://uam-acc.ms.myatos.net/SitePages/Home.aspx"
}else{
$strSharePointSiteURL = "https://uam.ms.myatos.net/"
$strSharePointServer = "uam.ms.myatos.net"
$trackerurl = "https://uam.ms.myatos.net/SitePages/Home.aspx"
}
$RunningAs = (Get-ItemProperty -Path $registryPath -Name RunningAs -ErrorAction SilentlyContinue).RunningAs
$script:OfflineMode = $false
Function PingTest {

  $checksp = Try {Invoke-WebRequest -uri $trackerurl -UseDefaultCredentials -disablekeepalive -method head -TimeoutSec 10} catch {}
  

	if((!(Test-Connection -Cn $strSharePointServer -BufferSize 16 -Count 1 -ea 0 -quiet)) -or ($checksp.StatusCode -ne "200")){
        $script:status = "Offline"
	}
	else {
	    $script:status = "Online"
	}
return $script:status
}

function ActiveApps {PAram($runningas)
 
    Clear-Variable Returnval -ErrorAction SilentlyContinue
    $1 = $Null
    $2 = $null
    $3 = $null
    $4 = $null
    $Returnval =@()

    [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
    [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($strSharePointSiteURL)
    [System.Net.CredentialCache]$credentials = New-Object -TypeName System.Net.CredentialCache
    $ctx.Credentials = $credentials.DefaultNetworkCredentials;
    $ctx.RequestTimeOut = 5000 * 60 * 10
    $web = $ctx.Web
    $list = $web.Lists.GetByTitle("Application Tracker")
    $camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
    $camlQuery.ViewXml = "<View>
<Query>
    <Where>
        <And>
            <Eq>
                <FieldRef Name='ApplicationTrackerAssignedTo' />
                <Value Type='Text'>$runningas</Value>
            </Eq>
            <Eq>
                <FieldRef Name='Active_x0020_request' />
                <Value Type='Text'>1</Value>
            </Eq>
        </And>
    </Where>
</Query>
<RowLimit>10000</RowLimit>
</View>"
    $spListItemCollection = $List.GetItems($camlQuery)
    $ctx.Load($spListItemCollection)
    $ctx.ExecuteQuery()

    foreach ($item in $spListItemCollection){
        $1 = $item['Title']
        $2 = $item['ApplicationName']
        $3 = $item['ApplicationStatus']     
        $4 = $item['PackageType']
        $5 = $item['Complexity']
        $6 = $item['No_x002e__x0020_of_x0020_test_x0']
        if($6 -eq $null){$6 = "1"}

        $Returnval += New-Object PsObject -Property @{ UAMID = "$1" ; ApplicationName = "$2" ; ApplicationStatus = "$3" ; PackageType = "$4" ; Complexity = "$5"; Numberoftestbuilds = "$6"}
        }
Return $Returnval
        } ##$runningas gives you their active apps 2##


$VMRunPath = "${env:ProgramFiles(x86)}\VMware\VMware Workstation\vmrun.exe"

$CurrentComboTextOptions = "$env:Temp\DiscoveryTool"
IF(!(Test-Path "$CurrentComboTextOptions\DefaultTextItems.txt")){
MD "$CurrentComboTextOptions" -Force 
"NULL NO TEXT WITH SCREEN CAPTURE
Double Click the installer to begin.
Click ""Next >""
Please Wait.
Click ""Install""
Check ""I Accept""
Click `"Finish`"" > "$CurrentComboTextOptions\DefaultTextItems.txt"
}


$icon = [System.Convert]::FromBase64String('
AAABAAMAEBAAAAEACABoBQAANgAAACAgAAABAAgAqAgAAJ4FAAAwMAAAAQAYAKgcAABGDgAAKAAA
ABAAAAAgAAAAAQAIAAAAAAAAAQAAAAAAAAAAAAAAAQAAAAAAAMmnbQC5jUMAzrB9AMKcWwDfzKoA
uo9GAPXv5ADdyKUA+/j0AMurcwDu5NIAtIQ0AMmnbgD07eIAt4o8ANrDmwDStoYA/fz6AOnbwwDD
nl8AzKt0APbx6ADHpGcA59e+AMWgYgC1hjgA6NnBALqPSADbxZ8A1LiKAP/+/gDq3ccA4tCyAL2T
TgDAmVYAwZlWAOvfygDGomYArnsnAOTUuADdx6MA1bqOAMqqcQC7kEcAyKZsAMCZVwDm1rsA1ryR
ALSFNQDJqG8A/Pv4ALKBMAD17uMA3cekALWHOADLqnIAw51dAP79+wDu49EAtok7ALeJOwDhz68A
7+XUAOfYvwD9+/kAz7GAALB+KQDaxJ0AtYc5ANO3iADo2sIA/v38AMajZgDUuYsA6dzFAP///wCy
gi8Ax6VpALiLPwC9lE8A8OfYAMGaVwC2iT0A3MahAPn28ADx6dsA697JALuRSAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABLS0tLS0tLS0tLS0tLS0tLS0tLS0tL
S0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tL
Sx9WETQ/VRIaBidKCAo9RggUVw1TQgkOSDs4TCoxFlc3UDAjCxMQDCAtD04ZPjoDTkclKQEoEE0g
TyECKwA2LCRLNTNEF1ImHS8iPA8bQQdDSwQFRREcSVQ5FTQeLhhRU0sRQEdLR0dLS0tLS0sRMh5L
S0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tL
S0tLS0tLAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAACgAAAAgAAAAQAAAAAEACAAAAAAAAAQAAAAAAAAAAAAAAAEAAAAAAADQsoEA
+fbwAPr38wC6jkQAoGUBANS5jQChZgQAyqlwAP79/ADy6dwAy6pzANi/lgC+lVAApGsKALKCMwDB
m1kAzrB8ALSGNgDo2sIAqHETAM+xfwDDnV8A+fXuAPr28QDgzKsAxqJlAKx4HwDTt4gArXkiAKBl
AgChZQIA1r2RAP38+gDXvpQA8ejaAKNqCADk07cAyqlxANi/lwDn2cAA9O7jAOjZwACncBEAzq96
APXu4wDcxaAAqHEUAKhyFwDfy6kA0raGAMaiZgC5jUMArHggAMajaQCfZAAAoGQAAPz7+ACiaAMA
r34pANa9kgDKqG8Ao2kGAL2TTACwfikAyqlyAKRqCQD07uQAwZpYAKdwEgDOr3sAtIU1ANvEngDr
38oA0bWEALeLPgD49O0A38uqAOzgzQCseCEArXkkAKBkAQD8+vYA/Pv5AOLRswD9+/kAomgEAPDm
1gC8kkoAyahwAKNpBwCjagoAsYAwAPPs3wD07eIAs4QzANrDnADn2L8Aza55AKdvEADOr3wAtIU2
ALeLPwD49O4A38qoAMWgYgDStYUAuIs/APz69wCufCUA7+XUALuRSACiZwIA1ryRALySSwCiaAUA
o2kIAMytdwC/mFQA2cKaAKZuDgDaw50Aza56AKdvEQDq3cYA0bSDALeKPQDeyaYA697JALeLQACr
dh0A0rWGAJ9jAAD7+fUA1buPAKFnAwC7kUkA/Pr4AOLQsgDv5dUAomcDALySTADMrHUA2cGYAPPr
3gCzgzIA0LOBAKl0GADq3ccA9/LqAKp1GwC3ij4A3smnAPv59gD8+fYArnskAO/k0wChZgEAoWcE
AO/l1gC8kUoAvJJNANjAlgCkbAoA8urcAP///wClbAoAsoIwAL+XUwDMrHYApW0NAKZuEACzgzMA
59e/APbx6AC2iTwAt4k8APr38QD7+PQA7uPRANS5iwChZgIA5NS3AP7+/QDy6t0A5dW6ANjAlwC+
llEApWwLALKCMQDZwZoAzKx3AOnbwwDq3MYA0bODAJ5iAAD69/IAn2IAANS5jACgZQAA4c6vAKFl
AACjagYA/v37APHp2wCkawkA//7+ALGALAClbQ8AqHIVAM+xfgAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAApKSkpKSk
pKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSk
pKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSk
pKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSk
pKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSk
pKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSk
pKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSk
pKSkpCShuZeYpDhMubkww7hjRbVrXXYKkSlUtrd+FEB8uISkjarGnxakwwzGhqeFd50TKg+mtKLJ
d3QirjcEVQRBMks4qMgLUlFUC8irIjM3EAkAHs+RSGE9BsVwPDBgLTqdIaQSqRpKr0o0NisXkMzA
pB83CrYwlSNVvK2kpKSZdcQ1pLCWNxFXZJx3vwKQu3+kGZwky04vR316m6SjeD5Zc46kpDHGO6Rw
N2UCAZC7f6QlxmcLg0UBRpJNSRpQUBoxAqSkTWKfZp+DG6RLkKWTpMdyQ1jGfl9yDLOdc4d2j82k
pKRRhy690C6yJ4xiOWq6GGgOvpzOz4FgFcIlOKTHcAKkpKQ7ixzGcSBTvKVvP14YFgWXbERuYKR5
wgx7VqBPKKSknoAdg1AmpMpCW1qshCCkCIiKbbGkpKMDiw1Ei5qPpKRcANHRgiikpGsmwSykpKSk
pKSkpKSkpJRnYQdpiYSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSk
pKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSk
pKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSk
pKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSk
pKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSk
pKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAKAAAADAAAABgAAAAAQAYAAAAAAAAGwAAAAAAAAAAAAAAAAAAAAAAAP//////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
//////////////////7+/fv59fv59f7+/f////////////////////z7+Pr38vz69v/+/v//////
//////////////////////38+vv49Pr38vz6+P///v///////////////+TUuMimbMimbMimbcil
a9K2h/v49P///////////+zgy8mnbcimbMimbcimbMimbOreyf7+/eTUuMSgYreKPbeKPcWgY+PS
tfz6+P/+/u7j0c+xf7uQSLWGN7mNQsqpcejZwf38+v/////+/vbx6OHOr8urdL2UTraJO7WHOLqP
RcimbePRtPv49P///////9G0g6BlAJ9jAJ9jAJ9jALGBLvjz7P///////////97JpaFmAp9kAJ9j
AJ9jAKBlANvFoNrDnaZuD59jAJ9iAJ5hAJ5hAKRrC8qpcdS5jKt2HKBkAJ5iAJ5hAJ5iAJ9jAKdv
EMyrdfn28PLq3LaIO6JoBZ9jAJ9iAJ5iAJ5iAJ9iAJ9jAKRqCcGbWvTu4/////Hp2trCm6lzF6Bk
ALB/K+TTt/38+v////////////bw5sysdaJnAqBlAKhxFdnCmunbxK98KKBkAKFmALOENMurdcCY
VadvEaFmAqJnBKBkAKNqCLWGN8GaWLqORKdvEKBlAKFlAselavfz68CYV59iAKNqB6x3HrSGNriL
QLB/KqNqCKBlAJ9kAMSeYPz69v////37+b+XU59jAKl0GO7j0f////79/f79/f79/f7+/smnbaBk
AKBkAL2TTv79/OTTuKNpB6BlAK15Ie/k1P////79/NS6jaFnAqBlAKhxFdnBmfn28P///urcx7OE
NKBlAKFlAKNpCN7Jp+3hzsurdeLQs/Lr3vr28fv59vfy6tnBmadvEKBlAKdxFOnbxP///////9vE
n6FnAqFmAbaIOsCZV8CYVcCYVcCYVr6WUqhyFqBlAKFmAdnBmf///9/Lq6JnA59jALyRSf37+f//
/////8qqcqBkAKFnAtO3if///////+LQs616I6BlAKBkAKBlAKBkALqPRvv49P//////////////
//////////fy6bGALKBkAKJoBdvFoP////////Lq3a15IqBlAKBkAJ5iAJ5hAJ5hAJ5hAJ9jAKFl
AKBlAKt3HvHo2v///9/Lq6JnAp9jAL2UT/38+/////r38bWHOZ9jAKx4IPLq3f///+veyqx4IKBk
AKJoBbqPRq9+KaBkAKp0GvDn2f////////////////7+/fPr38ysdqNqB6FlAKRqCeDNrf//////
//79/cSfYZ9kAKFnAruQR8ilasekaMaiZKp1GqBlAKBkAMKbW/79+////9/Lq6JnAp9jAL2UTv38
+/////bw5658Jp9jALWHOPr38/38+sGaWZ9jAKJnA8uqc/r38sWhY59jAKVtDerdyP////////bw
59/LqsWgYq16I6FmAqFmAJ9jALSFNvTu4////////////+DMrKNpBqBkAM+xfv////////Lp26x4
IKBlAKJoBN7Jp////////9/Lq6JnAp9jAL2UTv38+/////jz7LGBL59jALGALPjz7O/l1alzGJ9j
ALWGN/fy6v79/L+XU59jAKdwEu3i0fv49Na8kbGBLqJoBZ9kAKBkAKBkAKFmArSENOjZwv//////
//////////Xv5rB/K59jALSFNvj07f///9nCm6FmAqBkAK99KPTu4////////9/Lq6JnAp9jAL2U
Tv38+/////79+8CZV59jAKRsCuTTt+LQsaJoBKBkAM2tef///+/l1ax3Hp9jALKCMfPs4MekaaFn
A6BkAKBlAKBkAKZvD7iLP9a8kfXv5f///////////////////////8mnbp9jAKRrCuXVuv79/L2U
T6BkAKBkAMekav///v///////9/Lq6JnA59jAL2UT/38+////////97KqKNpB59jAMWhZNzGoqFm
AaBkANS5jPjz7MGZWKBlAKBkAM+xgOPRtaNqCaFlAKFlAK16Is2ueurdx/r38v//////////////
/////////////////+TUuaRrCp9jAMqpcu7j0alzF6BlAKRqCePRtf///+PRtcurdb2UT6FmAaBl
AK57JMuqc8urdc+xgPHo2cGbWp5iALqOReHOr6RqCKBlALaJO7eKPaFmAp9jALSFNfTu49G0hKBk
AKBkALSFNe7k0/////////z6+OfYv+HPsfPs4P////////////////////j07bSFNZ9jALKBL8uq
dKFmAaBkALODMvfy6////8qqdpxeAJ9jAKFmAKFmAKBlAJ5hAJ5gAKRsEe3i0PXu5L2TTax5IubX
vq99KKBkAKBkAJ9jAKBkALOEM+veyf///9G0hKBkAJ9kAMSgYv7+/v////////bw57B+KqFmAdnB
mf///////////////////////86wfaBkAKNpBqZuEKFmAKBkAMyteP///////9rDnbqORrKBMKFm
AaFlAKlyFrqPRrqPRr+XVPLq3P////jz7NW7kOjaw9CygKNpBqVtDa99KMysdvLr3v///////+TU
uaRrDKBlAKhxE9CygOnbxO/k0+TTt6t2HZ9jANW6jf////////////////38+9e+lbuQSKFnAqFm
AKFlAKFlAKVtDujZwf////////79/P79/d3IpqJnAp9jAL2TTfv49f37+f37+f/+/v/////////+
/v79+/j07ebXvOrcxvXv5v7+/f////////////z698imbKFmA59jAKBlAaVsC6hxE6ZuDaFmAZ9j
ANG0hv////////////////z6+LmORaBlAKNpBaNpBaNpBaFnAriLQPr28f///////////////+DM
rKNqB6FmAb6WUv38+///////////////////////////////////////////////////////////
//////r28dW7j7KCMKVtDKFmAaBlAKJnAqVtDKt2HdnAmP////////////////7+/erdyOLQtOPR
tePRtePRteLQs+3i0f////////////////////Xv5uPRtOLQsuvey/7+/v//////////////////
/////////////////////////////////////////////////////vfy6ujaw97Jp9vFoeDMrOnb
xPPr3/z69///////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////
/////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==')

Function Invoke-VMProgramExe {
<#
.DESCRIPTION
Invokes Psexec on the target VM
#>


    [CmdletBinding()]
    Param(       
                      
        [Parameter(Mandatory=$False,
                   ValueFromPipeline=$True,
                   ValueFromPipelineByPropertyName=$True,
                   HelpMessage = "Location of Running VM VMX file")]
        [ValidateScript({(Test-Path $_) -and ((Get-Item $_).Extension -eq ".vmx")})] 
        [Alias('Path')]
        [String]$VMXpath = $DefaultVMXPath,

        [Parameter(Mandatory=$False,
                   HelpMessage = "User Name for Target VMX")]
        [String]$VMUserName = $Script:DefaultVMUserName,

        [Parameter(Mandatory=$False,
                   HelpMessage = "Password for Target VMX")]
        [String]$VMPassword = $Script:DefaultVMPassword,

        [Parameter(Mandatory=$False,
                   HelpMessage = "See the command on the VM (Useful for error handling)")]
        [ValidateSet('True','False')]
        [String]$Visible = "True",

        [Parameter(Mandatory=$False,
                   HelpMessage = "Path and EXE you may also add Parameters in this string.")]
        [String]$EXE

    
    )

Begin{}
Process{

    If($VMUserName -eq ""){
        Write-Error '-VMUserName or $Script:DefaultVMUserName must be Set' -ErrorAction Stop
    }
    If($VMPassword -eq ""){
    Write-Error '-VMPassword or $Script:DefaultVMPassword must be Set' -ErrorAction Stop
    }
    If($VMXpath -eq ""){
       Write-Error '-VMXpath or $DefaultVMXpath must be Set' -ErrorAction Stop
    }


        & "$VMRunPath" -gu "$VMUserName" -gp "$VMPassword" runProgramInGuest "$VMXpath" -interactive -activewindow -nowait $EXE


}
END{}
}

Function Invoke-VMPSCommand {
<#
.DESCRIPTION
Run a powershell command on a VM.
Get-RunningVMs| Invoke-VMCommand -Command 'notepad.exe'
Powershell commands will return variables set as the -PSReturnVariable '$Variableyouwantreturning'
multiple variables can be returned -PSReturnVariable '$Variable1,$variable2'

Powershell return requires Z: drive to be mapped in the virtual machine
Require Powershell 3.0 on the target machine for powershell switch.
.EXAMPLE
Running a Command
$Script:DefaultVMUserName = "Packaging User"
$Script:DefaultVMPassword = "P4ckag!ng"
$final = Get-RunningVMs |Invoke-VMCommand -Command '$return = Get-Process' -PSReturnVariable '$return' -Visible True
$final will be the $return variable from the command ran on the machine
.EXAMPLE
Returning multiple Variables from a script block
$Script:DefaultVMUserName = "Packaging User"
$Script:DefaultVMPassword = "P4ckag!ng"
$final = Get-RunningVMs |Invoke-VMCommand -Command '$return = Get-Process
$return2 = Get-Service' -PSReturnVariable '$return,$return2' -Visible True
$final[0] will be the $return variable from the command ran on the machine
$final[1] will be the $return2 variable from the command ran on the machine
#>


    [CmdletBinding()]
    Param(       
                      
        [Parameter(Mandatory=$False,
                   ValueFromPipeline=$True,
                   ValueFromPipelineByPropertyName=$True,
                   HelpMessage = "Location of Running VM VMX file")]
        [ValidateScript({(Test-Path $_) -and ((Get-Item $_).Extension -eq ".vmx")})] 
        [Alias('Path')]
        [String]$VMXpath = $DefaultVMXPath,

        [Parameter(Mandatory=$False,
                   HelpMessage = "User Name for Target VMX")]
        [String]$VMUserName = $Script:DefaultVMUserName,

        [Parameter(Mandatory=$False,
                   HelpMessage = "Password for Target VMX")]
        [String]$VMPassword = $Script:DefaultVMPassword,

        [Parameter(Mandatory=$False,
                   HelpMessage = "Powershell command")]
        [String]$Command,

        [Parameter(Mandatory=$False,
                   HelpMessage = "Powershell command")]
        [String]$PSReturnVariable = '$Return',

        [Parameter(Mandatory=$False,
                   HelpMessage = "See the command on the VM (Useful for error handling)")]
        [ValidateSet('True','False')]
        [String]$Visible = "True"
    
    )

Begin{}
Process{

    If($VMUserName -eq ""){
        Write-Error '-VMUserName or $Script:DefaultVMUserName must be Set' -ErrorAction Stop
    }
    If($VMPassword -eq ""){
    Write-Error '-VMPassword or $Script:DefaultVMPassword must be Set' -ErrorAction Stop
    }
    If($VMXpath -eq ""){
       Write-Error '-VMXpath or $DefaultVMXpath must be Set' -ErrorAction Stop
    }

    $HostPath = $ENV:Temp.Replace(":","")
    MD "$ENV:Temp\VMRun" -ErrorAction SilentlyContinue
    #$PSReturnVariable = '$return'
    #$Command = '$return = Get-Process'
  
  IF($Command -ne $null){
        $Command = "$PSReturnVariable = ''
        $Command
        IF($PSReturnVariable -ne ''){$PSReturnVariable|Export-Clixml 'Z:\$HostPath\VMRun\Return.XML'}
        start-sleep 5
        "
        $version = 1
        $written = $false
        While($written -ne $true){
        Try{
        $Command | Out-File -FilePath "$ENV:Temp\VMRun\Input-$version.PS1" -Encoding ascii -Force
        $written = $true
        }Catch{$written = $false
        $Version++
        }
      
        }
        }

    IF($Visible -eq "True"){
    &  "$VMRunPath" -gu "$VMUserName" -gp "$VMPassword" runProgramInGuest "$VMXpath" -interactive -activewindow "$env:windir\System32\WindowsPowerShell\v1.0\PowerShell.exe" -executionpolicy bypass -file "Z:\$HostPath\VMRun\Input-$version.PS1"
    }
    else{& "$VMRunPath" -gu "$VMUserName" -gp "$VMPassword" runProgramInGuest "$VMXpath" -interactive "$env:windir\System32\WindowsPowerShell\v1.0\PowerShell.exe" -executionpolicy bypass -file "Z:\$HostPath\VMRun\Input-$version.PS1"}

    $final = Import-Clixml "$ENV:Temp\VMRun\Return.XML" -ErrorAction SilentlyContinue
    Del "$ENV:Temp\VMRun\Return.xml" -ErrorAction SilentlyContinue
    Del "$ENV:Temp\VMRun\Input.PS1" -ErrorAction SilentlyContinue
    Return $final

}
END{}
}

Function Get-RunningVMs {

$Paths = (& "$VMRunPath" list)|Select-Object -Skip 1
$return =@()
$paths|foreach{
$return += New-Object Psobject -Property @{
Path = $_
}
}
Return $return
}

Function Get-Difference {
    [CmdletBinding()]
    Param(       
                       
        [Parameter(Mandatory=$True,
                   HelpMessage = "-snapshots CreatesSnapshot, DeleteSnapshot,RevertSnapshot")]
        [ValidateSet('First','Second','Compare')]
        [String]$Capture,

        [Parameter(Mandatory=$False,
                   ValueFromPipeline=$True,
                   ValueFromPipelineByPropertyName=$True,
                   HelpMessage = "Location of VMX file")]
        #[ValidateScript({(Test-Path $_) -and ((Get-Item $_).Extension -eq ".vmx")})]
        [Alias('Path')]
        [String]$VMXpath = $DefaultVMXPath,

        [Parameter(Mandatory=$False,
                   HelpMessage = "User Name for Target VMX")]
        [String]$VMUserName = $Script:DefaultVMUserName,

        [Parameter(Mandatory=$False,
                   HelpMessage = "Password for Target VMX")]
        [String]$VMPassword = $Script:DefaultVMPassword


        )

BEGIN{}
PROCESS{
If($VMXpath -eq "" -and $Capture -ne "Compare"){
Write-Error '-VMXpath or $DefaultVMXpath must be Set' -ErrorAction Stop
}


$command = '


Write-Progress -Activity "Capturing System State" -percentComplete 5 -Status "Getting Processes"
$Processes = Get-Process |select-object -Property "Name"
start-sleep 0.4
Write-Progress -Activity "Capturing System State" -percentComplete 15 -Status "Getting Services"
$Services = Get-service |select-object -Property "Name"
start-sleep 0.4


$allEnvironmentals =@()
$allEnvironmentals += Get-Item -Path "Registry::HKEY_CURRENT_USER\Environment" |Select-Object -ExpandProperty property |ForEach-Object {
New-Object psobject -Property @{property=$_;
Value = (Get-ItemProperty -Path "Registry::HKEY_CURRENT_USER\Environment" -Name $_).$_}} 

$allEnvironmentals += Get-Item -Path "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" |Select-Object -ExpandProperty property |ForEach-Object {
New-Object psobject -Property @{property=$_;
Value = (Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" -Name $_).$_}} 

$EnvironmentalVars = Get-ChildItem Env: | sort { $_.value.length } -Descending

Write-Progress -Activity "Capturing System State" -percentComplete 20 -Status "Getting FireWall Rules"
$fw = New-object -ComObject HNetCfg.FwPolicy2
$FireWallrules =@()
$Firewallrules += $fw.Rules|Select-Object -Property Name  

Write-Progress -Activity "Capturing System State" -percentComplete 25 -Status "Getting Files"
$Files =@()
$files += Dir "$Env:ProgramFiles\*\*\*\*" -erroraction Silentlycontinue  |select-object -Property "FullName"
$files += Dir "$Env:ProgramFiles\*\*\*"  |select-object -Property "FullName"
$files += Dir "$Env:ProgramFiles\*\*"  |select-object -Property "FullName"
$files += Dir "$Env:ProgramFiles\*" |select-object -Property "FullName"
Try{
IF(Test-Path "${ENv:ProgramFiles(x86)}"){
$files += Dir "${ENv:ProgramFiles(x86)}\*\*\*\*" |select-object -Property "FullName"
$files += Dir "${ENv:ProgramFiles(x86)}\*\*\*" |select-object -Property "FullName"
$files += Dir "${ENv:ProgramFiles(x86)}\*\*" |select-object -Property "FullName"
$files += Dir "${ENv:ProgramFiles(x86)}\*" |select-object -Property "FullName"
}}Catch{}

$files += Dir "$ENV:HOMEDRIVE\*" |select-object -Property "FullName"
$files += Dir "$ENV:windir\*" |select-object -Property "FullName"
$files += Dir "$ENV:windir\System32\*\*" |select-object -Property "FullName"
$files += Dir "$ENV:windir\System32\*" |select-object -Property "FullName"
$files += Dir "$ENV:APPDATA\*\*\*" |select-object -Property "FullName"
$files += Dir "$ENV:APPDATA\*\*" |select-object -Property "FullName"
$files += Dir "$ENV:APPDATA\*" |select-object -Property "FullName"
$files += Dir "$ENV:LOCALAPPDATA\*\*\*" |select-object -Property "FullName"
$files += Dir "$ENV:LOCALAPPDATA\*\*" |select-object -Property "FullName"
$files += Dir "$ENV:LOCALAPPDATA\*" |select-object -Property "FullName"
$files += Dir "$ENV:APPDATA\..\LocalLow\*\*\*" |select-object -Property "FullName"
$files += Dir "$ENV:APPDATA\..\LocalLow\*\*" |select-object -Property "FullName"
$files += Dir "$ENV:APPDATA\..\LocalLow\*" |select-object -Property "FullName"
$files += Dir "$ENV:ProgramData\*\*\*" |select-object -Property "FullName"
$files += Dir "$ENV:ProgramData\*\*" |select-object -Property "FullName"
$files += Dir "$ENV:ProgramData\*" |select-object -Property "FullName"
if(test-path "C:\Packages"){
$files += Dir "C:\Packages\*" |select-object -Property "FullName"
$files += Dir "C:\Packages\*\*" |select-object -Property "FullName"
}

Write-Progress -Activity "Capturing System State" -percentComplete 50 -Status "Getting Registry"
$Registry =@()
$Registry += Dir "Registry::HKEY_CURRENT_USER\SoftWare\*\*\*" -ErrorAction SilentlyContinue |select-object -Property "Name"
$Registry += Dir "Registry::HKEY_LOCAL_MACHINE\SoftWare\*\*\*" -ErrorAction SilentlyContinue |select-object -Property "Name"

$Registry += Dir "Registry::HKEY_CURRENT_USER\" -ErrorAction SilentlyContinue |select-object -Property "Name"
$Registry += Dir "Registry::HKEY_CURRENT_USER\*\*" -ErrorAction SilentlyContinue |select-object -Property "Name"
$Registry += Dir "Registry::HKEY_CURRENT_USER\*\*\*" -ErrorAction SilentlyContinue |select-object -Property "Name"
$Registry += Dir "Registry::HKEY_LOCAL_MACHINE\" -ErrorAction SilentlyContinue |select-object -Property "Name"
$Registry += Dir "Registry::HKEY_LOCAL_MACHINE\*\*" -ErrorAction SilentlyContinue |select-object -Property "Name"
$Registry += Dir "Registry::HKEY_LOCAL_MACHINE\*\*\*" -ErrorAction SilentlyContinue |select-object -Property "Name"
$Registry += Dir "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall\" -ErrorAction SilentlyContinue |select-object -Property "Name"
$Registry += Dir "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall\*\*" -ErrorAction SilentlyContinue |select-object -Property "Name"

Write-Progress -Activity "Capturing System State" -percentComplete 55 -Status "Getting ODBCs"
$ODBCs =@()
$ODBCs += Get-Item -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers" -ErrorAction SilentlyContinue
$ODBCs += Get-Item -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\ODBC\ODBCINST.INI\ODBC Drivers" -ErrorAction SilentlyContinue

$ODBCs += Get-Item -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI\*" -ErrorAction SilentlyContinue
$ODBCs += Get-Item -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\ODBC\ODBC.INI\*" -ErrorAction SilentlyContinue


Write-Progress -Activity "Capturing System State" -percentComplete 60 -Status "Getting ShortCuts"
$Shortcuts = @()
$Shortcuts += Dir "$ENV:APPDATA\Microsoft\Windows\Start Menu" -Recurse |select-object -Property "FullName"
$Shortcuts += Dir "$ENV:ProgramData\Microsoft\Windows\Start Menu" -Recurse |select-object -Property "FullName"
$shortcuts += Dir "$env:USERPROFILE\Desktop" -Filter "*.lnk" |select-object -Property "FullName"
$shortcuts += Dir "$env:PUBLIC\desktop" -Filter "*.lnk" |select-object -Property "FullName"

if(Test-path "$ENV:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"){
$shortcuts += Dir "$ENV:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" -Filter "*.lnk" |select-object -Property "FullName"
}
if(Test-path "$ENV:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu"){
$shortcuts += Dir "$ENV:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu" -Filter "*.lnk" |select-object -Property "FullName"
}
####ShortCuts####

Function Get-ShortcutTarget {
    [CmdletBinding()] 
    param ( 
    [Parameter(Mandatory=$True, 
        ValueFromPipeline=$True, 
        ValueFromPipelineByPropertyName=$True)] 
        [ValidateScript({(Test-Path $_) -and ((Get-Item $_).Extension -eq ".lnk")})]
    $lnk
    )
$WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
$ShortcutTarget = $WindowsInstaller.GetType().InvokeMember("ShortcutTarget","GetProperty",$null,$WindowsInstaller,$lnk)
$StringData1 = $ShortcutTarget.GetType().InvokeMember("StringData","GetProperty",$null,$ShortcutTarget,1)
$StringData3 = $ShortcutTarget.GetType().InvokeMember("StringData","GetProperty",$null,$ShortcutTarget,3)
$WindowsInstaller.GetType().InvokeMember("ComponentPath","GetProperty",$null,$WindowsInstaller,@($StringData1,$StringData3))
$WindowsInstaller = $null
}

Function Get-ShortcutInfo {
    [CmdletBinding()] 
    param ( 
    [Parameter(Mandatory=$True, 
        ValueFromPipeline=$True, 
        ValueFromPipelineByPropertyName=$True)] 
        [ValidateScript({(Test-Path $_)})]
    $Path
    )

Process {
$shortcuts = Dir "$Path" -Filter "*.LNK" -Recurse
$sh = New-Object -COM WScript.Shell
$Discoshortcuts = @()
if($shortcuts -ne $null){
foreach($item in $shortcuts){
$argument = $sh.CreateShortcut("$($item.FullName)").Arguments
if($argument -eq ""){$argument = "N/A"}
$TargetPath = $sh.CreateShortcut("$($item.FullName)").TargetPath
if($TargetPath -match "$env:HOMEDRIVE\\Windows\\Installer"){
$targetpath = Get-ShortcutTarget -lnk $item.FullName
}elseif($targetpath -eq ""){
$targetpath = "N/A"
}
Try{$StartIn = Split-Path $TargetPath}
Catch{$StartIn = "N/A"}
$Discoshortcuts += New-Object PSObject -Property @{
ShortcutName = $item.BaseName;
Target = $TargetPath
Parameters = $argument
StartsIn = $StartIn
IconFileName = $sh.CreateShortcut("$($item.FullName)").IconLocation
LocationInStartMenu = $item.DirectoryName
}
}
}
}
END{
$sh = $null
return $Discoshortcuts
}
}

$allShortCuts =@()
$allShortCuts += Get-ShortcutInfo -Path "$ENV:APPDATA\Microsoft\Windows\Start Menu"
$allShortCuts += Get-ShortcutInfo -Path "$ENV:ProgramData\Microsoft\Windows\Start Menu"
$allShortCuts += Get-ShortcutInfo -Path "$env:USERPROFILE\Desktop"
$allShortCuts += Get-ShortcutInfo -Path "$env:PUBLIC\desktop"

if(Test-path "$ENV:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"){
$allShortCuts += Get-ShortcutInfo -Path "$ENV:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
}
if(Test-path "$ENV:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu"){
$allShortCuts += Get-ShortcutInfo -Path "$ENV:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu"
}
####ShortCuts####



Write-Progress -Activity "Capturing System State" -percentComplete 70 -Status "Getting Drivers"
$Drivers =@()
$Drivers += Dir "$env:windir\System32\DriverStore\FileRepository" |select-object -Property "FullName"
Write-Progress -Activity "Capturing System State" -percentComplete 80 -Status "Getting FTAs"
$FTAs1 =@()
$FTAs1 += cmd /C assoc
$FTAs =@()
foreach($item in $FTAs1){
$FTAs += New-Object PsObject -Property @{Name = $Item}
}

foreach($Item in $FTAs1){
$return = $null
$($return = cmd /C ftype ($item -split "=")[1]) 2>&1 | out-null
IF($return.length -gt 5){
$FTAs += New-Object PsObject -Property @{Name = $return}
}
}

$fw = $null

IF($allEnvironmentals.count -eq 0){$allEnvironmentals = "No Entry"}
IF($FireWallrules.count -eq 0){$FireWallrules = "No Entry"}
IF($Files.count -eq 0){$Files = "No Entry"}
IF($Registry.count -eq 0){$Registry = "No Entry"}
IF($ODBCs.count -eq 0){$ODBCs = "No Entry"}
IF($Shortcuts.count -eq 0){$Shortcuts = "No Entry"}
IF($Discoshortcuts.count -eq 0){$Discoshortcuts = "No Entry"}
IF($allShortCuts.count -eq 0){$allShortCuts = "No Entry"}
IF($Drivers.count -eq 0){$Drivers = "No Entry"}
IF($FTAs1.count -eq 0){$FTAs1 = "No Entry"}
IF($FTAs.count -eq 0){$FTAs = "No Entry"}
IF($Processes.count -eq 0){$Processes = "No Entry"}
IF($Services.count -eq 0){$Services = "No Entry"}



Write-Progress -Activity "Done" -percentComplete 100 -Status "Saving File Back From VM."

'
#$VMXpath = "E:\Windows 10 x64 Academy\Windows 10.vmx"
IF($Capture -eq "First"){
$Script:FirstCapture = Invoke-VMPSCommand -VMXpath $VMXpath  -Command $command -PSReturnVariable '$processes,$Services,$files,$Registry,$Shortcuts,$Drivers,$FTAs,$Firewallrules,$ODBCs,$allEnvironmentals'
return "First Capture Complete"
}
elseIf($Capture -eq "Second"){
$Script:SecondCapture = Invoke-VMPSCommand  -VMXpath $VMXpath -Command $command -PSReturnVariable '$processes,$Services,$files,$Registry,$Shortcuts,$Drivers,$FTAs,$Firewallrules,$ODBCs,$allEnvironmentals,$allShortCuts,$EnvironmentalVars'
return "Second Capture Complete"
}
else{
IF($Script:FirstCapture -ne $null -and $Script:SecondCapture -ne $null){

$Script:Compare =@()

#"Processes"
foreach($item in $(Compare-Object -ReferenceObject $Script:FirstCapture[0].Name -DifferenceObject $Script:SecondCapture[0].Name -PassThru) ){
$Script:Compare += New-Object PsObject -Property @{
Name = "Process"
Difference = "$item"
}
}
#"Services"
foreach($item in $(Compare-Object -ReferenceObject $Script:FirstCapture[1].Name -DifferenceObject $Script:SecondCapture[1].Name -PassThru) ){
$Script:Compare += New-Object PsObject -Property @{
Name = "Services"
Difference = "$item"
}
}
#"Files"
foreach($item in $(Compare-Object -ReferenceObject $Script:FirstCapture[2].FullName -DifferenceObject $Script:SecondCapture[2].FullName -PassThru) ){
$Script:Compare += New-Object PsObject -Property @{
Name = "Files"
Difference = "$item"
}
}

#"Registry"
foreach($item in $(Compare-Object -ReferenceObject $Script:FirstCapture[3].name -DifferenceObject $Script:SecondCapture[3].name -PassThru) ){
$Script:Compare += New-Object PsObject -Property @{
Name = "Registry"
Difference = "$item"
}
}

#"Shortcuts"
foreach($item in $(Compare-Object -ReferenceObject $Script:FirstCapture[4].FullName -DifferenceObject $Script:SecondCapture[4].FullName -PassThru) ){
$Script:Compare += New-Object PsObject -Property @{
Name = "Shortcuts"
Difference = "$item"
}
}


#"Drivers"
foreach($item in $(Compare-Object -ReferenceObject $Script:FirstCapture[5].FullName -DifferenceObject $Script:SecondCapture[5].FullName -PassThru) ){
$Script:Compare += New-Object PsObject -Property @{
Name = "Drivers"
Difference = "$item"
}
}

#"FTAs"
foreach($item in $(Compare-Object -ReferenceObject $Script:FirstCapture[6].name -DifferenceObject $Script:SecondCapture[6].name -PassThru) ){
$Script:Compare += New-Object PsObject -Property @{
Name = "FTAs"
Difference = "$item"
}
}

#"FireWallRules"
foreach($item in $(Compare-Object -ReferenceObject $Script:FirstCapture[7].Name -DifferenceObject $Script:SecondCapture[7].Name -PassThru) ){
$Script:Compare += New-Object PsObject -Property @{
Name = "FireWallRules"
Difference = "$item"
}
}

#"ODBCs"
foreach($item in $(Compare-Object -ReferenceObject $Script:FirstCapture[8].Property -DifferenceObject $Script:SecondCapture[8].Property -PassThru) ){
$Script:Compare += New-Object PsObject -Property @{
Name = "ODBCs"
Difference = "$item"
}
}

#"Environmentals"
foreach($item in $(Compare-Object -ReferenceObject $Script:FirstCapture[9].Property -DifferenceObject $Script:SecondCapture[9].Property -PassThru) ){
$Script:Compare += New-Object PsObject -Property @{
Name = "Environmentals"
Difference = "$item"
}
}




}
Return $Script:Compare
}


}
END{}




}
#Get-Difference -Capture Compare
Function Validate {

$captureFirstToolStripMenuItem.Enabled = $True 
$captureSecondToolStripMenuItem.Enabled = $True 
$compareToolStripMenuItem.Enabled = $True 


$exportfirstToolStripMenuItem.Enabled = $True 
$exportsecondToolStripMenuItem.Enabled = $True 
$exportcompareToolStripMenuItem1.Enabled = $True 
$exportexclusionListToolStripMenuItem.Enabled = $True

$captureFirstToolStripMenuItem.backcolor = $menustrip2.backcolor
$captureSecondToolStripMenuItem.backcolor = $menustrip2.backcolor
$compareToolStripMenuItem.backcolor = $menustrip2.backcolor

if($Script:FirstCapture -ne $null){
$captureFirstToolStripMenuItem.backcolor = "192, 255, 192"
}
if($script:SecondCapture -ne $null){
$captureSecondToolStripMenuItem.backcolor = "192, 255, 192"
}
if($Script:Compare -ne $null ){
$compareToolStripMenuItem.backcolor = "192, 255, 192"
}
IF($Script:FirstCapture -eq $null){
 
$captureSecondToolStripMenuItem.Enabled = $False 
$compareToolStripMenuItem.Enabled = $False

$exportfirstToolStripMenuItem.Enabled = $False 
$exportsecondToolStripMenuItem.Enabled = $False 
$exportcompareToolStripMenuItem1.Enabled = $False 


}
elseif($script:SecondCapture -eq $null){

$compareToolStripMenuItem.Enabled = $False 
 
$exportsecondToolStripMenuItem.Enabled = $False 
$exportcompareToolStripMenuItem1.Enabled = $False 
$exportexclusionListToolStripMenuItem.Enabled = $False 

}


$rowcount = $Script:datagridview1.RowCount
$i = 0
While($i -lt $rowcount){

$dataItem = $Script:datagridview1.Rows[$i].DataBoundItem
#$Script:Compare
IF($dataItem.Name -eq "Shortcuts" -and $dataItem.Difference -match ".lnk"){
$Script:datagridview1.Rows[$i].Cells["Selected"].Value = $True
}
elseif($dataItem.Name -eq "FTAs" -and $dataItem.Difference[0] -eq "."){
$Script:datagridview1.Rows[$i].Cells["Selected"].Value = $True
}
elseif($dataItem.Name -eq "Environmentals"){
Write-host "Yes"
$Script:datagridview1.Rows[$i].Cells["Selected"].Value = $True
}
elseif($dataItem.Name -eq "FireWallRules"){
Write-host "Yes"
$Script:datagridview1.Rows[$i].Cells["Selected"].Value = $True
}
elseif($dataItem.Name -eq "Drivers"){
Write-host "Yes"
$Script:datagridview1.Rows[$i].Cells["Selected"].Value = $True
}

$i++
}


}

Function Add-Exclusions {

if($script:Compare -ne $null){

$Final = @()
$FinalExcluded = @()

$ExcludedProcesses =@()
$Processes = $script:Compare|Where-Object -Property Name -eq Process
foreach($item in ($Script:Exclusions|Where-Object -Property Name -eq Process).Exclusion){
$item = $item.Replace("\","\\")
$ExcludedProcesses +=  $Processes |Where-Object{$_.Difference -eq $item}
$Processes =  $Processes |Where-Object{$_.Difference -ne $item}

}
$Final += $Processes
$FinalExcluded += $ExcludedProcesses

$ExcludedServices =@()
$Services = $script:Compare|Where-Object -Property Name -eq Services
foreach($item in ($Script:Exclusions|Where-Object -Property Name -eq Services).Exclusion){
$item = $item.Replace("\","\\")
$ExcludedServices +=  $Services |Where-Object{$_.Difference -ne $item}
$Services =  $Services |Where-Object{$_.Difference -ne $item}

}
$Final += $Services
$FinalExcluded += $ExcludedServices

$Files = $script:Compare|Where-Object -Property Name -eq Files

IF(($Script:Exclusions|Where-Object -Property Name -eq Files).Exclusion.length -eq 0 ){
#Replace ENV Vars Top
$FinalNames = @()
$allEnvironmentals = $Script:SecondCapture[11]
Foreach($FullPath in $Files){
$match = $false
foreach($item1 in $allEnvironmentals){
IF($match -ne $True){
Try{$StartString = $FullPath.Difference.Substring(0,$item1.Value.length)}catch{}
If($StartString -eq $item1.Value ){

$FinalNames += New-Object PsObject -Property @{Difference = "%$($item1.Name)%" + "$($FullPath.Difference.substring($item1.Value.length))"; Name = "Files";  }
$match = $true
}
$StartString = $null
}
}
IF($match -eq $false){$FinalNames += New-Object PsObject -Property @{Difference = "$($FullPath.Difference)"; Name = "Files"; }}
}

}else{

#Replace ENV Vars Top
$FinalNames = @()
$allEnvironmentals = $Script:SecondCapture[11]
Foreach($FullPath in $Files){
$match = $false
foreach($item1 in $allEnvironmentals){
IF($match -ne $True){
Try{$StartString = $FullPath.Difference.Substring(0,$item1.Value.length)}catch{}
If($StartString -eq $item1.Value ){

$FinalNames += New-Object PsObject -Property @{Difference = "%$($item1.Name)%" + "$($FullPath.Difference.substring($item1.Value.length))"; Name = "Files";  }
$match = $true
}
$StartString = $null
}
}
IF($match -eq $false){$FinalNames += New-Object PsObject -Property @{Difference = "$($FullPath.Difference)"; Name = "Files"; }}
}

$ExcludedFinalNames =@()
foreach($item in ($Script:Exclusions|Where-Object -Property Name -eq Files).Exclusion){

$count = "$item".Length

$ExcludedFinalNames +=  $FinalNames |Where-Object{"$(if("$($_.Difference)".Length -ge $count){"$($_.Difference)".Substring(0,$count)}else{} )" -eq "$item"}
$FinalNames =  $FinalNames |Where-Object{"$(if("$($_.Difference)".Length -ge $count){"$($_.Difference)".Substring(0,$count)}else{} )" -ne "$item"}

#Replace ENV Vars Bottom

}
}
$Final += $FinalNames
$FinalExcluded += $ExcludedFinalNames

$ExcludedRegistry =@()
$Registry = $script:Compare|Where-Object -Property Name -eq Registry
foreach($item in ($Script:Exclusions|Where-Object -Property Name -eq Registry).Exclusion){
$item = $item.Replace("\","\\")
$ExcludedRegistry +=  $Registry |Where-Object{$_.Difference -match $item}
$Registry =  $Registry |Where-Object{$_.Difference -notmatch $item}
}
$Final += $Registry
$FinalExcluded += $ExcludedRegistry

$ExcludedShortcuts =@()
$Shortcuts = $script:Compare|Where-Object -Property Name -eq Shortcuts
foreach($item in ($Script:Exclusions|Where-Object -Property Name -eq Shortcuts).Exclusion){
$item = $item.Replace("\","\\")
$ExcludedShortcuts +=  $Shortcuts |Where-Object{$_.Difference -match $item}
$Shortcuts =  $Shortcuts |Where-Object{$_.Difference -notmatch $item}
}
$Final += $Shortcuts
$FinalExcluded += $ExcludedShortcuts

$ExcludedDrivers =@()
$Drivers = $script:Compare|Where-Object -Property Name -eq Drivers
foreach($item in ($Script:Exclusions|Where-Object -Property Name -eq Drivers).Exclusion){
$item = $item.Replace("\","\\")
$ExcludedDrivers +=  $Drivers |Where-Object{$_.Difference -match $item}
$Drivers =  $Drivers |Where-Object{$_.Difference -notmatch $item}
}
$Final += $Drivers
$FinalExcluded += $ExcludedDrivers

$ExcludedFTAs =@()
$FTAs = $script:Compare|Where-Object -Property Name -eq FTAs
foreach($item in ($Script:Exclusions|Where-Object -Property Name -eq FTAs).Exclusion){
$item = $item.Replace("\","\\")
$ExcludedFTAs +=  $FTAs |Where-Object{$_.Difference -match $item}
$FTAs =  $FTAs |Where-Object{$_.Difference -notmatch $item}
}
$Final += $FTAs
$FinalExcluded += $ExcludedFTAs

$ExcludedFirewallRules=@()
$FirewallRules = $script:Compare|Where-Object -Property Name -eq FireWallRules
foreach($item in ($Script:Exclusions|Where-Object -Property Name -eq FireWallRules).Exclusion){
$item = $item.Replace("\","\\")
$ExcludedFirewallRules +=  $FirewallRules |Where-Object{$_.Difference -match $item}
$FirewallRules =  $FirewallRules |Where-Object{$_.Difference -notmatch $item}
}
$Final += $FirewallRules
$FinalExcluded += $ExcludedFirewallRules

$ExcludedODBCs =@()
$ODBCs = $script:Compare|Where-Object -Property Name -eq ODBCs
foreach($item in ($Script:Exclusions|Where-Object -Property Name -eq ODBCs).Exclusion){
$item = $item.Replace("\","\\")
$ExcludedODBCs +=  $ODBCs |Where-Object{$_.Difference -match $item}
$ODBCs =  $ODBCs |Where-Object{$_.Difference -notmatch $item}
}
$Final += $ODBCs
$FinalExcluded += $ExcludedODBCs

$ExcludedEnvironmentals =@()
$Environmentals = $script:Compare|Where-Object -Property Name -eq Environmentals
foreach($item in ($Script:Exclusions|Where-Object -Property Name -eq Environmentals).Exclusion){
$item = $item.Replace("\","\\")
$ExcludedEnvironmentals +=  $Environmentals |Where-Object{$_.Difference -eq $item}
$Environmentals =  $Environmentals |Where-Object{$_.Difference -ne $item}
}
$Final += $Environmentals
$FinalExcluded += $ExcludedEnvironmentals

#$FinalExcluded = $FinalExcluded|Where-Object {$_.Name = "Excluded:$($_.Name)" }

$Script:Selected.Visible = $True
if($toolstripcombobox3.selecteditem -eq "Filtered"){
$Script:datagridview1.DataSource=[collections.arraylist]@($Final|where -Property "Difference" -ne $null)

}elseif($toolstripcombobox3.selecteditem -eq "Exclusions"){
$Script:ExclusionsDone =@()
if("$Script:Exclusions".Length -gt 0){
foreach($itemexclusion in $Script:Exclusions){
if($itemexclusion.Exclusion -ne ""){
[array]$Script:ExclusionsDone += New-Object PsObject -Property @{Exclusion = $itemexclusion.Exclusion; Name = "Exclusion:$($itemexclusion.Name)";}
}
}
}
$Script:datagridview1.DataSource=[collections.arraylist]@($Script:ExclusionsDone|Select-Object "Exclusion","Name")
$Script:Selected.Visible = $False


}else{
$FinalExcluded1 =@()
if("$FinalExcluded".Length -gt 0){
foreach($itemexclusion in $FinalExcluded){
if($itemexclusion.Difference -ne "" ){
$FinalExcluded1 += New-Object PsObject -Property @{Difference = $itemexclusion.Difference; Name = "Exclusion:$($itemexclusion.Name)"}
}
}
}
$Script:datagridview1.DataSource=[collections.arraylist]@(($Final += $FinalExcluded1 )|where -Property "Difference" -ne $null)

}




}elseif($toolstripcombobox3.selecteditem -eq "Exclusions"){
if("$Script:Exclusions".Length -gt 0){
foreach($itemexclusion in $Script:Exclusions){
[array]$Script:ExclusionsDone += New-Object PsObject -Property @{Exclusion = $itemexclusion.Exclusion; Name = "Exclusion:$($itemexclusion.Name)"}
}
}
$Script:datagridview1.DataSource=[collections.arraylist]@($Script:ExclusionsDone|Select-Object "Exclusion","Name")
$Script:Selected.Visible = $False


}else{
$Script:datagridview1.DataSource=[collections.arraylist]@()
$Script:Selected.Visible = $True
}
foreach($item in $script:dataGridView1.Columns.Name){
if($item -ne "Selected"){
$script:dataGridView1.Columns["$item"].ReadOnly = $True
}
}
}

Function ValidationCredentials{

IF($usernametoolstriptextbox2.text -eq ""){
$usernametoolstriptextbox2.text = Read-Host -Prompt "VM Username"
}
IF($script:password -eq ""){
$script:password = Read-Host -Prompt "VM Password"
$toolstriptextbox3.Text = "*" * "$script:Password".length
}
$Script:DefaultVMUserName = $usernametoolstriptextbox2.text
$Script:DefaultVMPassword = $script:password

}

function Set-DataGridViewDoubleBuffer {
	param (
		[Parameter(Mandatory = $true)][System.Windows.Forms.DataGridView]$grid,
		[Parameter(Mandatory = $true)][boolean]$Enabled
	)
	$type = $grid.GetType();
	$propInfo = $type.GetProperty('DoubleBuffered',('Instance','NonPublic'))
	$propInfo.SetValue($grid, $Enabled, $null)
}

Function Invoke-MessageBox {
   [CmdletBinding(DefaultParameterSetName='DefaultIcon1')]
    Param(  
    
        [Parameter(Mandatory=$True,
                   HelpMessage = "Place Any String Here")]
                   [ValidateScript({$_.length -in 1..50})]
        [String]$FormHeader,
        
        [Parameter(Mandatory=$True,
                   HelpMessage = "Place Any String Here")]
        [String]$Message,

        [Parameter(Mandatory=$False,
                   HelpMessage = "Form Background Colour")]
        [ValidateSet('black','white','red','lime','blue','yellow','fuchsia','aqua','maroon','green','navy','olive','purple','teal','silver','gray')]
        [String]$BackColour = "White",

        [Parameter(Mandatory=$False,
                   HelpMessage = "Form Style")]
        [ValidateSet('None','FixedSingle','Fixed3D','FixedDialog','Sizable','FixedToolWindow','SizableToolWindow')]
        [String]$FormStyle = "FixedSingle",

        [Parameter(Mandatory=$False,
                   HelpMessage = "Text Size")]
        [ValidateScript({$_ -in 1..72})]
        [int]$FontSize = 12,

        [Parameter(Mandatory=$False,
                   HelpMessage = "Text to be displayed as the first button, this will also be the returned value when the button is clicked")]
                   [ValidateScript({$_.length -in 1..30})]
        [String]$ButtonOne = "Accept",

        [Parameter(Mandatory=$False,
                   HelpMessage = "Text to be displayed as the Second button, this will also be the returned value when the button is clicked")]
                   [ValidateScript({$_.length -in 1..30})]
        [String]$ButtonTwo,

        [Parameter(Mandatory=$False,
                   HelpMessage = "Do you want a commant box or not.")]
                   [ValidateSet('Yes','No')]
        [String]$DisplayCommentBox = 'No',
  
        [Parameter(Mandatory = $False, ParameterSetName = 'DefaultIcon1')]
        [ValidateSet('Information','Exclamation','Question','Error','Warning')]
        [String]$DefaultIcon = "Information",

        [Parameter(Mandatory = $False,
                   HelpMessage = "Enter a base 64 ico string or path to .ico file",
                   ParameterSetName="CustomIcon1")]
        [String]$CustomIcon,

        [Parameter(Mandatory = $False,
                   HelpMessage = "Time in minutes before the form closes automatically")]
        [ValidateScript({$_.length -in 0.1..10000})]
        [decimal]$TimeOut_Minutes,

        [Parameter(Mandatory = $False,
                   HelpMessage = "Do you wish to show the time remaining until the form closes automatically on the hearder?")]
        [ValidateSet('Yes','No')]
        [String]$ShowTimeOutRemaining = 'No',

        [Parameter(Mandatory = $False,
                   HelpMessage = "Make this window always on top.")]
        [ValidateSet('True','False')]
        [String]$Topmost = 'False'
    
    
    )

    DynamicParam {
            # Set the dynamic parameters' name
            $TextFont = 'TextFont'
            
            # Create the dictionary 
            $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

            # Create the collection of attributes
            $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            
            # Create and set the parameters' attributes
            $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
            $ParameterAttribute.Mandatory = $False
            $ParameterAttribute.Position = 1

            # Add the attributes to the attributes collection
            $AttributeCollection.Add($ParameterAttribute)

            # Generate and set the ValidateSet 
            [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
            $objFonts = New-Object System.Drawing.Text.InstalledFontCollection
            $arrSet = $objFonts.Families.name|foreach{$_ = "$_".Replace(" ","")
            $_}

            $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)

            # Add the ValidateSet to the attributes collection
            $AttributeCollection.Add($ValidateSetAttribute)
            #$AttributeCollection.Validvalues

            # Create and return the dynamic parameter
            $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TextFont, [string],$AttributeCollection)
            $RuntimeParameter.Value = "Calibri"
            $RuntimeParameterDictionary.Add($TextFont, $RuntimeParameter)
            return $RuntimeParameterDictionary
    }

BEGIN {
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
        $objFonts = New-Object System.Drawing.Text.InstalledFontCollection
        $Fonts = $objFonts.Families.name
        $TextFont = $PsBoundParameters[$TextFont]
        $TextFont = $Fonts |Where-Object {"$_".Replace(" ","") -eq $TextFont}
        IF($TextFont.count -ne 1){$TextFont = "Calibri"}
        
        $FormTimeoutDate = (Get-date).AddMinutes($TimeOut_Minutes)

    }


PROCESS {



function OnApplicationLoad {
	#Note: This function is not called in Projects
	#Note: This function runs before the form is created
	#Note: To get the script directory in the Packager use: Split-Path $hostinvocation.MyCommand.path
	#Note: To get the console output in the Packager (Windows Mode) use: $ConsoleOutput (Type: System.Collections.ArrayList)
	#Important: Form controls cannot be accessed in this function
	#TODO: Add snapins and custom code to validate the application load
	
	return $true #return true for success or false for failure
}

function OnApplicationExit {
	#Note: This function is not called in Projects
	#Note: This function runs after the form is closed
	#TODO: Add custom code to clean up and unload snapins when the application exits
	
	$script:ExitCode = 0 #Set the exit code for the Packager
}

#----------------------------------------------
# Generated Form Function
#----------------------------------------------
function Call-MessageBox_pff {

	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System.Core, Version=3.5.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.ServiceProcess, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$FormPrompt = New-Object 'System.Windows.Forms.Form'
	$panelLabels = New-Object 'System.Windows.Forms.Panel'
	$labelPleaseEnterInformati = New-Object 'System.Windows.Forms.Label'
	$panelComment = New-Object 'System.Windows.Forms.Panel'
	$textbox1 = New-Object 'System.Windows.Forms.TextBox'
	$PanelButtons = New-Object 'System.Windows.Forms.Panel'
	$buttonYes = New-Object 'System.Windows.Forms.Button'
	$buttonNo = New-Object 'System.Windows.Forms.Button'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	$timer1 = New-Object 'System.Windows.Forms.Timer'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
	
	$FormPrompt_Load={
		
	$Script:MessageReturn = $null
    $Script:MessageCommentBox = $null
	$font = New-Object System.Drawing.Font("$TextFont", $FontSize, [System.Drawing.FontStyle]'Regular')
	$size = [System.Windows.Forms.TextRenderer]::MeasureText("$($labelPleaseEnterInformati.Text)", $font)
	IF($size.Width -lt 400){
	$FormPrompt.Width = 400		
	} elseif($size.Width -gt 1500){
	$FormPrompt.Width = 1500	
	}else{
	$FormPrompt.Width = $size.Width
		}
	
	if($DisplayCommentBox -eq "Yes"){


	if($size.height + 180 -gt 1500){
	$FormPrompt.height = 1500	
	}else{
	$FormPrompt.height = $size.height + 180
		}
	}else{

	if($size.height + 100 -gt 1500){
	$FormPrompt.height = 1500	
	}else{
	$FormPrompt.height = $size.height + 100
		}



    }
		#TODO: Initialize Form Controls here
	$timer1.Enabled = $true
	}
	
	$PanelButtons_Resize={

    If($ButtonTwo -ne ""){
		#TODO: Place custom script here
		#$position $panel1.Width
		$buttonNo.width = ($PanelButtons.Width / 2 ) - 11
		$buttonNo.Location.X = ($PanelButtons.Width / 2 ) + 5
		$buttonYes.Width = ($PanelButtons.Width / 2 ) - 11
		$buttonYes.Location.X = ($PanelButtons.Location.X) + 5
		#$panel1
     }else{
        $buttonYes.Width = ($PanelButtons.Width ) - 11
        }
	}


    $timer1_Tick={
    IF($TimeOut_Minutes -ne ""){
    $FormTimeout = (Get-Date -Date $FormTimeoutDate).TimeOfDay.TotalSeconds - (Get-date).TimeOfDay.TotalSeconds
    $ts2 =  [timespan]::fromseconds($FormTimeout)
    If($ShowTimeOutRemaining -eq 'Yes'){
    $FormPrompt.Text = "$FormHeader |Time Remaining: $($ts2.minutes):$("{0:D2}" -f $ts2.seconds)"
    }
    IF($FormTimeout -lt 0){
    $Script:MessageReturn = "TimeOut"
    $FormPrompt.Close()}
    }
	#TODO: Place custom script here
	
    }

    $buttonYes_Click={
    $Script:MessageCommentBox = $textbox1.text
    $Script:MessageReturn = $buttonYes.Text
    $FormPrompt.Close()
	#TODO: Place custom script here
	
    }

    $buttonNo_Click={
    $Script:MessageCommentBox = $textbox1.text
    $Script:MessageReturn = $buttonNo.Text
    $FormPrompt.Close()
	#TODO: Place custom script here
	
    }
	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$FormPrompt.WindowState = $InitialFormWindowState
	}
	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$PanelButtons.remove_Resize($PanelButtons_Resize)
			$FormPrompt.remove_Load($FormPrompt_Load)
			$FormPrompt.remove_Load($Form_StateCorrection_Load)
			$FormPrompt.remove_FormClosed($Form_Cleanup_FormClosed)
			$timer1.remove_Tick($timer1_Tick)
			$buttonYes.remove_Click($buttonYes_Click)
			$buttonNo.remove_Click($buttonNo_Click)
		}
		catch [Exception]
		{ }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	#
	# FormPrompt
	#
	$FormPrompt.Controls.Add($panelLabels)
	$FormPrompt.Controls.Add($panelComment)
	$FormPrompt.Controls.Add($PanelButtons)
	$FormPrompt.BackColor = "$BackColour"
	$FormPrompt.ClientSize = '518, 224'
	$FormPrompt.Font = "$TextFont, 12pt"
    $FormPrompt.FormBorderStyle = "$FormStyle"
    $FormPrompt.TopMost = $Topmost

    if($CustomIcon -ne ""){

        $FormIcon = $CustomIcon

        }else{
        Switch ($DefaultIcon) {

        Information {$FormPrompt.Icon = [System.Drawing.SystemIcons]::Information}
        Exclamation {$FormPrompt.Icon = [System.Drawing.SystemIcons]::Exclamation}
        Question    {$FormPrompt.Icon = [System.Drawing.SystemIcons]::Question   }
        Error       {$FormPrompt.Icon = [System.Drawing.SystemIcons]::Error      }
        Warning     {$FormPrompt.Icon = [System.Drawing.SystemIcons]::Warning    }
        
        }
     
       }

	$FormPrompt.MinimumSize = '400, 100'
	$FormPrompt.Name = "FormPrompt"

	$FormPrompt.Text = $FormHeader
    IF($ShowTimeOutRemaining -eq 'Yes'){
    $FormPrompt.Text = "$FormHeader |Time Remaining: 0"
    }
	$FormPrompt.add_Load($FormPrompt_Load)
	#
	# panelLabels
	#
	$panelLabels.Controls.Add($labelPleaseEnterInformati)
	$panelLabels.AutoSizeMode = 'GrowAndShrink'
	$panelLabels.Dock = 'Fill'
	$panelLabels.Location = '0, 0'
	$panelLabels.Name = "panelLabels"
	$panelLabels.Padding = '5, 5, 5, 5'
	$panelLabels.Size = '518, 108'
	$panelLabels.TabIndex = 4
	#
	# labelPleaseEnterInformati
	#
	$labelPleaseEnterInformati.BorderStyle = 'FixedSingle'
	$labelPleaseEnterInformati.Dock = 'Fill'
	$labelPleaseEnterInformati.Location = '5, 5'
    $labelPleaseEnterInformati.Font = "$TextFont, $FontSize`pt"
	$labelPleaseEnterInformati.Name = "labelPleaseEnterInformati"
	$labelPleaseEnterInformati.Size = '508, 98'
	$labelPleaseEnterInformati.TabIndex = 3
	$labelPleaseEnterInformati.Text = $Message
	$labelPleaseEnterInformati.TextAlign = 'MiddleCenter'
	#
	# panelComment
	#
	$panelComment.Controls.Add($textbox1)
	$panelComment.AutoSizeMode = 'GrowAndShrink'
	$panelComment.Dock = 'Bottom'
	$panelComment.Location = '0, 108'
	$panelComment.Name = "panelComment"
	$panelComment.Size = '518, 76'
	$panelComment.TabIndex = 5
    IF($DisplayCommentBox -eq 'No'){$panelComment.Visible = $False}
	#
	# textbox1
	#
	$textbox1.Anchor = 'Top, Bottom, Left, Right'
	$textbox1.BorderStyle = 'FixedSingle'
	$textbox1.Location = '5, 3'
	$textbox1.Multiline = $True
	$textbox1.Name = "textbox1"
	$textbox1.ScrollBars = 'Vertical'
	$textbox1.Size = '507, 67'
	$textbox1.TabIndex = 0
	#
	# PanelButtons
	#
	$PanelButtons.Controls.Add($buttonYes)
	$PanelButtons.Controls.Add($buttonNo)
	$PanelButtons.AutoSizeMode = 'GrowAndShrink'
	$PanelButtons.Dock = 'Bottom'
	$PanelButtons.Location = '0, 184'
	$PanelButtons.Name = "PanelButtons"
	$PanelButtons.Size = '518, 40'
	$PanelButtons.TabIndex = 2
	$PanelButtons.add_Resize($PanelButtons_Resize)
	#
	# buttonYes
	#
	$buttonYes.AutoSizeMode = 'GrowAndShrink'
	$buttonYes.FlatAppearance.MouseDownBackColor = '128, 255, 128'
	$buttonYes.FlatAppearance.MouseOverBackColor = '192, 255, 192'
	$buttonYes.FlatStyle = 'Flat'
	$buttonYes.Location = '5, 6'
	$buttonYes.Name = "buttonYes"
	$buttonYes.Size = '254, 28'
	$buttonYes.TabIndex = 0
	$buttonYes.Text = "$ButtonOne"
	$buttonYes.UseVisualStyleBackColor = $True
	$buttonYes.add_Click($buttonYes_Click)
	#
	# buttonNo
	#
	$buttonNo.Anchor = 'Top'
	$buttonNo.AutoSizeMode = 'GrowAndShrink'
	$buttonNo.FlatAppearance.MouseDownBackColor = '255, 128, 128'
	$buttonNo.FlatAppearance.MouseOverBackColor = '255, 192, 192'
	$buttonNo.FlatStyle = 'Flat'
	$buttonNo.Location = '264, 6'
	$buttonNo.Name = "buttonNo"
	$buttonNo.Size = '248, 28'
	$buttonNo.TabIndex = 1
	$buttonNo.Text = "$ButtonTwo"
	$buttonNo.UseVisualStyleBackColor = $True
	$buttonNo.add_Click($buttonNo_Click)
    If($ButtonTwo -eq ""){$buttonNo.Visible = $False
    $buttonYes.Size = '518, 28'
    }
    #
	# timer1
	#
	$timer1.add_Tick($timer1_Tick)
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $FormPrompt.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$FormPrompt.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$FormPrompt.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	Return $FormPrompt.ShowDialog()


} #End Function

#Call OnApplicationLoad to initialize
if((OnApplicationLoad) -eq $true)
{
	#Call the form
	Call-MessageBox_pff|Out-Null
	#Perform cleanup
	OnApplicationExit
}




}

END {Return $Script:MessageReturn,$Script:MessageCommentBox}

}

Function VMScreenCapture {

    [CmdletBinding()]
    Param(       
                      
        [Parameter(Mandatory=$False,
                   ValueFromPipeline=$True,
                   ValueFromPipelineByPropertyName=$True,
                   HelpMessage = "Location of Running VM VMX file")]
        [ValidateScript({(Test-Path $_) -and ((Get-Item $_).Extension -eq ".vmx")})] 
        [Alias('Path')]
        [String]$VMXpath = $DefaultVMXPath,

        [Parameter(Mandatory=$False,
                   HelpMessage = "User Name for Target VMX")]
        [String]$VMUserName = $Script:DefaultVMUserName,

        [Parameter(Mandatory=$False,
                   HelpMessage = "Password for Target VMX")]
        [String]$VMPassword = $Script:DefaultVMPassword,

        [Parameter(Mandatory=$False,
                   HelpMessage = "Image Background Colour")]
        [ValidateSet("Windowed","FullScreen")]
        [String]$Type = "Windowed"

    
    )
Begin {
$HostPath = $ENV:Temp.Replace(":","")
$error.clear()
Test-path "$ENV:Temp\VMRun\CaptureWindowed.PS1" -ErrorAction SilentlyContinue 
if("$error" -match "Access is denied"){
Invoke-MessageBox -FormHeader "VMWARE File Locked" -Message "Please Stop your VMs and close VMware then reload to resolve this issue." -ButtonOne "Okay"
  }  
}

Process {
$Temp = $ENV:Temp.Replace("C:","Z:\C")

IF($Type -eq "Windowed"){
$command = '
        [String]$Script:SaveToLocation = "' + $Temp + '\VMRun\CapturedScreen.PNG"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
        Add-Type -AssemblyName System.Drawing
        $Script:PNGCodec = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.FormatDescription -eq "PNG" }
        [Windows.Forms.Sendkeys]::SendWait("%{PrtSc}")
        Start-Sleep -Milliseconds 250
        $Script:bitmap = [Windows.Forms.Clipboard]::GetImage()    
        $Script:ep = New-Object Drawing.Imaging.EncoderParameters  
        $Script:ep.Param[0] = New-Object Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality, [long]100)  
        $Script:bitmap.Save("$Script:SaveToLocation", $Script:PNGCodec, $Script:ep) '

}else{
$command = '
        [String]$SaveToLocation = "' + $Temp + '\VMRun\CapturedScreen.PNG"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
        Add-Type -AssemblyName System.Drawing
        $PNGCodec = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.FormatDescription -eq "PNG" }
        [Windows.Forms.Sendkeys]::SendWait("{PrtSc}")
        Start-Sleep -Milliseconds 250
        $bitmap = [Windows.Forms.Clipboard]::GetImage()    
        $ep = New-Object Drawing.Imaging.EncoderParameters  
        $ep.Param[0] = New-Object Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality, [long]100)  
        $bitmap.Save("$SaveToLocation", $PNGCodec, $ep) '

}


IF($type -eq "Windowed"){
IF(!(Test-path "$ENV:Temp\VMRun\CaptureWindowed.PS1")){

$Command | Out-File -FilePath "$ENV:Temp\VMRun\CaptureWindowed.PS1" -Encoding ascii -Force
}

& "$VMRunPath" -gu "$VMUserName" -gp "$VMPassword" runProgramInGuest "$VMXpath" -interactive "$env:windir\system32\windowspowershell\v1.0\PowerShell.exe" -sta -executionpolicy bypass -file "Z:\$HostPath\VMRun\CaptureWindowed.PS1"
}else{

IF(!(Test-path "$ENV:Temp\VMRun\CaptureFullScreen.PS1")){

$Command | Out-File -FilePath "$ENV:Temp\VMRun\CaptureFullScreen.PS1" -Encoding ascii -Force
}

& "$VMRunPath" -gu "$VMUserName" -gp "$VMPassword" runProgramInGuest "$VMXpath" -interactive "$env:windir\system32\windowspowershell\v1.0\PowerShell.exe" -sta -executionpolicy bypass -file "Z:\$HostPath\VMRun\CaptureFullScreen.PS1"


}

}


}

Function PassCapToDoc {
Param(  [Parameter(Mandatory=$False,
        HelpMessage = "Capture Type windowed or full screen")]
        [ValidateSet("Windowed","FullScreen")]
        [String]$Type = "Windowed")


ValidationCredentials


Try{
$captured = $true
#$Type = "Windowed"
VMScreenCapture -VMXpath "$($Script:VMselectiontoolstripcombobox4.SelectedItem)" -Type $type
$capture = "$env:Temp\VMRun\CapturedScreen.PNG"
}Catch{$captured = $false}


IF(($captured -eq $true) -and (Test-path "$env:Temp\VMRun\CapturedScreen.PNG")){
If($toolstripcombobox2.Text -eq "Installation Steps"){

if($Script:ISobjRange -eq $null){
$Script:ISobjRange = $Script:Doc.BookMarks.Item("INSTALLATION_STEPS_DISCOTOOL").Range
}
    IF($Script:ISobjRange.Text -match "Step by step instructions detailing how the installation should be performed."){
    $Script:ISobjRange.Text = "Installation of Application: Version:"
    $Script:ISobjRange.font.Size = 16
    $Script:ISobjRange.font.Bold = 1
    }
$Script:ISobjRange.Select()
$Script:ISobjRange.Start = $Script:ISobjRange.End
$Script:ISobjRange.Text = "

"
$Script:ISobjRange.font.Size = 10
$Script:ISobjRange.font.Bold = 0

$Script:ISobjRange.Start = $Script:ISobjRange.End
$imagetoscale = $Script:ISobjRange.InlineShapes.AddPicture("$capture")
$imagetoscale.LockAspectRatio = 1
$imagetoscale.Range.ParagraphFormat.Alignment  = 1
$imagetoscale.Width = 365
$Script:ISobjRange.End = $Script:ISobjRange.End + 1
$Script:ISobjRange.Start = $Script:ISobjRange.End
$Script:ISobjRange.text = "
"
$Script:ISobjRange.Start = $Script:ISobjRange.End
$Script:ISobjRange.Select()
IF($DefaultTexttoolstripcombobox1.Text -ne "NULL NO TEXT WITH SCREEN CAPTURE"){
$quotecount = 0

$string = $($DefaultTexttoolstripcombobox1.Text).Replace("Default Text: ","")
$stringarray = $string.Split('"')
foreach($item in $stringarray){
IF($quotecount -eq 0){
$Script:ISobjRange.Text = "$item"
$Script:ISobjRange.font.Bold = 0
$quotecount = 1
$Script:ISobjRange.Start = $Script:ISobjRange.End
}elseif($quotecount -eq 1){
$Script:ISobjRange.Text = '"'
$Script:ISobjRange.font.Bold = 0
$Script:ISobjRange.Start = $Script:ISobjRange.End
$Script:ISobjRange.Text = "$item"
$Script:ISobjRange.font.Bold = 1
$Script:ISobjRange.Start = $Script:ISobjRange.End
$Script:ISobjRange.Text = '"'
$Script:ISobjRange.font.Bold = 0
$Script:ISobjRange.Start = $Script:ISobjRange.End
$quotecount = 0
}
}



}

}

elseIf($toolstripcombobox2.Text -eq "Start Menu"){

if($Script:FLobjRange -eq $null){
$Script:FLobjRange = $Script:Doc.BookMarks.Item("START_MENU_DISCOTOOL").Range
}

$Script:FLobjRange.Select()
$Script:FLobjRange.Text = "

"
$Script:FLobjRange.font.Size = 10
$Script:FLobjRange.font.Bold = 0

$Script:FLobjRange.Start = $Script:FLobjRange.End
$imagetoscale = $Script:FLobjRange.InlineShapes.AddPicture("$capture")
$imagetoscale.LockAspectRatio = 1
$imagetoscale.Range.ParagraphFormat.Alignment  = 1
$imagetoscale.Width = 365
$Script:FLobjRange.End = $Script:FLobjRange.End + 1
$Script:FLobjRange.Start = $Script:FLobjRange.End
$Script:FLobjRange.text = "
"
$Script:FLobjRange.Start = $Script:FLobjRange.End
$Script:FLobjRange.Select()
IF($DefaultTexttoolstripcombobox1.Text -ne "NULL NO TEXT WITH SCREEN CAPTURE"){
$quotecount = 0

$string = $($DefaultTexttoolstripcombobox1.Text).Replace("Default Text: ","")
$stringarray = $string.Split('"')
foreach($item in $stringarray){
IF($quotecount -eq 0){
$Script:FLobjRange.Text = "$item"
$Script:FLobjRange.font.Bold = 0
$quotecount = 1
$Script:FLobjRange.Start = $Script:FLobjRange.End
}elseif($quotecount -eq 1){
$Script:FLobjRange.Text = '"'
$Script:FLobjRange.font.Bold = 0
$Script:FLobjRange.Start = $Script:FLobjRange.End
$Script:FLobjRange.Text = "$item"
$Script:FLobjRange.font.Bold = 1
$Script:FLobjRange.Start = $Script:FLobjRange.End
$Script:FLobjRange.Text = '"'
$Script:FLobjRange.font.Bold = 0
$Script:FLobjRange.Start = $Script:FLobjRange.End
$quotecount = 0
}
}



}

}

elseIf($toolstripcombobox2.Text -eq "First Launch Test & Post Configuration"){

if($Script:FLobjRange -eq $null){
$Script:FLobjRange = $Script:Doc.BookMarks.Item("FIRST_LAUNCH_DISCOTOOL").Range
}

$Script:FLobjRange.Select()
$Script:FLobjRange.Text = "

"
$Script:FLobjRange.font.Size = 10
$Script:FLobjRange.font.Bold = 0

$Script:FLobjRange.Start = $Script:FLobjRange.End
$imagetoscale = $Script:FLobjRange.InlineShapes.AddPicture("$capture")
$imagetoscale.LockAspectRatio = 1
$imagetoscale.Range.ParagraphFormat.Alignment  = 1
$imagetoscale.Width = 365
$Script:FLobjRange.End = $Script:FLobjRange.End + 1
$Script:FLobjRange.Start = $Script:FLobjRange.End
$Script:FLobjRange.text = "
"
$Script:FLobjRange.Start = $Script:FLobjRange.End
$Script:FLobjRange.Select()
IF($DefaultTexttoolstripcombobox1.Text -ne "NULL NO TEXT WITH SCREEN CAPTURE"){
$quotecount = 0

$string = $($DefaultTexttoolstripcombobox1.Text).Replace("Default Text: ","")
$stringarray = $string.Split('"')
foreach($item in $stringarray){
IF($quotecount -eq 0){
$Script:FLobjRange.Text = "$item"
$Script:FLobjRange.font.Bold = 0
$quotecount = 1
$Script:FLobjRange.Start = $Script:FLobjRange.End
}elseif($quotecount -eq 1){
$Script:FLobjRange.Text = '"'
$Script:FLobjRange.font.Bold = 0
$Script:FLobjRange.Start = $Script:FLobjRange.End
$Script:FLobjRange.Text = "$item"
$Script:FLobjRange.font.Bold = 1
$Script:FLobjRange.Start = $Script:FLobjRange.End
$Script:FLobjRange.Text = '"'
$Script:FLobjRange.font.Bold = 0
$Script:FLobjRange.Start = $Script:FLobjRange.End
$quotecount = 0
}
}



}

}

elseIf($toolstripcombobox2.Text -eq "About Menu"){

if($Script:AMobjRange -eq $null){
$Script:AMobjRange = $Script:Doc.BookMarks.Item("ABOUT_MENU_DISCOTOOL").Range
}

$Script:AMobjRange.Select()
$Script:AMobjRange.Text = "

"
$Script:AMobjRange.font.Size = 10
$Script:AMobjRange.font.Bold = 0

$Script:AMobjRange.Start = $Script:AMobjRange.End
$imagetoscale = $Script:AMobjRange.InlineShapes.AddPicture("$capture")
$imagetoscale.LockAspectRatio = 1
$imagetoscale.Range.ParagraphFormat.Alignment  = 1
$imagetoscale.Width = 365
$Script:AMobjRange.End = $Script:AMobjRange.End + 1
$Script:AMobjRange.Start = $Script:AMobjRange.End
$Script:AMobjRange.text = "
"
$Script:AMobjRange.Start = $Script:AMobjRange.End
$Script:AMobjRange.Select()
IF($DefaultTexttoolstripcombobox1.Text -ne "NULL NO TEXT WITH SCREEN CAPTURE"){
$quotecount = 0

$string = $($DefaultTexttoolstripcombobox1.Text).Replace("Default Text: ","")
$stringarray = $string.Split('"')
foreach($item in $stringarray){
IF($quotecount -eq 0){
$Script:AMobjRange.Text = "$item"
$Script:AMobjRange.font.Bold = 0
$quotecount = 1
$Script:AMobjRange.Start = $Script:AMobjRange.End
}elseif($quotecount -eq 1){
$Script:AMobjRange.Text = '"'
$Script:AMobjRange.font.Bold = 0
$Script:AMobjRange.Start = $Script:AMobjRange.End
$Script:AMobjRange.Text = "$item"
$Script:AMobjRange.font.Bold = 1
$Script:AMobjRange.Start = $Script:AMobjRange.End
$Script:AMobjRange.Text = '"'
$Script:AMobjRange.font.Bold = 0
$Script:AMobjRange.Start = $Script:AMobjRange.End
$quotecount = 0
}
}



}

}

elseIf($toolstripcombobox2.Text -eq "Help Menu"){

if($Script:HMobjRange -eq $null){
$Script:HMobjRange = $Script:Doc.BookMarks.Item("HELP_MENU_DISCOTOOL").Range
}

$Script:HMobjRange.Select()
$Script:HMobjRange.Text = "

"
$Script:HMobjRange.font.Size = 10
$Script:HMobjRange.font.Bold = 0

$Script:HMobjRange.Start = $Script:HMobjRange.End
$imagetoscale = $Script:HMobjRange.InlineShapes.AddPicture("$capture")
$imagetoscale.LockAspectRatio = 1
$imagetoscale.Range.ParagraphFormat.Alignment  = 1
$imagetoscale.Width = 365
$Script:HMobjRange.End = $Script:HMobjRange.End + 1
$Script:HMobjRange.Start = $Script:HMobjRange.End
$Script:HMobjRange.text = "
"
$Script:HMobjRange.Start = $Script:HMobjRange.End
$Script:HMobjRange.Select()
IF($DefaultTexttoolstripcombobox1.Text -ne "NULL NO TEXT WITH SCREEN CAPTURE"){
$quotecount = 0

$string = $($DefaultTexttoolstripcombobox1.Text).Replace("Default Text: ","")
$stringarray = $string.Split('"')
foreach($item in $stringarray){
IF($quotecount -eq 0){
$Script:HMobjRange.Text = "$item"
$Script:HMobjRange.font.Bold = 0
$quotecount = 1
$Script:HMobjRange.Start = $Script:HMobjRange.End
}elseif($quotecount -eq 1){
$Script:HMobjRange.Text = '"'
$Script:HMobjRange.font.Bold = 0
$Script:HMobjRange.Start = $Script:HMobjRange.End
$Script:HMobjRange.Text = "$item"
$Script:HMobjRange.font.Bold = 1
$Script:HMobjRange.Start = $Script:HMobjRange.End
$Script:HMobjRange.Text = '"'
$Script:HMobjRange.font.Bold = 0
$Script:HMobjRange.Start = $Script:HMobjRange.End
$quotecount = 0
}
}



}

}

elseIf($toolstripcombobox2.Text -eq "Control Panel Entry"){

if($Script:CPEobjRange -eq $null){
$Script:CPEobjRange = $Script:Doc.BookMarks.Item("CONTROL_PANEL_DISCOTOOL").Range
}

$Script:CPEobjRange.Select()
$Script:CPEobjRange.Text = "

"
$Script:CPEobjRange.font.Size = 10
$Script:CPEobjRange.font.Bold = 0

$Script:CPEobjRange.Start = $Script:CPEobjRange.End
$imagetoscale = $Script:CPEobjRange.InlineShapes.AddPicture("$capture")
$imagetoscale.LockAspectRatio = 1
$imagetoscale.Range.ParagraphFormat.Alignment  = 1
$imagetoscale.Width = 365
$Script:CPEobjRange.End = $Script:CPEobjRange.End + 1
$Script:CPEobjRange.Start = $Script:CPEobjRange.End
$Script:CPEobjRange.text = "
"
$Script:CPEobjRange.Start = $Script:CPEobjRange.End
$Script:CPEobjRange.Select()
IF($DefaultTexttoolstripcombobox1.Text -ne "NULL NO TEXT WITH SCREEN CAPTURE"){
$quotecount = 0

$string = $($DefaultTexttoolstripcombobox1.Text).Replace("Default Text: ","")
$stringarray = $string.Split('"')
foreach($item in $stringarray){
IF($quotecount -eq 0){
$Script:CPEobjRange.Text = "$item"
$Script:CPEobjRange.font.Bold = 0
$quotecount = 1
$Script:CPEobjRange.Start = $Script:CPEobjRange.End
}elseif($quotecount -eq 1){
$Script:CPEobjRange.Text = '"'
$Script:CPEobjRange.font.Bold = 0
$Script:CPEobjRange.Start = $Script:CPEobjRange.End
$Script:CPEobjRange.Text = "$item"
$Script:CPEobjRange.font.Bold = 1
$Script:CPEobjRange.Start = $Script:CPEobjRange.End
$Script:CPEobjRange.Text = '"'
$Script:CPEobjRange.font.Bold = 0
$Script:CPEobjRange.Start = $Script:CPEobjRange.End
$quotecount = 0
}
}



}

}

elseIf($toolstripcombobox2.Text -eq "Context Menu"){

if($Script:CMobjRange -eq $null){
$Script:CMobjRange = $Script:Doc.BookMarks.Item("CONTEXT_MENU_DISCOTOOL").Range
}

$Script:CMobjRange.Select()
$Script:CMobjRange.Text = "

"
$Script:CMobjRange.font.Size = 10
$Script:CMobjRange.font.Bold = 0

$Script:CMobjRange.Start = $Script:CMobjRange.End
$imagetoscale = $Script:CMobjRange.InlineShapes.AddPicture("$capture")
$imagetoscale.LockAspectRatio = 1
$imagetoscale.Range.ParagraphFormat.Alignment  = 1
$imagetoscale.Width = 365
$Script:CMobjRange.End = $Script:CMobjRange.End + 1
$Script:CMobjRange.Start = $Script:CMobjRange.End
$Script:CMobjRange.text = "
"
$Script:CMobjRange.Start = $Script:CMobjRange.End
$Script:CMobjRange.Select()
IF($DefaultTexttoolstripcombobox1.Text -ne "NULL NO TEXT WITH SCREEN CAPTURE"){
$quotecount = 0

$string = $($DefaultTexttoolstripcombobox1.Text).Replace("Default Text: ","")
$stringarray = $string.Split('"')
foreach($item in $stringarray){
IF($quotecount -eq 0){
$Script:CMobjRange.Text = "$item"
$Script:CMobjRange.font.Bold = 0
$quotecount = 1
$Script:CMobjRange.Start = $Script:CMobjRange.End
}elseif($quotecount -eq 1){
$Script:CMobjRange.Text = '"'
$Script:CMobjRange.font.Bold = 0
$Script:CMobjRange.Start = $Script:CMobjRange.End
$Script:CMobjRange.Text = "$item"
$Script:CMobjRange.font.Bold = 1
$Script:CMobjRange.Start = $Script:CMobjRange.End
$Script:CMobjRange.Text = '"'
$Script:CMobjRange.font.Bold = 0
$Script:CMobjRange.Start = $Script:CMobjRange.End
$quotecount = 0
}
}



}

}

elseIf($toolstripcombobox2.Text -eq "Functional Testing"){

if($Script:FTobjRange -eq $null){
$Script:FTobjRange = $Script:Doc.BookMarks.Item("FUNCTIONAL_TESTING_DISCOTOOL").Range
}

$Script:FTobjRange.Select()
$Script:FTobjRange.Text = "

"
$Script:FTobjRange.font.Size = 10
$Script:FTobjRange.font.Bold = 0

$Script:FTobjRange.Start = $Script:FTobjRange.End
$imagetoscale = $Script:FTobjRange.InlineShapes.AddPicture("$capture")
$imagetoscale.LockAspectRatio = 1
$imagetoscale.Range.ParagraphFormat.Alignment  = 1
$imagetoscale.Width = 365
$Script:FTobjRange.End = $Script:FTobjRange.End + 1
$Script:FTobjRange.Start = $Script:FTobjRange.End
$Script:FTobjRange.text = "
"
$Script:FTobjRange.Start = $Script:FTobjRange.End
$Script:FTobjRange.Select()
IF($DefaultTexttoolstripcombobox1.Text -ne "NULL NO TEXT WITH SCREEN CAPTURE"){
$quotecount = 0

$string = $($DefaultTexttoolstripcombobox1.Text).Replace("Default Text: ","")
$stringarray = $string.Split('"')
foreach($item in $stringarray){
IF($quotecount -eq 0){
$Script:FTobjRange.Text = "$item"
$Script:FTobjRange.font.Bold = 0
$quotecount = 1
$Script:FTobjRange.Start = $Script:FTobjRange.End
}elseif($quotecount -eq 1){
$Script:FTobjRange.Text = '"'
$Script:FTobjRange.font.Bold = 0
$Script:FTobjRange.Start = $Script:FTobjRange.End
$Script:FTobjRange.Text = "$item"
$Script:FTobjRange.font.Bold = 1
$Script:FTobjRange.Start = $Script:FTobjRange.End
$Script:FTobjRange.Text = '"'
$Script:FTobjRange.font.Bold = 0
$Script:FTobjRange.Start = $Script:FTobjRange.End
$quotecount = 0
}
}



}

}

elseIf($toolstripcombobox2.Text -eq "Place At Cursor"){


$Script:Word.Selection.TypeText("

")
$imagetoscale = $Script:Word.Selection.InlineShapes.AddPicture("$capture")
$imagetoscale.LockAspectRatio = 1
$imagetoscale.Range.ParagraphFormat.Alignment  = 1
$imagetoscale.Width = 365
$Script:Word.Selection.TypeText("
")

IF($DefaultTexttoolstripcombobox1.Text -ne "NULL NO TEXT WITH SCREEN CAPTURE"){
$quotecount = 0
$string = $($DefaultTexttoolstripcombobox1.Text).Replace("Default Text: ","")
$stringarray = $string.Split('"')
foreach($item in $stringarray){
IF($quotecount -eq 0){
$Script:Word.Selection.Font.Bold = 0
$Script:Word.Selection.TypeText("$item")
$Script:Word.Selection.Font.Bold = 0
$quotecount = 1

}elseif($quotecount -eq 1){
$Script:Word.Selection.Font.Bold = 0
$Script:Word.Selection.TypeText('"')
$Script:Word.Selection.Font.Bold = 1
$Script:Word.Selection.TypeText("$item")
$Script:Word.Selection.Font.Bold = 0
$Script:Word.Selection.TypeText('"')
$quotecount = 0
}
}



}

}
Del "$env:Temp\VMRun\CapturedScreen.PNG" -ErrorAction SilentlyContinue
}else{
$toolstriptextbox1.text = "No Screen Captured`nIs Z:\ drive is mapped? `nIs Username and Password correct?"
}
}


if(test-path -path "$VMRunPath"){
#========================================================================
# Code Generated By: SAPIEN Technologies, Inc., PowerShell Studio 2012 v3.1.21
# Generated On: 01/04/2017 15:04
# Generated By: Benjamin
#========================================================================
#----------------------------------------------
#region Application Functions
#----------------------------------------------

function OnApplicationLoad {
	#Note: This function is not called in Projects
	#Note: This function runs before the form is created
	#Note: To get the script directory in the Packager use: Split-Path $hostinvocation.MyCommand.path
	#Note: To get the console output in the Packager (Windows Mode) use: $ConsoleOutput (Type: System.Collections.ArrayList)
	#Important: Form controls cannot be accessed in this function
	#TODO: Add snapins and custom code to validate the application load
	
	return $true #return true for success or false for failure
}

function OnApplicationExit {
	#Note: This function is not called in Projects
	#Note: This function runs after the form is closed
	#TODO: Add custom code to clean up and unload snapins when the application exits
	
	$script:ExitCode = 0 #Set the exit code for the Packager
}

#endregion Application Functions

#----------------------------------------------
# Generated Form Function
#----------------------------------------------
function Call-Doc_Cap_pff {

	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System.Core, Version=3.5.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.ServiceProcess, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$formDiscoveryTool = New-Object 'System.Windows.Forms.Form'
	$Script:datagridview1 = New-Object 'System.Windows.Forms.DataGridView'
	$menustrip1 = New-Object 'System.Windows.Forms.MenuStrip'
	$menustrip2 = New-Object 'System.Windows.Forms.MenuStrip'
	$combobox1 = New-Object 'System.Windows.Forms.ComboBox'    
	$captureFirstToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$captureSecondToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$compareToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$DefaultTexttoolstripcombobox1 = New-Object 'System.Windows.Forms.ToolStripComboBox'
	$Script:captureWindowToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$captureFullScreenToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$optionsToolStripMenuItem1 = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$vMControlToolStripMenuItem1 = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$toolstriptextbox1 = New-Object 'System.Windows.Forms.ToolStripTextBox'
	$exportToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$importToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$customiseCaptureTextToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$LoadDocumentationToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$exportfirstToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$exportsecondToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$exportcompareToolStripMenuItem1 = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$exportexclusionListToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$importfirstToolStripMenuItem1 = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$importsecondToolStripMenuItem1 = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$importcompareToolStripMenuItem2 = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$importexclusionListToolStripMenuItem1 = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ARPToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$CMDToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$controlPanelToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$driversToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$eventViewerToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$FTAsToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$MMCToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$processesToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$registryToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$servicesToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$showToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$minimiseToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$maximiseToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$closeToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$backgroundColourToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$colordialog1 = New-Object 'System.Windows.Forms.ColorDialog'
	$toolstripcombobox3 = New-Object 'System.Windows.Forms.ToolStripComboBox'
	$VirtualMachineToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$Script:VMselectiontoolstripcombobox4 = New-Object 'System.Windows.Forms.ToolStripComboBox'
	$usernametoolstriptextbox2 = New-Object 'System.Windows.Forms.ToolStripTextBox'
	$toolstriptextbox3 = New-Object 'System.Windows.Forms.ToolStripTextBox'
	$moveToBottomToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$contextmenustrip1 = New-Object 'System.Windows.Forms.ContextMenuStrip'
	$excludeToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$showAllToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$showExclusionsToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$Script:Selected = New-Object 'System.Windows.Forms.DataGridViewCheckBoxColumn'
	$toolstripcombobox2 = New-Object 'System.Windows.Forms.ToolStripComboBox'
	$timer1 = New-Object 'System.Windows.Forms.Timer'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	$openfiledialog1 = New-Object 'System.Windows.Forms.OpenFileDialog'
    $SaveFileDialog = New-Object 'System.Windows.Forms.SaveFileDialog'
    $OpenFileDialog = New-Object 'System.Windows.Forms.OpenFileDialog'
    $sendToDiscoveryToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	
	
	
	
	
	
	
	$formDiscoveryTool_Load={

if(Test-path "$env:TEMP\CompareExclusion\LastUsed.XML"){
$Script:Exclusions =@()
$Script:Exclusions += Import-Clixml "$env:TEMP\CompareExclusion\LastUsed.XML"
}
Del "$env:TEMP\VMRun\*" -Recurse -ErrorAction SilentlyContinue 

$RunningVMs = Get-RunningVMs
$Script:VMselectiontoolstripcombobox4.Items.Clear()
$RunningVMs|foreach{$Script:VMselectiontoolstripcombobox4.Items.Add($_.path)}
IF("$RunningVMs".Length -gt 0){
$Script:VMselectiontoolstripcombobox4.SelectedIndex = 0
}


Validate
Set-DataGridViewDoubleBuffer -grid $Script:datagridview1 -Enabled $true
		#TODO: Initialize Form Controls here

$Script:DefaultVMUserName = "Packaging"
$Script:DefaultVMPassword = "P4ckag!ng"

$textitems = Get-Content "$CurrentComboTextOptions\DefaultTextItems.txt"
Foreach($item in $textitems){
$DefaultTexttoolstripcombobox1.Items.Add("$item")
}

$DefaultTexttoolstripcombobox1.SelectedIndex = 0
		$textlength = $DefaultTexttoolstripcombobox1.Items| sort-object -property length | select-object -last 1
		$font = New-Object System.Drawing.Font("Calibri", 8.25, [System.Drawing.FontStyle]'Regular')
		$size = [System.Windows.Forms.TextRenderer]::MeasureText($textlength, $font)
		$DefaultTexttoolstripcombobox1.DropDownWidth = $size.Width

$toolstripcombobox2.SelectedIndex = 0
$toolstripcombobox3.SelectedIndex = 0
		$textlength = $toolstripcombobox2.Items| sort-object -property length | select-object -last 1
		$font = New-Object System.Drawing.Font("Calibri", 8.25, [System.Drawing.FontStyle]'Regular')
		$size = [System.Windows.Forms.TextRenderer]::MeasureText($textlength, $font)
		$toolstripcombobox2.DropDownWidth = $size.Width


		#TODO: Initialize Form Controls herev
		$textlength = $Script:VMselectiontoolstripcombobox4.Items| sort-object -property length | select-object -last 1
        if($textlength -gt 0){
		$Script:VMselectiontoolstripcombobox4.SelectedIndex = 0
		$font = New-Object System.Drawing.Font("Calibri", 8.25, [System.Drawing.FontStyle]'Regular')
		$size = [System.Windows.Forms.TextRenderer]::MeasureText($textlength, $font)
		$Script:VMselectiontoolstripcombobox4.DropDownWidth = $size.Width
        }
		$script:Password = "P4ckag!ng"
		$toolstriptextbox3.Text = "*********"
		$script:PasswordLength = $script:Password.Length
		$toolstriptextbox3.ShortcutsEnabled = $false
        #$Script:fixedheight = 88
		$Script:fixedheight = $formDiscoveryTool.Height
		$formDiscoveryTool.MinimumSize.Height = $script:fixedheight
        $Script:captureWindowToolStripMenuItem.Text = "Open Discovery"
        $captureFullScreenToolStripMenuItem.Visible = $false
        $timer1.Enabled = $true

}
	
	
	$showToolStripMenuItem_Click={
		#TODO: Place custom script here
		if($Script:datagridview1.Visible -eq $false){
		$Script:datagridview1.Visible = $True
		$showToolStripMenuItem.Text = "Hide"
		if(($formDiscoveryTool.WindowState -eq 'Maximized') -or ($formDiscoveryTool.Height -Gt "450") ){}else{$formDiscoveryTool.Height = "450" }
			}else{
		$Script:datagridview1.Visible = $False
		$showToolStripMenuItem.Text = "Show"
		if($formDiscoveryTool.WindowState -eq 'Maximized'){}else{$formDiscoveryTool.Height = $Script:fixedheight }	
			}
	}
	
	
	
	$formDiscoveryTool_Resize={
		#TODO: Place custom script here
	if($Script:datagridview1.Visible -eq $false){$formDiscoveryTool.Height = $Script:fixedheight
}
	}
	
	
	
	$minimiseToolStripMenuItem_Click={
		#TODO: Place custom script here
		$formDiscoveryTool.WindowState = 'Minimized'
	}
	
	$maximiseToolStripMenuItem_Click={
		#TODO: Place custom script here
		if($formDiscoveryTool.WindowState -eq 'Normal' ){
		$formDiscoveryTool.FormBorderStyle = 'None'
		$formDiscoveryTool.WindowState = 'Maximized'
		$formDiscoveryTool.TopMost = $true
		$maximiseToolStripMenuItem.Text = "Restore"
		$closeToolStripMenuItem.Visible = $true
		$moveToBottomToolStripMenuItem.Visible = $true
		}elseif($formDiscoveryTool.WindowState -eq 'Maximized'){
		$formDiscoveryTool.WindowState = 'Minimized'
		$formDiscoveryTool.TopMost = $false
		$closeToolStripMenuItem.Visible = $False
		$moveToBottomToolStripMenuItem.Visible = $False
		$formDiscoveryTool.WindowState = 'Normal'
		$maximiseToolStripMenuItem.Text = "Maximise"
		$formDiscoveryTool.FormBorderStyle = 'Sizable'
		$menustrip1.Dock = 'Top'
		$menustrip2.Dock = 'Top'
		$moveToBottomToolStripMenuItem.Text = "Move To Bottom"	
			if($Script:datagridview1.Visible -eq $false){
				$formDiscoveryTool.Height = $Script:fixedheight
				}else{
				if($formDiscoveryTool.Height -lt "450"){
				$formDiscoveryTool.Height = "450"
				}
			}
	
			}	
			
	}
	
	$closeToolStripMenuItem_Click={
		#TODO: Place custom script here
		$formDiscoveryTool.Close()
	}
	
	
	
	$toolstripcombobox3_SelectedIndexChanged={
Add-Exclusions
Validate

IF($toolstripcombobox3.SelectedItem -eq "Exclusions"){
$showAllToolStripMenuItem.Text = "Show All"
$showExclusionsToolStripMenuItem.Text = "Show Filtered"
}elseIF($toolstripcombobox3.SelectedItem -eq "All"){
$showAllToolStripMenuItem.Text = "Show Exclusions"
$showExclusionsToolStripMenuItem.Text = "Show Filtered"
}elseif($toolstripcombobox3.SelectedItem -eq "Filtered"){
$showAllToolStripMenuItem.Text = "Show Exclusions"
$showExclusionsToolStripMenuItem.Text = "Show All"
}
	    $combobox1.Focus()

		#$menustrip1.BackColor = 
		#TODO: Place custom script here

	}
	
	$backgroundColourToolStripMenuItem_Click={
		#TODO: Place custom script here
		$colordialog1.Color = 'Azure'
		$colordialog1.ShowDialog()
		if($colordialog1.Color -ne $null){
		$menustrip1.BackColor = $colordialog1.Color
		$menustrip2.BackColor = $colordialog1.Color
		}
	}
	
	$moveToBottomToolStripMenuItem_Click={
		#TODO: Place custom script here
		if($moveToBottomToolStripMenuItem.Text -eq "Move To Bottom"){
		$menustrip1.Dock = 'Bottom'
		$menustrip2.Dock = 'Bottom'
		$moveToBottomToolStripMenuItem.Text = "Move To Top"	
			}else{
		$menustrip1.Dock = 'Top'
		$menustrip2.Dock = 'Top'
		$moveToBottomToolStripMenuItem.Text = "Move To Bottom"	
			}
	}
	
	
	$toolstriptextbox3_TextChanged={
		#TODO: Place custom script here
	#Event Argument: $_ = [System.ComponentModel.CancelEventArgs]
		#TODO: Place custom script here
		if($toolstriptextbox3.Text.Length -eq 0){$script:Password = ""}
		elseif($toolstriptextbox3.Text.Length -gt "$script:Password".length ){
		$script:Password = $script:Password + $toolstriptextbox3.Text.Substring($toolstriptextbox3.Text.Length - 1)
		$toolstriptextbox3.Text = "*" * "$script:Password".length
		$toolstriptextbox3.SelectionStart = $toolstriptextbox3.Text.Length
			}elseif($toolstriptextbox3.Text.Length -lt "$script:Password".length ){
			
		$script:Password = $script:Password.substring(0,"$script:Password".length -("$script:Password".length - $toolstriptextbox3.Text.Length ))
		$toolstriptextbox3.Text = "*" * "$script:Password".length
		$toolstriptextbox3.SelectionStart = $toolstriptextbox3.Text.Length
			}
        $Script:DefaultVMPassword = $script:Password
	}
	
	
	
	
	
	$exportfirstToolStripMenuItem_Click={

    $SaveFileDialog.ShowDialog() | Out-Null
    IF("$($SaveFileDialog.filename)" -ne $null){
    if(test-path "$(Split-Path "$($SaveFileDialog.filename)" -ErrorAction SilentlyContinue)"){
    $Script:FirstCapture |Export-Clixml "$($SaveFileDialog.filename)"
    }
    }
		#TODO: Place custom script here
		
	}
	
	$exportsecondToolStripMenuItem_Click={

    $SaveFileDialog.ShowDialog() | Out-Null
    IF("$($OpenFileDialog.filename)" -ne $null){
    if(test-path "$(Split-Path "$($SaveFileDialog.filename)" -ErrorAction SilentlyContinue)"){
    $Script:SecondCapture |Export-Clixml "$($SaveFileDialog.filename)"
    }
    }
		#TODO: Place custom script here
		
	}
	
	$exportcompareToolStripMenuItem1_Click={

    $SaveFileDialog.ShowDialog() | Out-Null
    IF("$($SaveFileDialog.filename)" -ne $null){
    if(test-path "$(Split-Path "$($SaveFileDialog.filename)" -ErrorAction SilentlyContinue)"){
    $Script:Compare,$script:SecondCapture,$script:FirstCapture |Export-Clixml "$($SaveFileDialog.filename)"
    }
}
		#TODO: Place custom script here
		
	}
	
	$exportexclusionListToolStripMenuItem_Click={

    $SaveFileDialog.ShowDialog() | Out-Null
    IF("$($SaveFileDialog.filename)" -ne $null){
    if(test-path "$(Split-Path "$($SaveFileDialog.filename)")"){
    $Script:Exclusions |Export-Clixml "$($SaveFileDialog.filename)"
}
}
		#TODO: Place custom script here
		
	}
	
	$importfirstToolStripMenuItem1_Click={

    $OpenFileDialog.ShowDialog() | Out-Null
    IF("$($OpenFileDialog.filename)" -ne $null){
    if(test-path "$($OpenFileDialog.filename)"){
    $Script:FirstCapture = Import-Clixml "$($OpenFileDialog.filename)"
}
Validate
}
		#TODO: Place custom script here
		
	}
	
	$importsecondToolStripMenuItem1_Click={

    $OpenFileDialog.ShowDialog() | Out-Null
    IF("$($OpenFileDialog.filename)" -ne $null){
    if(test-path "$($OpenFileDialog.filename)"){
    $Script:SecondCapture = Import-Clixml "$($OpenFileDialog.filename)"
}
Validate
}
		#
		#TODO: Place custom script here
		
	}
	
	$importcompareToolStripMenuItem2_Click={

    $OpenFileDialog.ShowDialog() | Out-Null
    IF("$($OpenFileDialog.filename)" -ne $null){
    if(test-path "$($OpenFileDialog.filename)"){
    $Script:Compare = (Import-Clixml "$($OpenFileDialog.filename)")[0]
    $script:SecondCapture = (Import-Clixml "$($OpenFileDialog.filename)")[1]
    $script:FirstCapture = (Import-Clixml "$($OpenFileDialog.filename)")[2]
}
Add-Exclusions
Validate
}
		#TODO: Place custom script here
		
	}
	
	$importexclusionListToolStripMenuItem1_Click={

    $OpenFileDialog.ShowDialog() | Out-Null
    IF("$($OpenFileDialog.filename)" -ne $null){
    if(test-path "$($OpenFileDialog.filename)" -ErrorAction SilentlyContinue){
    $Script:Exclusions = Import-Clixml "$($OpenFileDialog.filename)"
}
Validate
}
		#TODO: Place custom script here
		
	}
	
	$customiseCaptureTextToolStripMenuItem_Click={
    Invoke-item "$CurrentComboTextOptions\DefaultTextItems.txt"
    $return = Invoke-MessageBox -FormHeader "Done?" -Message "Have you finished and saved your modifications to the default text? `n Click Yes when you have." -ButtonOne "Yes"
    if($return -eq "Yes"){
    $DefaultTexttoolstripcombobox1.Items.Clear()
    foreach($item in (Get-content "$CurrentComboTextOptions\DefaultTextItems.txt")){
    $DefaultTexttoolstripcombobox1.Items.Add("$item")
    }
    $DefaultTexttoolstripcombobox1.SelectedIndex = 0
    }
		#TODO: Place custom script here
		#TODO: Place custom script here
		
	}
	
	$ARPToolStripMenuItem_Click={
		#TODO: Place custom script here
	ValidationCredentials
If($VMselectiontoolstripcombobox4.selecteditem -ne ""){
IF("$(Get-RunningVMs|where-Object -Property "path" -eq $VMselectiontoolstripcombobox4.selecteditem )".Length -gt 0 ){
Invoke-VMProgramExe -EXE "`"CMD.exe`" /C Appwiz.cpl" -VMXpath "$($VMselectiontoolstripcombobox4.selecteditem)"
}
}	
	}
	
	$CMDToolStripMenuItem_Click={

ValidationCredentials
If($VMselectiontoolstripcombobox4.selecteditem -ne ""){
IF("$(Get-RunningVMs|where-Object -Property "path" -eq $VMselectiontoolstripcombobox4.selecteditem )".Length -gt 0 ){
Invoke-VMProgramExe -EXE "CMD.exe" -VMXpath "$($VMselectiontoolstripcombobox4.selecteditem)"
}
}
		#TODO: Place custom script here
		
	}
	
	$controlPanelToolStripMenuItem_Click={

ValidationCredentials
If($VMselectiontoolstripcombobox4.selecteditem -ne ""){
IF("$(Get-RunningVMs|where-Object -Property "path" -eq $VMselectiontoolstripcombobox4.selecteditem )".Length -gt 0 ){
Invoke-VMProgramExe -EXE "`"Control.exe`"" -VMXpath "$($VMselectiontoolstripcombobox4.selecteditem)"
}
}
		#TODO: Place custom script here
		
	}
	
	$driversToolStripMenuItem_Click={


ValidationCredentials
If($VMselectiontoolstripcombobox4.selecteditem -ne ""){
IF("$(Get-RunningVMs|where-Object -Property "path" -eq $VMselectiontoolstripcombobox4.selecteditem )".Length -gt 0 ){
Invoke-VMProgramExe -EXE "`"C:\Windows\Explorer.exe`" `"C:\Windows\System32\DriverStore\FileRepository`"" -VMXpath "$($VMselectiontoolstripcombobox4.selecteditem)"
}
}
		#TODO: Place custom script here
		
	}
	
	$eventViewerToolStripMenuItem_Click={

ValidationCredentials
If($VMselectiontoolstripcombobox4.selecteditem -ne ""){
IF("$(Get-RunningVMs|where-Object -Property "path" -eq $VMselectiontoolstripcombobox4.selecteditem )".Length -gt 0 ){
Invoke-VMProgramExe -EXE "`"C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe`" /C `"eventvwr.msc`"" -VMXpath "$($VMselectiontoolstripcombobox4.selecteditem)"
}
}

		#TODO: Place custom script here
		
	}
	
	$FTAsToolStripMenuItem_Click={

ValidationCredentials
If($VMselectiontoolstripcombobox4.selecteditem -ne ""){
IF("$(Get-RunningVMs|where-Object -Property "path" -eq $VMselectiontoolstripcombobox4.selecteditem )".Length -gt 0 ){
Invoke-VMProgramExe -EXE "`"C:\Windows\system32\control.exe`" /name Microsoft.DefaultPrograms /page pageFileAssoc" -VMXpath "$($VMselectiontoolstripcombobox4.selecteditem)"
}
}
		#TODO: Place custom script here
		
	}
	
	$MMCToolStripMenuItem_Click={

ValidationCredentials
If($VMselectiontoolstripcombobox4.selecteditem -ne ""){
IF("$(Get-RunningVMs|where-Object -Property "path" -eq $VMselectiontoolstripcombobox4.selecteditem )".Length -gt 0 ){
Invoke-VMProgramExe -EXE "`"C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe`" /C `"MMC.exe`"" -VMXpath "$($VMselectiontoolstripcombobox4.selecteditem)"
}
}
		#TODO: Place custom script here
		
	}
	
	$processesToolStripMenuItem_Click={

ValidationCredentials
If($VMselectiontoolstripcombobox4.selecteditem -ne ""){
IF("$(Get-RunningVMs|where-Object -Property "path" -eq $VMselectiontoolstripcombobox4.selecteditem )".Length -gt 0 ){
Invoke-VMProgramExe -EXE "`"C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe`" /C `"taskmgr.exe`"" -VMXpath "$($VMselectiontoolstripcombobox4.selecteditem)"
}
}
		#TODO: Place custom script here
		
	}
	
	$registryToolStripMenuItem_Click={
		#TODO: Place custom script here
		ValidationCredentials
If($VMselectiontoolstripcombobox4.selecteditem -ne ""){
IF("$(Get-RunningVMs|where-Object -Property "path" -eq $VMselectiontoolstripcombobox4.selecteditem )".Length -gt 0 ){
Invoke-VMProgramExe -EXE "`"C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe`" /C `"C:\Windows\Regedit.exe`"" -VMXpath "$($VMselectiontoolstripcombobox4.selecteditem)"
}
}
	}
	
	$servicesToolStripMenuItem_Click={

ValidationCredentials
If($VMselectiontoolstripcombobox4.selecteditem -ne ""){
IF("$(Get-RunningVMs|where-Object -Property "path" -eq $VMselectiontoolstripcombobox4.selecteditem )".Length -gt 0 ){
Invoke-VMProgramExe -EXE "`"C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe`" /C Services.msc" -VMXpath "$($VMselectiontoolstripcombobox4.selecteditem)"
}
}
		#TODO: Place custom script here
		
	}
	
	$VirtualMachineToolStripMenuItem_Click={
		#TODO: Place custom script here
$RunningVMs = Get-RunningVMs
$Script:VMselectiontoolstripcombobox4.Items.Clear()
$RunningVMs|foreach{$Script:VMselectiontoolstripcombobox4.Items.Add($_.path)}
IF("$RunningVMs".Length -gt 0){
$textlength = $Script:VMselectiontoolstripcombobox4.Items| sort-object -property length | select-object -last 1
    if($textlength -gt 0){
	$Script:VMselectiontoolstripcombobox4.SelectedIndex = 0
	$font = New-Object System.Drawing.Font("Calibri", 8.25, [System.Drawing.FontStyle]'Regular')
	$size = [System.Windows.Forms.TextRenderer]::MeasureText($textlength, $font)
	$Script:VMselectiontoolstripcombobox4.DropDownWidth = $size.Width
    }
}

		
	}
	
	
	$captureFirstToolStripMenuItem_Click={

ValidationCredentials

If($Script:VMselectiontoolstripcombobox4.text -ne ""){
IF((Get-RunningVMs).Path -eq "$($Script:VMselectiontoolstripcombobox4.text)"){
$toolstriptextbox1.Text = "Starting First Capture`n This may take a minute or two..."
#[System.Windows.Forms.Application]::DoEvents()
Try{
Get-Difference -Capture First -VMXpath "$($Script:VMselectiontoolstripcombobox4.Text)"
$toolstriptextbox1.Text = "First Capture Done!`nInstall your application then click Capture Second."
}Catch{
$toolstriptextbox1.Text = "Nothing captured, is Z drive is mapped?`nUsername and Password is correct?"
}
#[System.Windows.Forms.Application]::DoEvents()
}
}

Validate
		#TODO: Place custom script here
		
	}
	
	$captureSecondToolStripMenuItem_Click={

ValidationCredentials

If($Script:VMselectiontoolstripcombobox4.text -ne ""){
IF((Get-RunningVMs).Path -eq "$($Script:VMselectiontoolstripcombobox4.text)"){
$toolstriptextbox1.Text = "Starting Second Capture`n This may take a minute or two..."
#[System.Windows.Forms.Application]::DoEvents()
Try{
Get-Difference -Capture Second -VMXpath "$($Script:VMselectiontoolstripcombobox4.Text)"
$toolstriptextbox1.Text = "Second Capture Done!`nClick Compare to get results."
}Catch{
$toolstriptextbox1.Text = "Nothing captured, Is Z drive is mapped?`nUsername and Password is correct?"
}
#[System.Windows.Forms.Application]::DoEvents()
}
}

Validate
		#TODO: Place custom script here
		
	}
	
	$compareToolStripMenuItem_Click={

$toolstriptextbox1.Text = "Starting Compare`n This may take a minute."
#[System.Windows.Forms.Application]::DoEvents()
Get-Difference -Capture Compare
$toolstriptextbox1.Text = "Compare Done!`nYou may wish to exclude some values."
#[System.Windows.Forms.Application]::DoEvents()
Add-Exclusions


	if($Script:datagridview1.Visible -eq $false){
	$Script:datagridview1.Visible = $True
	$showToolStripMenuItem.Text = "Hide"
	if(($formDiscoveryTool.WindowState -eq 'Maximized') -or ($formDiscoveryTool.Height -Gt "450") ){}else{$formDiscoveryTool.Height = "450" }
	}
		#TODO: Place custom script here
Validate	
	}
	
	$DefaultTexttoolstripcombobox1_SelectedIndexChanged={
		#TODO: Place custom script here
	    $combobox1.Focus()
       
	}

    $toolstripcombobox2_SelectedIndexChanged={
	#TODO: Place custom script here
	$combobox1.Focus()
    }

    $Script:VMselectiontoolstripcombobox4_SelectedIndexChanged={
	#TODO: Place custom script here
	$combobox1.Focus()	
    }
	
	$Script:captureWindowToolStripMenuItem_Click={

if($Script:captureWindowToolStripMenuItem.text -eq "Open Discovery"){

		if($openfiledialog1.ShowDialog() -eq 'OK')
		{

if((Test-Path "$($openfiledialog1.FileName)") -and ("$($openfiledialog1.FileName)" -match ".docx")){
$Script:DiscoLocation = $openfiledialog1.FileName
$script:Word = New-Object -Com Word.Application
$Script:Word.Visible = $True #set this to true for debugging
	# http://msdn.microsoft.com/en-us/library/bb238158%28v=office.12%29.aspx

	#$Script:Doc = $Script:Word.Documents.Open("$($textboxFile.text)")

$Script:ISobjRange = $null
$Script:FLobjRange = $null
$Script:AMobjRange = $null
$Script:HMobjRange = $null
$Script:CPEobjRange = $null
$Script:CMobjRange = $null
$Script:FTobjRange = $null

$Script:FTASobjRange = $null
$Script:AdditionalobjRange = $null
$Script:FWRSobjRange = $null
$Script:INSTALLDIRobjRange = $null
$Script:COMMONDIRobjRange = $null
$Script:OTHERDIRobjRange = $null

$Script:Doc = $Script:Word.Documents.Open("$Script:DiscoLocation")
$Script:Doc.Activate()

$Script:captureWindowToolStripMenuItem.Text = "Capture Windowed"
$captureFullScreenToolStripMenuItem.Visible = "True"
}


}



}else{
PassCapToDoc -Type Windowed
}

		#TODO: Place custom script here
		
	}

$timer1_Tick={

if($Script:Doc -ne $null){
if($script:Word.ActiveDocument.count -eq 0){
$captureFullScreenToolStripMenuItem.Visible = $False
$Script:captureWindowToolStripMenuItem.Text = "Open Discovery"
}else{
if($Script:Compare -ne $null){
$sendToDiscoveryToolStripMenuItem.Visible = $true
}else{
$sendToDiscoveryToolStripMenuItem.Visible = $false
}

}

}else{
$sendToDiscoveryToolStripMenuItem.Visible = $false
}

	#TODO: Place custom script here
	
}
	
	$captureFullScreenToolStripMenuItem_Click={
PassCapToDoc -Type FullScreen
		#TODO: Place custom script here
		
	}
	
	$usernametoolstriptextbox2_TextChanged={
		#TODO: Place custom script here
$Script:DefaultVMUserName = $usernametoolstriptextbox2.Text
		
	}

	$excludeToolStripMenuItem_Click={

$Selectedrow =  $Script:datagridview1.SelectedCells.RowIndex|Select -First 1
if($Selectedrow -ne $null){

$dataRow = $Script:datagridview1.Rows[$Selectedrow].DataBoundItem

if(($dataRow|Select -ExpandProperty "Name") -match "Exclusion:"){
if("$(($dataRow|get-member).Name)" -match "Exclusion"){
$Script:Exclusions1 =@()
foreach($Exclusion in $Script:Exclusions ){
if($Exclusion.Name -eq ("$($dataRow|Select -ExpandProperty "Name")").TrimStart("Exclusion:")){
if($Exclusion.Exclusion -ne "$($dataRow|Select -ExpandProperty "Exclusion")" ){
$Script:Exclusions1 += $Exclusion
}
}else{
$Script:Exclusions1 += $Exclusion
}
}
$Script:Exclusions = $Script:Exclusions1

}else{

$Script:Exclusions1 =@()
foreach($Exclusion in $Script:Exclusions ){
if($Exclusion.Name -eq ("$($dataRow|Select -ExpandProperty "Name")").TrimStart("Exclusion:")){
if($Exclusion.Exclusion -ne "$($dataRow|Select -ExpandProperty "Difference")" ){
$Script:Exclusions1 += $Exclusion
}
}else{
$Script:Exclusions1 += $Exclusion
}
}
$Script:Exclusions = $Script:Exclusions1
}}else{

$Script:Exclusions += New-Object PsObject -Property @{Name = "$($dataRow|Select -ExpandProperty "Name")"
Exclusion = "$($dataRow|Select -ExpandProperty "Difference")"    }


}
		#TODO: Place custom script here
Add-Exclusions
Validate
}	
	}
	
	$showAllToolStripMenuItem_Click={
IF($showAllToolStripMenuItem.Text -eq "Show Exclusions"){
$toolstripcombobox3.SelectedItem = "Exclusions"
$showAllToolStripMenuItem.Text = "Show All"
$showExclusionsToolStripMenuItem.Text = "Show Filtered"
}elseIF($showAllToolStripMenuItem.Text -eq "Show All"){
$toolstripcombobox3.SelectedItem = "All"
$showAllToolStripMenuItem.Text = "Show Exclusions"
$showExclusionsToolStripMenuItem.Text = "Show Filtered"
}elseif($showAllToolStripMenuItem.Text -eq "Show Filtered"){
$toolstripcombobox3.SelectedItem = "Filtered"
$showAllToolStripMenuItem.Text = "Show Exclusions"
$showExclusionsToolStripMenuItem.Text = "Show All"
}
	
	}
	
	$showExclusionsToolStripMenuItem_Click={
		#TODO: Place custom script here
IF($showExclusionsToolStripMenuItem.Text -eq "Show Exclusions"){
$toolstripcombobox3.SelectedItem = "Exclusions"
$showAllToolStripMenuItem.Text = "Show All"
$showExclusionsToolStripMenuItem.Text = "Show Filtered"
}elseIF($showExclusionsToolStripMenuItem.Text -eq "Show All"){
$toolstripcombobox3.SelectedItem = "All"
$showAllToolStripMenuItem.Text = "Show Exclusions"
$showExclusionsToolStripMenuItem.Text = "Show Filtered"
}elseif($showExclusionsToolStripMenuItem.Text -eq "Show Filtered"){
$toolstripcombobox3.SelectedItem = "Filtered"
$showAllToolStripMenuItem.Text = "Show Exclusions"
$showExclusionsToolStripMenuItem.Text = "Show All"
}
	}

	$sendToDiscoveryToolStripMenuItem_Click={

$ping = PingTest 
if($ping -eq "Online"){
$Active =  (ActiveApps -runningas "$RunningAs" |Where-Object { ($_.ApplicationStatus -eq "Discovery") -or ($_.ApplicationStatus -eq "Packaging") -or ($_.ApplicationStatus -eq "Quality assurance")}).count
}
if("$Active".Length -gt 1 -or ($ping -eq "Offline") -or ($script:OfflineMode -eq $true)){

if($toolstripcombobox3.SelectedItem -ne "Filtered"){
$toolstripcombobox3.SelectedItem = "Filtered"
}

$arrayExport =@()
$i = 0
$rowcount = $Script:datagridview1.RowCount
While($i -lt $rowcount){
IF($Script:datagridview1.Rows[$i].Cells["Selected"].Value -eq $true ){
$arrayExport += $Script:datagridview1.Rows[$i]| select -expand DataBoundItem
}
$i++
}

if(($Script:Doc.ActiveWindow -ne $null) -and ($arrayExport.count -gt 0)){


if($Script:FTASobjRange.text -eq $null){
$Script:FTASobjRange = $Script:Doc.BookMarks.Item("FTAS_DISCOTOOL").Range
$Script:Doc.BookMarks.Item("FTAS_DISCOTOOL").select()
$Script:SelectedExports = $arrayExport|Where-Object -Property Name -eq "FTAs"
IF($Script:SelectedExports.name.count -gt 0){
$Script:FTASobjRange.text = "$($Script:SelectedExports.Difference|foreach{"$_
"})"
}else{$Script:FTASobjRange.text = "N/A"}
}else{
$Script:SelectedExports = $arrayExport|Where-Object -Property Name -eq "FTAs"
IF($Script:SelectedExports.name.count -gt 0){
$Script:FTASobjRange.text = "$($Script:SelectedExports.Difference|foreach{"$_
"})"
}else{$Script:FTASobjRange.text = "N/A"}
}



if($Script:FWRSobjRange.text -eq $null){
$Script:FWRSobjRange = $Script:Doc.BookMarks.Item("FIREWALL_DISCOTOOL").Range
$Script:Doc.BookMarks.Item("FIREWALL_DISCOTOOL").Select()
$Script:SelectedExports = $arrayExport|Where-Object -Property Name -eq "FireWallRules"
IF($Script:SelectedExports.name.count -gt 0){
$Script:FWRSobjRange.text = "$($Script:SelectedExports.Difference|foreach{"$_
"})"
$Script:FWRSobjRange.Font.Color = "wdColorBlack"
}else{$Script:FWRSobjRange.text = "N/A"
$Script:FWRSobjRange.Font.Color = "wdColorBlack"}
}else{
$Script:SelectedExports = $arrayExport|Where-Object -Property Name -eq "FireWallRules"
IF($Script:SelectedExports.name.count -gt 0){
$Script:FWRSobjRange.text = "$($Script:SelectedExports.Difference|foreach{"$_
"})"
$Script:FWRSobjRange.Font.Color = "wdColorBlack"
}else{$Script:FWRSobjRange.text = "N/A"
$Script:FWRSobjRange.Font.Color = "wdColorBlack"}
}

#$Script:SelectedExports += New-Object Psobject -Property @{
#Name = "Drivers"
#Difference = "Drivers1"
#}


$Script:SelectedExports=@()
$Script:SelectedExports += $arrayExport|Where-Object -Property Name -eq "ODBCs"
$Script:SelectedExports += $arrayExport|Where-Object -Property Name -eq "Environmentals"
$Script:SelectedExports += $arrayExport|Where-Object -Property Name -eq "Drivers"

If($Script:SelectedExports.name.count -gt 0){

	$AdditionalInformation = $Doc.SelectContentControlsByTag("AdditionalInformation")
	$r=$AdditionalInformation | select -expand range
	$r.Text = "YES - See below for full details"


		
$fields = $Doc.FormFields
if(($Script:SelectedExports|where -property "Name" -eq "Environmentals").Difference.count -gt 0){
$fields.Item("ENV_YES").CheckBox.Value = $true
$Textone = "Application contains the following Environmental(s):
$(($Script:SelectedExports|where -property "Name" -eq "Environmentals").Difference|foreach{"'$_'`n"})
"
$Script:ENVobjRange = $Script:Doc.BookMarks.Item("ENV_DISCOTOOL").Range
$Script:Doc.BookMarks.Item("ENV_DISCOTOOL").select()
$Script:ENVobjRange.text = $Textone
}else{$fields.Item("ENV_NO").CheckBox.Value = $true}


if(($Script:SelectedExports|where -property "Name" -eq "Drivers").Difference.count -gt 0){
$fields.Item("DRIVERS_YES").CheckBox.Value = $true
$Textone = "Application contains the following Driver(s):
$(($Script:SelectedExports|where -property "Name" -eq "Drivers").Difference|foreach{"'$_'`n"})
"
$Script:DRIVERSobjRange = $Script:Doc.BookMarks.Item("DRIVERS_DISCOTOOL").Range
$Script:DRIVERSobjRange.text = $Textone
}else{$fields.Item("DRIVERS_NO").CheckBox.Value = $true}


if(($Script:SelectedExports|where -property "Name" -eq "ODBCs").Difference.count -gt 0){
$fields.Item("ODBC_YES").CheckBox.Value = $true
$Textone = "Application contains the following ODBC(s):
$(($Script:SelectedExports|where -property "Name" -eq "ODBCs").Difference|foreach{"'$_'`n"})
"
$Script:ODBCSobjRange = $Script:Doc.BookMarks.Item("ODBC_DISCOTOOL").Range
$Script:ODBCSobjRange.text = $Textone
}else{$fields.Item("ODBC_NO").CheckBox.Value = $true}

}else{
	$AdditionalInformation = $Doc.SelectContentControlsByTag("AdditionalInformation")
	$r=$AdditionalInformation | select -expand range
	$r.Text = "No"
}


if($Script:INSTALLDIRobjRange.text -eq $null){

$Script:INSTALLDIRobjRange = $Script:Doc.BookMarks.Item("INSTALLDIR_DISCOTOOL").Range
$Script:COMMONDIRobjRange = $Script:Doc.BookMarks.Item("COMMONFILES_DISCOTOOL").Range
$Script:OTHERDIRobjRange = $Script:Doc.BookMarks.Item("OTHERLOCATIONS_DISCOTOOL").Range
$Script:SelectedExports = $arrayExport|Where-Object -Property Name -eq "Files"

$installdirs =@()
$CommonDir =@()
$otherDir =@()

foreach($item in $Script:SelectedExports){
if($item |where -Property "Difference" -match "%Common"){
$CommonDir += $item
}
elseIF($item |where -Property "Difference" -match "%ProgramFiles"){
$installdirs += $item
}
else{$otherDir += $item}
}

IF($installdirs.name.count -gt 0){
$Script:INSTALLDIRobjRange.text = "$($installdirs.Difference|foreach{"$_
"})"
}else{$Script:INSTALLDIRobjRange.text = "N/A"}

IF($CommonDir.name.count -gt 0){
$Script:COMMONDIRobjRange.text = "$($CommonDir.Difference|foreach{"$_
"})"
}else{$Script:COMMONDIRobjRange.text = "N/A"}

IF($otherDir.name.count -gt 0){
$Script:OTHERDIRobjRange.text = "$($otherDir.Difference|foreach{"$_
"})"
}else{$Script:OTHERDIRobjRange.text = "N/A"}


##adding shortcuts
$numberofTables = $Script:Doc.Tables.count
$Script:SelectedExports = $arrayExport|Where-Object -Property Name -eq "Shortcuts"
$shortcutstoadd =@()
foreach($item in $Script:SelectedExports|where -Property "Difference" -match ".lnk"){
$shortcutname = ([IO.FileInfo[]]$item.Difference).BaseName
$shortcutstoadd += $Script:SecondCapture[10]|Where-Object -Property "ShortcutName" -eq $shortcutname
}

$tabletitle = $null
while($tabletitle -ne "Shortcuts" -and ($numberofTables -gt 1)){
 $tabletitle = $Script:Doc.Tables.Item($numberofTables).title
 $shortcuttable = $Script:Doc.Tables.Item($numberofTables)
 $numberofTables--
 }
 $shortcuttable.select()
 $shortcuttable.range.text
 while(($shortcuttable.Rows.count -1) -lt $Script:SelectedExports.Difference.count){
 $shortcuttable.Rows.add()
 }
 $row = 2
 if($shortcutstoadd.Shortcutname.count -gt 0){
 foreach($shortcut in $shortcutstoadd){
 $shortcuttable.cell($row,1).range.text = $shortcut.ShortcutName
 $shortcuttable.cell($row,2).range.text = $shortcut.Target
 $shortcuttable.cell($row,3).range.text = $shortcuT.Parameters
 $shortcuttable.cell($row,4).range.text = $shortcut.StartsIn
 $shortcuttable.cell($row,5).range.text = $shortcut.IconFileName
 $shortcuttable.cell($row,6).range.text = $shortcut.LocationInStartMenu
 $row++
 }
 }else{
 $shortcuttable.cell($row,1).range.text = "No Shortcuts"
 }
 ##adding shortcuts




}
}


}else{Invoke-MessageBox -FormHeader "No Active Applications" -Message "You can't run this with no applications in an active status.`n Sorry guys" -ButtonOne "Damn it! Okay..." }	
#TODO: Place custom script here
		
	}

	$LoadDocumentationToolStripMenuItem_Click={
		$ie = New-Object -com InternetExplorer.Application 
        $ie.navigate2("https://uam.ms.myatos.net/Config%20%20Automation%20Templates/Discovery%20Assistant/Documentation.pdf")

}

$formDiscoveryTool_KeyDown=[System.Windows.Forms.KeyEventHandler]{
#Event Argument: $_ = [System.Windows.Forms.KeyEventArgs]
	#TODO: Place custom script here
	if ($_.Control -and $_.KeyCode -eq 'O') { 
$messageBoxOffline = Invoke-MessageBox -FormHeader "Go Offline" -Message "Would you like to use the tool in offlinemode? `nThis will bypass active status validation." -ButtonOne "Yes" -ButtonTwo "No"
if($messageBoxOffline -eq "Yes"){
$script:OfflineMode = $true

}

}
	
}
		
	
	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$formDiscoveryTool.WindowState = $InitialFormWindowState
	}
	
	$Form_Cleanup_FormClosed=
	{

    MD "$env:TEMP\CompareExclusion\" -ErrorAction SilentlyContinue
    $Script:Exclusions| Export-Clixml "$env:TEMP\CompareExclusion\LastUsed.XML"
    Del "$env:TEMP\VMRun\*" -Recurse -ErrorAction SilentlyContinue 
		#Remove all event handlers from the controls
		try
		{
			$formDiscoveryTool.remove_Load($formDiscoveryTool_Load)
			$formDiscoveryTool.remove_Resize($formDiscoveryTool_Resize)
            $formDiscoveryTool.remove_KeyDown($formDiscoveryTool_KeyDown)
			$captureFirstToolStripMenuItem.remove_Click($captureFirstToolStripMenuItem_Click)
			$captureSecondToolStripMenuItem.remove_Click($captureSecondToolStripMenuItem_Click)
			$compareToolStripMenuItem.remove_Click($compareToolStripMenuItem_Click)
			$DefaultTexttoolstripcombobox1.remove_SelectedIndexChanged($DefaultTexttoolstripcombobox1_SelectedIndexChanged)
            $Script:VMselectiontoolstripcombobox4.remove_SelectedIndexChanged($Script:VMselectiontoolstripcombobox4_SelectedIndexChanged)
            $toolstripcombobox2.remove_SelectedIndexChanged($toolstripcombobox2_SelectedIndexChanged)
			$Script:captureWindowToolStripMenuItem.remove_Click($Script:captureWindowToolStripMenuItem_Click)
			$captureFullScreenToolStripMenuItem.remove_Click($captureFullScreenToolStripMenuItem_Click)
			$customiseCaptureTextToolStripMenuItem.remove_Click($customiseCaptureTextToolStripMenuItem_Click)
			$LoadDocumentationToolStripMenuItem.remove_Click($LoadDocumentationToolStripMenuItem_Click)
			$exportfirstToolStripMenuItem.remove_Click($exportfirstToolStripMenuItem_Click)
			$exportsecondToolStripMenuItem.remove_Click($exportsecondToolStripMenuItem_Click)
			$exportcompareToolStripMenuItem1.remove_Click($exportcompareToolStripMenuItem1_Click)
			$exportexclusionListToolStripMenuItem.remove_Click($exportexclusionListToolStripMenuItem_Click)
			$importfirstToolStripMenuItem1.remove_Click($importfirstToolStripMenuItem1_Click)
			$importsecondToolStripMenuItem1.remove_Click($importsecondToolStripMenuItem1_Click)
			$importcompareToolStripMenuItem2.remove_Click($importcompareToolStripMenuItem2_Click)
			$importexclusionListToolStripMenuItem1.remove_Click($importexclusionListToolStripMenuItem1_Click)
			$ARPToolStripMenuItem.remove_Click($ARPToolStripMenuItem_Click)
			$CMDToolStripMenuItem.remove_Click($CMDToolStripMenuItem_Click)
			$controlPanelToolStripMenuItem.remove_Click($controlPanelToolStripMenuItem_Click)
			$driversToolStripMenuItem.remove_Click($driversToolStripMenuItem_Click)
			$eventViewerToolStripMenuItem.remove_Click($eventViewerToolStripMenuItem_Click)
			$FTAsToolStripMenuItem.remove_Click($FTAsToolStripMenuItem_Click)
			$MMCToolStripMenuItem.remove_Click($MMCToolStripMenuItem_Click)
			$processesToolStripMenuItem.remove_Click($processesToolStripMenuItem_Click)
			$registryToolStripMenuItem.remove_Click($registryToolStripMenuItem_Click)
			$servicesToolStripMenuItem.remove_Click($servicesToolStripMenuItem_Click)
			$showToolStripMenuItem.remove_Click($showToolStripMenuItem_Click)
			$minimiseToolStripMenuItem.remove_Click($minimiseToolStripMenuItem_Click)
			$maximiseToolStripMenuItem.remove_Click($maximiseToolStripMenuItem_Click)
			$closeToolStripMenuItem.remove_Click($closeToolStripMenuItem_Click)
			$backgroundColourToolStripMenuItem.remove_Click($backgroundColourToolStripMenuItem_Click)
			$toolstripcombobox3.remove_SelectedIndexChanged($toolstripcombobox3_SelectedIndexChanged)
			$VirtualMachineToolStripMenuItem.remove_Click($VirtualMachineToolStripMenuItem_Click)
			$usernametoolstriptextbox2.remove_TextChanged($usernametoolstriptextbox2_TextChanged)
			$toolstriptextbox3.remove_TextChanged($toolstriptextbox3_TextChanged)
			$moveToBottomToolStripMenuItem.remove_Click($moveToBottomToolStripMenuItem_Click)
			$excludeToolStripMenuItem.remove_Click($excludeToolStripMenuItem_Click)
			$showAllToolStripMenuItem.remove_Click($showAllToolStripMenuItem_Click)
			$showExclusionsToolStripMenuItem.remove_Click($showExclusionsToolStripMenuItem_Click)
			$sendToDiscoveryToolStripMenuItem.remove_Click($sendToDiscoveryToolStripMenuItem_Click)
			$formDiscoveryTool.remove_Load($Form_StateCorrection_Load)
			$formDiscoveryTool.remove_FormClosed($Form_Cleanup_FormClosed)
			$timer1.remove_Tick($timer1_Tick)
		}
		catch [Exception]
		{ }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	#
	# formDiscoveryTool
	#
	$formDiscoveryTool.Controls.Add($Script:datagridview1)
	$formDiscoveryTool.Controls.Add($menustrip1)
	$formDiscoveryTool.Controls.Add($menustrip2)
	$formDiscoveryTool.Controls.Add($combobox1)
	$formDiscoveryTool.AutoSize = $True
	$formDiscoveryTool.BackColor = 'Silver'
	$formDiscoveryTool.ClientSize = '971, 88'
	$formDiscoveryTool.Font = "Calibri, 8.25pt"
	$formDiscoveryTool.ForeColor = 'Black'
	#region Binary Data
	$formDiscoveryTool.Icon = $icon
	#endregion
	$formDiscoveryTool.MainMenuStrip = $menustrip1
	$formDiscoveryTool.MaximizeBox = $False
	$formDiscoveryTool.MinimizeBox = $False
	$formDiscoveryTool.MinimumSize = '971, 88'
	$formDiscoveryTool.Name = "formDiscoveryTool"
	$formDiscoveryTool.Text = "Discovery Tool"
	$formDiscoveryTool.TransparencyKey = 'Silver'
	$formDiscoveryTool.add_Load($formDiscoveryTool_Load)
	$formDiscoveryTool.add_Resize($formDiscoveryTool_Resize)
    $formDiscoveryTool.KeyPreview = $True
    $formDiscoveryTool.add_KeyDown($formDiscoveryTool_KeyDown)
	#
	# datagridview1
	#
	$Script:datagridview1.AllowUserToAddRows = $False
	$Script:datagridview1.AllowUserToDeleteRows = $False
	$Script:datagridview1.AutoSizeColumnsMode = 'Fill'
	$Script:datagridview1.AutoSizeRowsMode = 'AllCells'
	$Script:datagridview1.BackgroundColor = 'White'
	$Script:datagridview1.ColumnHeadersBorderStyle = 'Single'
	$System_Windows_Forms_DataGridViewCellStyle_1 = New-Object 'System.Windows.Forms.DataGridViewCellStyle'
	$System_Windows_Forms_DataGridViewCellStyle_1.Alignment = 'MiddleLeft'
	$System_Windows_Forms_DataGridViewCellStyle_1.BackColor = 'White'
	$System_Windows_Forms_DataGridViewCellStyle_1.Font = "Calibri, 8.25pt"
	$System_Windows_Forms_DataGridViewCellStyle_1.ForeColor = 'WindowText'
	$System_Windows_Forms_DataGridViewCellStyle_1.SelectionBackColor = 'Highlight'
	$System_Windows_Forms_DataGridViewCellStyle_1.SelectionForeColor = 'HighlightText'
	$System_Windows_Forms_DataGridViewCellStyle_1.WrapMode = 'True'
	$Script:datagridview1.ColumnHeadersDefaultCellStyle = $System_Windows_Forms_DataGridViewCellStyle_1
	$Script:datagridview1.ColumnHeadersHeightSizeMode = 'AutoSize'
	[void]$Script:datagridview1.Columns.Add($Script:Selected)
	$Script:datagridview1.ContextMenuStrip = $contextmenustrip1
	$System_Windows_Forms_DataGridViewCellStyle_2 = New-Object 'System.Windows.Forms.DataGridViewCellStyle'
	$System_Windows_Forms_DataGridViewCellStyle_2.Alignment = 'MiddleLeft'
	$System_Windows_Forms_DataGridViewCellStyle_2.BackColor = 'Window'
	$System_Windows_Forms_DataGridViewCellStyle_2.Font = "Calibri, 8.25pt"
	$System_Windows_Forms_DataGridViewCellStyle_2.ForeColor = 'Black'
	$System_Windows_Forms_DataGridViewCellStyle_2.SelectionBackColor = 'PaleTurquoise'
	$System_Windows_Forms_DataGridViewCellStyle_2.SelectionForeColor = 'Black'
	$System_Windows_Forms_DataGridViewCellStyle_2.WrapMode = 'True'
	$Script:datagridview1.DefaultCellStyle = $System_Windows_Forms_DataGridViewCellStyle_2
	$Script:datagridview1.Dock = 'Fill'
	$Script:datagridview1.GridColor = 'White'
	$Script:datagridview1.Location = '0, 60'
	$Script:datagridview1.Name = "datagridview1"
	$Script:datagridview1.RowHeadersVisible = $False
	$Script:datagridview1.RowTemplate.Height = 24
	$Script:datagridview1.Size = '971, 0'
	$Script:datagridview1.TabIndex = 2
	$Script:datagridview1.Visible = $False
    #$Script:datagridview1.ReadOnly = $True
	#
	# menustrip1
	#
	$menustrip1.BackColor = 'Azure'
	$menustrip1.Font = "Calibri, 8.25pt"
	[void]$menustrip1.Items.Add($captureFirstToolStripMenuItem)
	[void]$menustrip1.Items.Add($captureSecondToolStripMenuItem)
	[void]$menustrip1.Items.Add($compareToolStripMenuItem)
	[void]$menustrip1.Items.Add($sendToDiscoveryToolStripMenuItem)
	[void]$menustrip1.Items.Add($DefaultTexttoolstripcombobox1)
	[void]$menustrip1.Items.Add($toolstripcombobox2)
	[void]$menustrip1.Items.Add($Script:captureWindowToolStripMenuItem)
	[void]$menustrip1.Items.Add($captureFullScreenToolStripMenuItem)
	[void]$menustrip1.Items.Add($moveToBottomToolStripMenuItem)
	[void]$menustrip1.Items.Add($showToolStripMenuItem)
	$menustrip1.LayoutStyle = 'HorizontalStackWithOverflow'
	$menustrip1.Location = '0, 28'
	$menustrip1.Name = "menustrip1"
	$menustrip1.Size = '971, 32'
	$menustrip1.TabIndex = 0
	$menustrip1.Text = "menustrip1"
	#
	# menustrip2
	#
	$menustrip2.BackColor = 'Azure'
	$menustrip2.Font = "Calibri, 8.25pt"
	[void]$menustrip2.Items.Add($optionsToolStripMenuItem1)
	[void]$menustrip2.Items.Add($vMControlToolStripMenuItem1)
	[void]$menustrip2.Items.Add($VirtualMachineToolStripMenuItem)
	[void]$menustrip2.Items.Add($toolstriptextbox1)
	[void]$menustrip2.Items.Add($closeToolStripMenuItem)
	[void]$menustrip2.Items.Add($maximiseToolStripMenuItem)
	[void]$menustrip2.Items.Add($minimiseToolStripMenuItem)
	$menustrip2.Location = '0, 0'
	$menustrip2.Name = "menustrip2"
	$menustrip2.Size = '971, 28'
	$menustrip2.TabIndex = 1
	$menustrip2.Text = "menustrip2"
	#
	# combobox1
	#
	$combobox1.FormattingEnabled = $True
	$combobox1.Location = '534, 0'
	$combobox1.Name = "combobox1"
	$combobox1.Size = '121, 21'
	$combobox1.TabIndex = 4
	#
	# captureFirstToolStripMenuItem
	#
	$captureFirstToolStripMenuItem.Name = "captureFirstToolStripMenuItem"
	$captureFirstToolStripMenuItem.Size = '92, 28'
	$captureFirstToolStripMenuItem.Text = "Capture First"
	$captureFirstToolStripMenuItem.add_Click($captureFirstToolStripMenuItem_Click)
	#
	# captureSecondToolStripMenuItem
	#
	$captureSecondToolStripMenuItem.Name = "captureSecondToolStripMenuItem"
	$captureSecondToolStripMenuItem.Size = '108, 28'
	$captureSecondToolStripMenuItem.Text = "Capture Second"
	$captureSecondToolStripMenuItem.add_Click($captureSecondToolStripMenuItem_Click)
	#
	# compareToolStripMenuItem
	#
	$compareToolStripMenuItem.Name = "compareToolStripMenuItem"
	$compareToolStripMenuItem.Size = '71, 28'
	$compareToolStripMenuItem.Text = "Compare"
	$compareToolStripMenuItem.add_Click($compareToolStripMenuItem_Click)
	#
	# DefaultTexttoolstripcombobox1
	#
	$DefaultTexttoolstripcombobox1.DropDownStyle = 'DropDownList'
	$DefaultTexttoolstripcombobox1.Font = "Calibri, 8.25pt"
	$DefaultTexttoolstripcombobox1.Name = "DefaultTexttoolstripcombobox1"
	$DefaultTexttoolstripcombobox1.Size = '121, 28'
	$DefaultTexttoolstripcombobox1.add_SelectedIndexChanged($DefaultTexttoolstripcombobox1_SelectedIndexChanged)
    $DefaultTexttoolstripcombobox1.CausesValidation = $false
	#
	# captureWindowToolStripMenuItem
	#
	$Script:captureWindowToolStripMenuItem.Name = "captureWindowToolStripMenuItem"
	$Script:captureWindowToolStripMenuItem.Size = '114, 28'
	$Script:captureWindowToolStripMenuItem.Text = "Capture Window"
	$Script:captureWindowToolStripMenuItem.add_Click($Script:captureWindowToolStripMenuItem_Click)
	#
	# captureFullScreenToolStripMenuItem
	#
	$captureFullScreenToolStripMenuItem.Name = "captureFullScreenToolStripMenuItem"
	$captureFullScreenToolStripMenuItem.Size = '128, 28'
	$captureFullScreenToolStripMenuItem.Text = "Capture Full Screen"
	$captureFullScreenToolStripMenuItem.add_Click($captureFullScreenToolStripMenuItem_Click)
	#
	# optionsToolStripMenuItem1
	#
	[void]$optionsToolStripMenuItem1.DropDownItems.Add($exportToolStripMenuItem)
	[void]$optionsToolStripMenuItem1.DropDownItems.Add($importToolStripMenuItem)
	[void]$optionsToolStripMenuItem1.DropDownItems.Add($customiseCaptureTextToolStripMenuItem)
	[void]$optionsToolStripMenuItem1.DropDownItems.Add($LoadDocumentationToolStripMenuItem)
	[void]$optionsToolStripMenuItem1.DropDownItems.Add($backgroundColourToolStripMenuItem)
	[void]$optionsToolStripMenuItem1.DropDownItems.Add($toolstripcombobox3)
	$optionsToolStripMenuItem1.Name = "optionsToolStripMenuItem1"
	$optionsToolStripMenuItem1.Size = '63, 24'
	$optionsToolStripMenuItem1.Text = "Options"
	#
	# vMControlToolStripMenuItem1
	#
	[void]$vMControlToolStripMenuItem1.DropDownItems.Add($ARPToolStripMenuItem)
	[void]$vMControlToolStripMenuItem1.DropDownItems.Add($CMDToolStripMenuItem)
	[void]$vMControlToolStripMenuItem1.DropDownItems.Add($controlPanelToolStripMenuItem)
	[void]$vMControlToolStripMenuItem1.DropDownItems.Add($driversToolStripMenuItem)
	[void]$vMControlToolStripMenuItem1.DropDownItems.Add($eventViewerToolStripMenuItem)
	[void]$vMControlToolStripMenuItem1.DropDownItems.Add($FTAsToolStripMenuItem)
	[void]$vMControlToolStripMenuItem1.DropDownItems.Add($MMCToolStripMenuItem)
	[void]$vMControlToolStripMenuItem1.DropDownItems.Add($processesToolStripMenuItem)
	[void]$vMControlToolStripMenuItem1.DropDownItems.Add($registryToolStripMenuItem)
	[void]$vMControlToolStripMenuItem1.DropDownItems.Add($servicesToolStripMenuItem)
	$vMControlToolStripMenuItem1.Name = "vMControlToolStripMenuItem1"
	$vMControlToolStripMenuItem1.Size = '84, 24'
	$vMControlToolStripMenuItem1.Text = "VM Control"
	#
	# toolstriptextbox1
	#
	$toolstriptextbox1.BackColor = 'Azure'
	$toolstriptextbox1.Font = "Calibri, 8.25pt"
	$toolstriptextbox1.Name = "toolstriptextbox1"
	$toolstriptextbox1.ReadOnly = $True
	$toolstriptextbox1.Size = '250, 24'
	#
	# exportToolStripMenuItem
	#
	$exportToolStripMenuItem.BackColor = 'Control'
	[void]$exportToolStripMenuItem.DropDownItems.Add($exportfirstToolStripMenuItem)
	[void]$exportToolStripMenuItem.DropDownItems.Add($exportsecondToolStripMenuItem)
	[void]$exportToolStripMenuItem.DropDownItems.Add($exportcompareToolStripMenuItem1)
	[void]$exportToolStripMenuItem.DropDownItems.Add($exportexclusionListToolStripMenuItem)
	$exportToolStripMenuItem.Name = "exportToolStripMenuItem"
	$exportToolStripMenuItem.Size = '208, 22'
	$exportToolStripMenuItem.Text = "Export"
	#
	# importToolStripMenuItem
	#
	$importToolStripMenuItem.BackColor = 'Control'
	[void]$importToolStripMenuItem.DropDownItems.Add($importfirstToolStripMenuItem1)
	[void]$importToolStripMenuItem.DropDownItems.Add($importsecondToolStripMenuItem1)
	[void]$importToolStripMenuItem.DropDownItems.Add($importcompareToolStripMenuItem2)
	[void]$importToolStripMenuItem.DropDownItems.Add($importexclusionListToolStripMenuItem1)
	$importToolStripMenuItem.Name = "importToolStripMenuItem"
	$importToolStripMenuItem.Size = '208, 22'
	$importToolStripMenuItem.Text = "Import"
	#
	# customiseCaptureTextToolStripMenuItem
	#
	$customiseCaptureTextToolStripMenuItem.BackColor = 'Control'
	$customiseCaptureTextToolStripMenuItem.Name = "customiseCaptureTextToolStripMenuItem"
	$customiseCaptureTextToolStripMenuItem.Size = '208, 22'
	$customiseCaptureTextToolStripMenuItem.Text = "Customise Capture Text"
	$customiseCaptureTextToolStripMenuItem.add_Click($customiseCaptureTextToolStripMenuItem_Click)
	#
	# LoadDocumentationToolStripMenuItem
	#
	$LoadDocumentationToolStripMenuItem.BackColor = 'Control'
	$LoadDocumentationToolStripMenuItem.Name = "LoadDocumentationToolStripMenuItem"
	$LoadDocumentationToolStripMenuItem.Size = '208, 22'
	$LoadDocumentationToolStripMenuItem.Text = "Documentation"
	$LoadDocumentationToolStripMenuItem.add_Click($LoadDocumentationToolStripMenuItem_Click)
	#
	# exportfirstToolStripMenuItem
	#
	$exportfirstToolStripMenuItem.BackColor = 'Control'
	$exportfirstToolStripMenuItem.Name = "exportfirstToolStripMenuItem"
	$exportfirstToolStripMenuItem.Size = '127, 22'
	$exportfirstToolStripMenuItem.Text = "First"
	$exportfirstToolStripMenuItem.add_Click($exportfirstToolStripMenuItem_Click)
	#
	# exportsecondToolStripMenuItem
	#
	$exportsecondToolStripMenuItem.BackColor = 'Control'
	$exportsecondToolStripMenuItem.Name = "exportsecondToolStripMenuItem"
	$exportsecondToolStripMenuItem.Size = '127, 22'
	$exportsecondToolStripMenuItem.Text = "Second"
	$exportsecondToolStripMenuItem.add_Click($exportsecondToolStripMenuItem_Click)
	#
	# exportcompareToolStripMenuItem1
	#
	$exportcompareToolStripMenuItem1.BackColor = 'Control'
	$exportcompareToolStripMenuItem1.Name = "exportcompareToolStripMenuItem1"
	$exportcompareToolStripMenuItem1.Size = '127, 22'
	$exportcompareToolStripMenuItem1.Text = "Compare"
	$exportcompareToolStripMenuItem1.add_Click($exportcompareToolStripMenuItem1_Click)
	#
	# exportexclusionListToolStripMenuItem
	#
	$exportexclusionListToolStripMenuItem.BackColor = 'Control'
	$exportexclusionListToolStripMenuItem.Name = "exportexclusionListToolStripMenuItem"
	$exportexclusionListToolStripMenuItem.Size = '152, 22'
	$exportexclusionListToolStripMenuItem.Text = "Exclusion List"
	$exportexclusionListToolStripMenuItem.add_Click($exportexclusionListToolStripMenuItem_Click)
	#
	# importfirstToolStripMenuItem1
	#
	$importfirstToolStripMenuItem1.Name = "importfirstToolStripMenuItem1"
	$importfirstToolStripMenuItem1.Size = '127, 22'
	$importfirstToolStripMenuItem1.Text = "First"
	$importfirstToolStripMenuItem1.add_Click($importfirstToolStripMenuItem1_Click)
	#
	# importsecondToolStripMenuItem1
	#
	$importsecondToolStripMenuItem1.Name = "importsecondToolStripMenuItem1"
	$importsecondToolStripMenuItem1.Size = '127, 22'
	$importsecondToolStripMenuItem1.Text = "Second"
	$importsecondToolStripMenuItem1.add_Click($importsecondToolStripMenuItem1_Click)
	#
	# importcompareToolStripMenuItem2
	#
	$importcompareToolStripMenuItem2.Name = "importcompareToolStripMenuItem2"
	$importcompareToolStripMenuItem2.Size = '127, 22'
	$importcompareToolStripMenuItem2.Text = "Compare"
	$importcompareToolStripMenuItem2.add_Click($importcompareToolStripMenuItem2_Click)
	#
	# importexclusionListToolStripMenuItem1
	#
	$importexclusionListToolStripMenuItem1.Name = "importexclusionListToolStripMenuItem1"
	$importexclusionListToolStripMenuItem1.Size = '152, 22'
	$importexclusionListToolStripMenuItem1.Text = "Exclusion List"
	$importexclusionListToolStripMenuItem1.add_Click($importexclusionListToolStripMenuItem1_Click)
	#
	# ARPToolStripMenuItem
	#
	$ARPToolStripMenuItem.Name = "ARPToolStripMenuItem"
	$ARPToolStripMenuItem.Size = '151, 22'
	$ARPToolStripMenuItem.Text = "ARP"
	$ARPToolStripMenuItem.add_Click($ARPToolStripMenuItem_Click)
	#
	# CMDToolStripMenuItem
	#
	$CMDToolStripMenuItem.Name = "CMDToolStripMenuItem"
	$CMDToolStripMenuItem.Size = '151, 22'
	$CMDToolStripMenuItem.Text = "CMD"
	$CMDToolStripMenuItem.add_Click($CMDToolStripMenuItem_Click)
	#
	# controlPanelToolStripMenuItem
	#
	$controlPanelToolStripMenuItem.Name = "controlPanelToolStripMenuItem"
	$controlPanelToolStripMenuItem.Size = '151, 22'
	$controlPanelToolStripMenuItem.Text = "Control Panel"
	$controlPanelToolStripMenuItem.add_Click($controlPanelToolStripMenuItem_Click)
	#
	# driversToolStripMenuItem
	#
	$driversToolStripMenuItem.Name = "driversToolStripMenuItem"
	$driversToolStripMenuItem.Size = '151, 22'
	$driversToolStripMenuItem.Text = "Drivers"
	$driversToolStripMenuItem.add_Click($driversToolStripMenuItem_Click)
	#
	# eventViewerToolStripMenuItem
	#
	$eventViewerToolStripMenuItem.Name = "eventViewerToolStripMenuItem"
	$eventViewerToolStripMenuItem.Size = '151, 22'
	$eventViewerToolStripMenuItem.Text = "Event Viewer"
	$eventViewerToolStripMenuItem.add_Click($eventViewerToolStripMenuItem_Click)
	#
	# FTAsToolStripMenuItem
	#
	$FTAsToolStripMenuItem.Name = "FTAsToolStripMenuItem"
	$FTAsToolStripMenuItem.Size = '151, 22'
	$FTAsToolStripMenuItem.Text = "FTAs"
	$FTAsToolStripMenuItem.add_Click($FTAsToolStripMenuItem_Click)
	#
	# MMCToolStripMenuItem
	#
	$MMCToolStripMenuItem.Name = "MMCToolStripMenuItem"
	$MMCToolStripMenuItem.Size = '151, 22'
	$MMCToolStripMenuItem.Text = "MMC"
	$MMCToolStripMenuItem.add_Click($MMCToolStripMenuItem_Click)
	#
	# processesToolStripMenuItem
	#
	$processesToolStripMenuItem.Name = "processesToolStripMenuItem"
	$processesToolStripMenuItem.Size = '151, 22'
	$processesToolStripMenuItem.Text = "Processes"
	$processesToolStripMenuItem.add_Click($processesToolStripMenuItem_Click)
	#
	# registryToolStripMenuItem
	#
	$registryToolStripMenuItem.Name = "registryToolStripMenuItem"
	$registryToolStripMenuItem.Size = '151, 22'
	$registryToolStripMenuItem.Text = "Registry"
	$registryToolStripMenuItem.add_Click($registryToolStripMenuItem_Click)
	#
	# servicesToolStripMenuItem
	#
	$servicesToolStripMenuItem.Name = "servicesToolStripMenuItem"
	$servicesToolStripMenuItem.Size = '152, 22'
	$servicesToolStripMenuItem.Text = "Services"
	$servicesToolStripMenuItem.add_Click($servicesToolStripMenuItem_Click)
	#
	# showToolStripMenuItem
	#
	$showToolStripMenuItem.Alignment = 'Right'
	$showToolStripMenuItem.Name = "showToolStripMenuItem"
	$showToolStripMenuItem.Size = '50, 28'
	$showToolStripMenuItem.Text = "Show"
	$showToolStripMenuItem.add_Click($showToolStripMenuItem_Click)
	#
	# minimiseToolStripMenuItem
	#
	$minimiseToolStripMenuItem.Alignment = 'Right'
	$minimiseToolStripMenuItem.Name = "minimiseToolStripMenuItem"
	$minimiseToolStripMenuItem.Size = '71, 24'
	$minimiseToolStripMenuItem.Text = "Minimise"
	$minimiseToolStripMenuItem.add_Click($minimiseToolStripMenuItem_Click)
	#
	# maximiseToolStripMenuItem
	#
	$maximiseToolStripMenuItem.Alignment = 'Right'
	$maximiseToolStripMenuItem.Name = "maximiseToolStripMenuItem"
	$maximiseToolStripMenuItem.Size = '74, 24'
	$maximiseToolStripMenuItem.Text = "Maximise"
	$maximiseToolStripMenuItem.add_Click($maximiseToolStripMenuItem_Click)
	#
	# closeToolStripMenuItem
	#
	$closeToolStripMenuItem.Alignment = 'Right'
	$closeToolStripMenuItem.Name = "closeToolStripMenuItem"
	$closeToolStripMenuItem.Size = '49, 24'
	$closeToolStripMenuItem.Text = "Close"
	$closeToolStripMenuItem.Visible = $False
	$closeToolStripMenuItem.add_Click($closeToolStripMenuItem_Click)
	#
	# backgroundColourToolStripMenuItem
	#
	$backgroundColourToolStripMenuItem.Name = "backgroundColourToolStripMenuItem"
	$backgroundColourToolStripMenuItem.Size = '208, 22'
	$backgroundColourToolStripMenuItem.Text = "Background Colour  "
	$backgroundColourToolStripMenuItem.add_Click($backgroundColourToolStripMenuItem_Click)
	#
	# colordialog1
	#
	#
	# toolstripcombobox3
	#
	$toolstripcombobox3.DropDownStyle = 'DropDownList'
	[void]$toolstripcombobox3.Items.Add("Filtered")
	[void]$toolstripcombobox3.Items.Add("Exclusions")
	[void]$toolstripcombobox3.Items.Add("All")
	$toolstripcombobox3.Font = "Calibri, 8.25pt"
	$toolstripcombobox3.Name = "toolstripcombobox3"
	$toolstripcombobox3.Size = '121, 28'
	$toolstripcombobox3.add_SelectedIndexChanged($toolstripcombobox3_SelectedIndexChanged)
	#
	# VirtualMachineToolStripMenuItem
	#
	[void]$VirtualMachineToolStripMenuItem.DropDownItems.Add($Script:VMselectiontoolstripcombobox4)
	[void]$VirtualMachineToolStripMenuItem.DropDownItems.Add($usernametoolstriptextbox2)
	[void]$VirtualMachineToolStripMenuItem.DropDownItems.Add($toolstriptextbox3)
	$VirtualMachineToolStripMenuItem.Name = "VirtualMachineToolStripMenuItem"
	$VirtualMachineToolStripMenuItem.Size = '110, 24'
	$VirtualMachineToolStripMenuItem.Text = "Virtual Machine"
	$VirtualMachineToolStripMenuItem.add_Click($VirtualMachineToolStripMenuItem_Click)
	#
	# VMselectiontoolstripcombobox4
	#
	$Script:VMselectiontoolstripcombobox4.DropDownStyle = 'DropDownList'
	$Script:VMselectiontoolstripcombobox4.Font = "Calibri, 8.25pt"
	$Script:VMselectiontoolstripcombobox4.Name = "VMselectiontoolstripcombobox4"
	$Script:VMselectiontoolstripcombobox4.Size = '121, 25'
	$Script:VMselectiontoolstripcombobox4.add_SelectedIndexChanged($Script:VMselectiontoolstripcombobox4_SelectedIndexChanged)
	#
	# usernametoolstriptextbox2
	#
	$usernametoolstriptextbox2.Font = "Calibri, 8.25pt"
	$usernametoolstriptextbox2.Name = "usernametoolstriptextbox2"
	$usernametoolstriptextbox2.Size = '100, 24'
	$usernametoolstriptextbox2.Text = "Packaging"
	$usernametoolstriptextbox2.add_TextChanged($usernametoolstriptextbox2_TextChanged)
	#
	# toolstriptextbox3
	#
	$toolstriptextbox3.Name = "toolstriptextbox3"
	$toolstriptextbox3.Size = '100, 27'
	$toolstriptextbox3.add_TextChanged($toolstriptextbox3_TextChanged)
	#
	# moveToBottomToolStripMenuItem
	#
	$moveToBottomToolStripMenuItem.Alignment = 'Right'
	$moveToBottomToolStripMenuItem.Name = "moveToBottomToolStripMenuItem"
	$moveToBottomToolStripMenuItem.Size = '113, 28'
	$moveToBottomToolStripMenuItem.Text = "Move To Bottom"
	$moveToBottomToolStripMenuItem.Visible = $False
	$moveToBottomToolStripMenuItem.add_Click($moveToBottomToolStripMenuItem_Click)
	#
	# contextmenustrip1
	#
	[void]$contextmenustrip1.Items.Add($excludeToolStripMenuItem)
	[void]$contextmenustrip1.Items.Add($showAllToolStripMenuItem)
	[void]$contextmenustrip1.Items.Add($showExclusionsToolStripMenuItem)
	$contextmenustrip1.Name = "contextmenustrip1"
	$contextmenustrip1.Font = "Calibri, 8.25pt"
	$contextmenustrip1.Size = '137, 52'
	#
	# excludeToolStripMenuItem
	#
	$excludeToolStripMenuItem.Name = "excludeToolStripMenuItem"
	$excludeToolStripMenuItem.Size = '136, 24'
	$excludeToolStripMenuItem.Text = "Include/Exclude"
	$excludeToolStripMenuItem.add_Click($excludeToolStripMenuItem_Click)
	#
	# showAllToolStripMenuItem
	#
	$showAllToolStripMenuItem.Name = "showAllToolStripMenuItem"
	$showAllToolStripMenuItem.Size = '136, 24'
	$showAllToolStripMenuItem.Text = "Show All"
	$showAllToolStripMenuItem.add_Click($showAllToolStripMenuItem_Click)
	#
	# showExclusionsToolStripMenuItem
	#
	$showExclusionsToolStripMenuItem.Name = "showExclusionsToolStripMenuItem"
	$showExclusionsToolStripMenuItem.Size = '185, 24'
	$showExclusionsToolStripMenuItem.Text = "Show Exclusions"
	$showExclusionsToolStripMenuItem.add_Click($showExclusionsToolStripMenuItem_Click)
	#
	# CheckColumn
	#
	$Script:Selected.HeaderText = "Selected"
	$Script:Selected.Name = "Selected"
    $Script:Selected.TrueValue = $true
	$Script:Selected.Width = 62
    $Script:Selected.AutoSizeMode = 'ColumnHeader'
    $Script:Selected.ReadOnly = $false
    #
    #SaveFileDialog
    #
    $initialDirectory = "$env:Desktop"
    $SaveFileDialog.initialDirectory = $initialDirectory
    $SaveFileDialog.filter = “XML File (*.xml)|*.xml”
	#
	# openfiledialog1
	#
    $OpenFileDialog1.initialDirectory = $initialDirectory
    $OpenFileDialog1.filter = “Docx File (*.Docx)|*.Docx”
    #
    # openfiledialog
    #
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = “XML File (*.xml)|*.xml”
	#
	# toolstripcombobox2
	#
	$toolstripcombobox2.DropDownStyle = 'DropDownList'
	$toolstripcombobox2.Name = "toolstripcombobox2"
	$toolstripcombobox2.Size = '121, 28'
	$toolstripcombobox2.Font = "Calibri, 8.25pt"
	[void]$toolstripcombobox2.Items.Add("Installation Steps")
	[void]$toolstripcombobox2.Items.Add("Start Menu")
	[void]$toolstripcombobox2.Items.Add("First Launch Test & Post Configuration")
	[void]$toolstripcombobox2.Items.Add("About Menu")
	[void]$toolstripcombobox2.Items.Add("Help Menu")
	[void]$toolstripcombobox2.Items.Add("Control Panel Entry")
	[void]$toolstripcombobox2.Items.Add("Context Menu")
	[void]$toolstripcombobox2.Items.Add("Functional Testing")
	[void]$toolstripcombobox2.Items.Add("Place At Cursor")
	$toolstripcombobox2.add_SelectedIndexChanged($toolstripcombobox2_SelectedIndexChanged)
	#
	# timer1
	#
	$timer1.add_Tick($timer1_Tick)
    $timer1.Interval = 3000
    #
	# sendToDiscoveryToolStripMenuItem
	#
	$sendToDiscoveryToolStripMenuItem.Name = "sendToDiscoveryToolStripMenuItem"
	$sendToDiscoveryToolStripMenuItem.Size = '102, 23'
	$sendToDiscoveryToolStripMenuItem.Text = "Send Selected To Discovery"
	$sendToDiscoveryToolStripMenuItem.visible = $Flase
	$sendToDiscoveryToolStripMenuItem.add_Click($sendToDiscoveryToolStripMenuItem_Click)
    #
    #
    #
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $formDiscoveryTool.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$formDiscoveryTool.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$formDiscoveryTool.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	return $formDiscoveryTool.ShowDialog()

} #End Function

#Call OnApplicationLoad to initialize
if((OnApplicationLoad) -eq $true)
{
	#Call the form
	Call-Doc_Cap_pff | Out-Null
	#Perform cleanup
	OnApplicationExit
}
}else{
Invoke-MessageBox -FormHeader "No VMRun.exe" -Message "Cound not find 'VMRun.exe in '$VMRunPath'`nPlease ensure you have VMware installed and the 'VMRun.exe' is in this location" -ButtonOne "Okay"

}