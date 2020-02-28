




$credential = Get-Credential

##### Set Credentials
$credential.Password | ConvertFrom-SecureString | Set-Content C:\HypervCopy\encrypted_password.txt
$User1 = "Badger\Dave"
$encrypted = Get-Content C:\HypervCopy\encrypted_password.txt | ConvertTo-SecureString
$credential = New-Object System.Management.Automation.PsCredential($user1, $encrypted)

####vm names
$VMname1 = "Sharepointserver1"
$VMName2 = "Autoseq"
get-vm -Name $VMname1 | Select-Object -Property vmid
$VMID1= "c204e19d-4c01-4dbe-bddd-5bbf80ea7238"


#### vm commands#####

#start VM
Start-VM $VMname1
#stop vm
Stop-vm $VMname1

#Copy item to session
Copy-Item -ToSession $s -path C:\HypervCopy -Destination C:\HypervCopy
#Copy item from session
Copy-Item -FromSession $s -path C:\HypervCopy -Destination C:\HypervCopy
#recursive
Copy-Item -ToSession $s -Path C:\HypervCopy\Source -Recurse -Destination C:\HypervCopy -Force

#create Vm Checkpoint
get-vm | Checkpoint-VM -SnapshotName "$vmname1`_$(get-date -format 'dd-MM-yyyy-hh-mm')"
#restore vm checkpoint
Restore-VMCheckpoint -Name "Base" -VMName $VMname1 -Confirm

get-command *pssession

#get vm ID
get-vm Sharepointserver1 | Select id

#Create PSSession
$PSsession1 = New-PSSession -VMName $VMname1 -Credential $credential
$S = New-PSSession -VMName $VMname1 -Credential $credential
#Remove Session
Remove-PSSession $s
#Enter Session
Enter-PSSession $s
#Exit Session
Exit-PSSession 


######Sharepoint configuration
Install-Module sharepointpnppowershell2019
get-module sharepointpnppowershell2019
get-command -module sharepointpnppowershell2019
$sharepointURl = "http://sharepointserve"
Connect-PnPOnline -Url -credentials $credential
## current conenction if on VM
Connect-PnPOnline -url $sharepointURl -CurrentCredentials
### connection to site pass creds
Connect-PnPOnline -Url $sharepointURl -Credentials $credential
#get sharepoint list
Get-PnPList -Identity "139af52b-697b-4edc-b469-4d77e2c20e8b"
#remove sharepoint list
Remove-PnPList -Identity "139af52b-697b-4edc-b469-4d77e2c20e8b"
#new sharepoint list

$SharepointList = "Badger"
New-PnPList -Title $SharepointList -Url $SharepointList -Template GenericList -OnQuickLaunch

$SharepointList = "Badger"

##crate column for list
Add-PnPField -list $SharepointList -DisplayName "Vendor" -InternalName "Vendor" -type Text -AddToDefaultView 
Add-PnPField -list $SharepointList -DisplayName "AppName" -InternalName "AppName" -type Text -AddToDefaultView 
Add-PnPField -list $SharepointList -DisplayName "Version" -InternalName "Version" -type Number -AddToDefaultView
Add-PnPField -list $SharepointList -DisplayName "BrickName" -InternalName "BrickName" -type Text -AddToDefaultView  
Add-PnPField -list $SharepointList -DisplayName "Release" -InternalName "Release" -type Text  -AddToDefaultView 
Add-PnPField -list $SharepointList -DisplayName "AppType" -InternalName "AppType" -type MultiChoice -AddToDefaultView -Choices "MSI","Scripted", "AppV5", "AppV4", "Other"
Add-PnPField -list $SharepointList -DisplayName "AD Group" -InternalName "ADGroup" -type Text -AddToDefaultView 
Add-PnPField -list $SharepointList -DisplayName "Assigned to:" -InternalName "AssignedTo" -type User -AddToDefaultView 
Add-PnPField -list $SharepointList -DisplayName "Architecture" -InternalName "Architecture" -type MultiChoice -AddToDefaultView -Choices "x86","x64", "Both"
Add-PnPField -list $SharepointList -DisplayName "Target" -InternalName "Target" -type MultiChoice -AddToDefaultView -Choices "Windows 7 32bit","Windows 7 64bit", "Windows 10", "Mac"
#example # PS:>Add-PnPField -List "Demo list" -DisplayName "Speakers" -InternalName "SPSSpeakers" -Type MultiChoice -Group "Demo Group" -AddToDefaultView -Choices "Obiwan Kenobi","Darth Vader", "Anakin Skywalker"

Add-PnPListItem -List $Badger -ContentType 
get-pnplistitem -List "badger"
Set-PnPField -List $SharepointList 

[String]$PslistAppName = 'Frogger'
[String]$PslistVendor = 'Apper'
[String]$PslistVersion = '12.0.29'
[String]$PSListRelease = 'R02'
[String]$PSListArch = "X86"
[String]$PSListAppType = "AppV5"
#[String]$PSListUserAssigned = "User02"
[String]$PSListUserAssigned = "User01"

[String]$PSListTarget = "Windows 10"

switch ($PSListAppType){
    'MSI'        {$PSListAppTypeSwitch = 'MSI'}
    'Scripted'   {$PSListAppTypeSwitch = 'SCP'}
    'AppV5'      {$PSListAppTypeSwitch = 'AV5'}
    'AppV4'      {$PSListAppTypeSwitch = 'AV4'}
    'Other'      {$PSListAppTypeSwitch = 'Oth'}    
}
#CREATE brick Name
[String]$PSListBrickName = "$PslistVendor`_$PslistAppName`_$PslistVersion`_$PSListAppTypeSwitch`_$PSlistArch`_$PSListRelease"
#Create ADGroup
[String]$ADGroup = "SFTG-$pslistbrickname"

#Add item to sharepoint list
Add-PnPListItem -List $sharepointlist -Values @{
                                                "Vendor"=$PslistVendor;
                                                "AppName" = $PslistAppName;
                                                "Version"=$PslistVersion;
                                                "BrickName"=$PSListBrickName;
                                                "Release"=$PSListRelease;
                                                "Architecture"=$PSListArch;
                                                "ADGroup"=$ADGroup;
                                                "Target"=$PSListTarget;
                                                "AssignedTo"=$PSListUserAssigned;
                                                "AppType"=$PSListAppType
                                            }



get-pnplistitem -list $SharepointList -Query "$CalmQuery"


$CalmQuery = "<View><Query><Where><Eq><FieldRef Name='AssignedTo' /><Value Type='User'>$PSListUserAssigned</Value></Eq></Where></Query></View>"
get-pnplistitem -list $SharepointList -Query $CalmQuery | fl *


@{
    "Vendor"=$PslistVendor;
    "AppName" = $PslistAppName;
    "Version"=$PslistVersion;
    "BrickName"=$PSListBrickName;
    "Release"=$PSListRelease;
    "Architecture"=$PSListArch;
    "ADGroup"=$ADGroup;
    "Target"=$PSListTarget;
    "AssignedTo"=$PSListUserAssigned;
    "AppType"=$PSListAppType
}


