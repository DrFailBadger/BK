




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
Copy-Item -ToSession $PSsession1 -path C:\HypervCopy\Orac -Destination C:\HypervCopy -Recurse -Force
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
$PSsession1 = New-PSSession -VMName $VMname2 -Credential $credential
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
Import-Module sharepointpnppowershell2019
get-command -module sharepointpnppowershell2019
$sharepointURl = "http://sharepointserve"
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

##crate column for lis
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

Add-PnPField -list $SharepointList -DisplayName "Status" -InternalName "Status" -type MultiChoice -AddToDefaultView -Choices "New","New-Assigned", "Discovery","Discovery On-Hold","Discovery Signed-Off", "Packaging","Packaging on-Hold", "QA", "Pre-Prod","Pre-Prod on-hold", "Prod", "Prod On-hold", "Pilot", "Live", "Completed"
Add-PnPField -list $SharepointList -DisplayName "Requestor" -InternalName "Requestor" -type Text -AddToDefaultView
Add-PnPField -list $SharepointList -DisplayName "Requestor Email" -InternalName "RequestorEmail" -type Text -AddToDefaultView
Add-PnPField -list $SharepointList -DisplayName "Requestor Phone" -InternalName "RequestorPhone" -type Text -AddToDefaultView
#example # PS:>Add-PnPField -List "Demo list" -DisplayName "Speakers" -InternalName "SPSSpeakers" -Type MultiChoice -Group "Demo Group" -AddToDefaultView -Choices "Obiwan Kenobi","Darth Vader", "Anakin Skywalker"

Add-PnPListItem -List $Badger -ContentType 
get-pnplistitem -List "badger"
Set-PnPField -List $SharepointList 

[String]$PslistAppName = 'Frogger2123123'
[String]$PslistVendor = 'Apper2123123'
[String]$PslistVersion = '12.0.222911231'
[String]$PSListRelease = 'R03123123'
[String]$PSListArch = "X86"
[String]$PSListAppType = "AppV5"
#[String]$PSListUserAssigned = "User02"
[String]$PSListUserAssigned = "User01"
[String]$PSListTarget = "Windows 10"
[String]$PSListRequestor = "John Smith"
[String]$PSListRequestorEmail = "John.Smith@badger.com"
[String]$PSListRequestorPhone = "01234567890"
[String]$PSListStatus ="New"
[String]$PSListUserAssigned = $null


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


$PsListValues =  @{
    "Vendor"=$PslistVendor;
    "AppName" = $PslistAppName;
    "Version"=$PslistVersion;
    "BrickName"=$PSListBrickName;
    "Release"=$PSListRelease;
    "Architecture"=$PSListArch;
    "ADGroup"=$ADGroup;
    "Target"=$PSListTarget;
    "AssignedTo"=$PSListUserAssigned;
    "AppType"=$PSListAppType;
    "Requestor"=$PSListRequestor;
    "RequestorEmail"=$PSListRequestorEmail;
    "RequestorPhone"=$PSListRequestorPhone;
    "Status"=$PSListStatus
}
#Add item to sharepoint list
####NEW###
Add-PnPListItem -List $sharepointlist -Values $PsListValues
###Update###
$PSListIdentity = 1
Set-PnplistItem -list $sharepointlist -identity $psListIdentity -Values $PsListValues



switch ($PSListStatus){
    'New'                       {$PSListStatusActual = 'MSI'}
    'New-Assigned'              {$PSListStatusActual = 'SCP'}
    'Discovery'                 {$PSListStatusActual = 'AV5'}
    'Discovery On-Hold'         {$PSListStatusActual = 'AV4'}
    'Discovery Signed-Off'      {$PSListStatusActual = 'Oth'}
    'Packaging'                 {$PSListStatusActual = 'MSI'}
    'Packaging on-Hold'         {$PSListStatusActual = 'SCP'}
    'QA'                        {$PSListStatusActual = 'AV5'}
    'Pre-Prod'                  {$PSListStatusActual = 'AV4'}
    'Pre-Prod on-hold'          {$PSListStatusActual = 'Oth'} 
    'Prod'                      {$PSListStatusActual = 'MSI'}
    'Prod On-Hold'              {$PSListStatusActual = 'SCP'}
    'Pilot'                     {$PSListStatusActual = 'AV5'}
    'Live'                      {$PSListStatusActual = 'AV4'}
    'Completed'                 {$PSListStatusActual = 'AV4'}
}
 

######## where items is not completed


###### where item is assinged to current user



##Set user


$PSListUserAssigned = $null
$PSListUserAssigned = "User01"
$PSListUserAssigned

#When record is assigned to user
$CamlQueryAssigned = 
"<View>
<Query>
    <OrderBy>
    <FieldRef Name='ID' Ascending='FALSE' />
    </OrderBy>
    <Where>
    <Eq>
        <FieldRef Name='AssignedTo' />
        <Value Type='User'>$PSListUserAssigned</Value>
    </Eq>
    </Where>
</Query>
</View>"

# When record is assigned to user but not completed
$CamlQueryAssignedNotCompleted = 
"<View>
<Query> 
    <OrderBy>
    <FieldRef Name='ID' Ascending='FALSE' />
    </OrderBy>
    <Where>
    <And>
    <Eq>
        <FieldRef Name='AssignedTo' />
        <Value Type='User'>$PSListUserAssigned</Value>
    </Eq>
    <Neq>
        <FieldRef Name='Status' />
        <Value Type='Choice'>Completed</Value>
    </Neq>
    </And>
    </Where>
</Query>
</View>"
#all records
$CamlQueryAllRecords =
"<View>
<Query>
<OrderBy>
    <FieldRef Name='ID' />
    </OrderBy>
</Query>
</View>"
#all records not completed
$CamlQueryAllRecordsNotCompleted =
"<View>
<Query>  
    <OrderBy>
    <FieldRef Name='ID' />
    </OrderBy>
    <Where>
    <Neq>
        <FieldRef Name='Status' />
        <Value Type='Choice'>Completed</Value>
    </Neq>
    </Where>
</Query>
</View>"

#New records not assigned
$CamlQueryNotAssignNewStatus = 
"<View>
<Query>  
    <OrderBy>
    <FieldRef Name='ID' />
    </OrderBy>
    <Where>
    <And>
    <Eq>
        <FieldRef Name='Status' />
        <Value Type='Choice'>New</Value>
    </Eq>
    <IsNull>
        <FieldRef Name='AssignedTo' />
    </IsNull>
    </And>
    </Where>
</Query>
</View>"

<#
get-pnplistitem -list $SharepointList -Query $CalmQuery
get-pnplistitem -list $SharepointList $calmq
get-pnplistitem -list $SharepointList -Query $CamlQueryAssigned
get-pnplistitem -list $SharepointList -Query $CamlQueryAssignedNotCompleted 
get-pnplistitem -list $SharepointList -Query $CamlQueryAllRecords
get-pnplistitem -list $SharepointList -Query $CamlQueryAllRecordsNotCompleted

get-pnplistitem -list $SharepointList -Query $CamlQueryNotAssignNewStatus 
#>


#assinged to me completed 3
# all records 0
# all Records Assigned to me 1
# not completed 2
#not completed assigned to me 3
# new records
[int]$Checkbox = 0
[int]$CheckBox += 1
[int]$CheckBox -= 1
[int]$CheckBox += 2
[int]$CheckBox -= 2

switch ($Checkbox) {
    1  { $PSListQuery = $CamlQueryAssigned  }
    2  { $PSListQuery = $CamlQueryAllRecordsNotCompleted }
    3  { $PSListQuery = $CamlQueryAssignedNotCompleted   }
    0  { $PSListQuery = $CamlQueryAllRecords  }
    10 { $PSListQuery = $CamlQueryNotAssignNewStatus  }
    Default {}
}
$PSListQuery = get-pnplistitem -list $SharepointList -Query $PSListQuery


$Checkbox =0
$PSListQuery







$listitems | select-object -property ID, Vendor

get-pnplistitem -list $SharepointList -Query $NewCalm
get-pnplistitem -list $SharepointList -Query $CalmQuery

$table1 = @()

foreach($listItem in $PSListQuery){  
    $table1 += New-object PsObject -property @{
        ID = $listItem["ID"]
        Vendor          =   $Listitem["Vendor"]
        Appname         =   $Listitem["AppName"]
        Version         =   $Listitem["Version"]
        BrickName       =   $Listitem["BrickName"]
        AppType         =   $Listitem["AppType"]
        'Ad Group'      =   $Listitem["ADGroup"]
        'Assigned to'   =   $Listitem["AssignedTo"]
        Architecture    =   $Listitem["Architecture"]
        target          =   $Listitem["Target"]
        Status          =   $Listitem["Status"]
        Requestor       =   $Listitem["Requestor"]
        Email           =   $Listitem["RequestorEmail"]
        Phone           =   $Listitem["RequestorPhone"]
    }

} 

$table1 > C:\HypervCopy\TEXT.txt



$Listitems | Select-Object -property "ID";"RequestorPhone"
$listItem
$Table1
$Listitem["Version"]
$Listitem["BrickName"]
$Listitem["AppType"]
$Listitem["ADGroup"]
$Listitem["AssignedTo"]
$Listitem["Architecture"]
$Listitem["Target"]
$Listitem["Status"]
$Listitem["Requestor"]
$Listitem["RequestorEmail"]
$Listitem["RequestorPhone"]

$ListName = $SharepointList 
$Fields = Get-PnPField -List $SharepointList   
$global:FieldsList = ""
foreach($Field in $Fields)
{
    $global:FieldsList = $global:FieldsList + $Field.InternalName + ','
}

$Items = (Get-PnPListItem -List $ListName -Fields $global:FieldsList).FieldValues    
foreach($Item in $Items)
{
    Write-Host $Item.Title -ForegroundColor Yellow
    foreach($Field in $Fields)
    {
        Write-Host $Field.InternalName -ForegroundColor Yellow
        Write-Host $Item[$Field.InternalName]
    }
}




###




$checkbox = $true
$PSlistcurrentUser = "Badger.com\dave"


if (!$checkbox) {
    $PSlistAssigned = $null
}
else {
    $PSlistAssigned = $PSlistcurrentUser
    
}

$PSlistAssigned



[System.Net.CredentialCache]$credentials = New-Object -TypeName System.Net.CredentialCache
$ctx.Credentials = $credentials.DefaultNetworkCredentials


(gmo -l *Sharepoint*).path

