




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



#### vm commands
Start-VM $VMname1
New-PSSession -VMName $VMname1 -Credential $credential
Copy-Item -ToSession $s -path C:\HypervCopy -Destination C:\HypervCopy
Start-vm $VMname1 -checkpoint
Stop-vm $VMname1

get-vm | Checkpoint-VM -SnapshotName "$vmname1`_$(get-date -format 'dd-MM-yyyy-hh-mm')"
Restore-VMCheckpoint -Name "Base" -VMName $VMname1 -Confirm



$username = "bad.admin@acme.com.au"
$password = "Ilizna234!"
$secureStringPwd = $password | ConvertTo-SecureString -AsPlainText -Force 
$secureStringText = "Ilizn234!"
$secureStringText = $secureStringPwd | ConvertFrom-SecureString 
Set-Content "C:\temp\ExportedPassword.txt" $secureStringText

$credential = Get-Credential
$credential.Password | ConvertFrom-SecureString | Set-Content C:\HypervCopy\encrypted_password.txt
$User1 = "Badger\Dave"
$encrypted = Get-Content C:\HypervCopy\encrypted_password.txt | ConvertTo-SecureString
$credential = New-Object System.Management.Automation.PsCredential($user1, $encrypted)


get-vm Sharepointserver1 | Select id

$S = New-PSSession -VMId $VMID1 -Credential $credential
Copy-Item -ToSession $s -Path C:\HypervCopy\Source -Recurse -Destination C:\HypervCopy -Force
Remove-PSSession $s