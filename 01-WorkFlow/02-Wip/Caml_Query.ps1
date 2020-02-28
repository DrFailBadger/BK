###  set


$Script:strSharePointSiteURL = "https://uam-acc.ms.myatos.net/"
$Returnval =@()

[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($Script:strSharePointSiteURL)
[System.Net.CredentialCache]$credentials = New-Object -TypeName System.Net.CredentialCache
$ctx.Credentials = $credentials.DefaultNetworkCredentials;
$ctx.RequestTimeOut = 5000 * 60 * 10
$web = $ctx.Web
$list = $web.Lists.GetByTitle("Config - Analyst groups")
$camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
$camlQuery.ViewXml = "<View>
<Query>
<Where>
<Eq>
<FieldRef Name='Title' />
<Value Type='Text'>$GBUanalystgroup</Value>
</Eq>
</Where>
</Query>
<RowLimit>10000</RowLimit>
</View>"
$spListItemCollection = $List.GetItems($camlQuery)
$ctx.Load($spListItemCollection)
$timeone = Get-Date
$ctx.ExecuteQuery()



foreach ($item in $spListItemCollection) { 
    $item["Title"] = "$Reason"
Try{
    $Item.Update()
    $ctx.ExecuteQuery()
}Catch{
$success = $false
}
 
    }

###return
$Script:strSharePointSiteURL = "https://uam-acc.ms.myatos.net/"
$Returnval =@()

[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($Script:strSharePointSiteURL)
[System.Net.CredentialCache]$credentials = New-Object -TypeName System.Net.CredentialCache
$ctx.Credentials = $credentials.DefaultNetworkCredentials;
$ctx.RequestTimeOut = 5000 * 60 * 10
$web = $ctx.Web
$list = $web.Lists.GetByTitle("Config - Analyst groups")
$camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
$camlQuery.ViewXml = "<View>
<Query>
<Where>
<Eq>
<FieldRef Name='Title' />
<Value Type='Text'>$GBUanalystgroup</Value>
</Eq>
</Where>
</Query>
<RowLimit>10000</RowLimit>
</View>"
$spListItemCollection = $List.GetItems($camlQuery)
$ctx.Load($spListItemCollection)
$timeone = Get-Date
$ctx.ExecuteQuery()

foreach ($item in $spListItemCollection){
$1 = $item['ServiceHoursFrom'] 
$2 = $item['ServiceHoursTo'] 
}
$Returnval = @()



foreach ($item in $spListItemCollection){

$Returnval += New-Object PsObject -Property @{
Title = $item['Title'] 
Version = $item['Version'] 

}
}

$Returnval |where -Property Title -eq "Dong"