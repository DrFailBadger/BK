#requires -version 3.0

function Add-RegKeyMember {

 

    [CmdletBinding()]

    param(

    [Parameter(Mandatory, ParameterSetName="ByKey", Position=0, ValueFromPipeline)]

    # Registry key object returned from Get-ChildItem or Get-Item

    [Microsoft.Win32.RegistryKey] $RegistryKey,

    [Parameter(Mandatory, ParameterSetName="ByPath", Position=0)]

    # Path to a registry key

    [string] $Path

    )



    begin {

    }



    process {

    }

}


dir HKLM:\SOFTWARE -Recurse | Add-RegKeyMember | Select-Object name, lastwritetime
