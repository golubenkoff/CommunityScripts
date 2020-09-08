<#
	.SYNOPSIS
        Script will user UserName or inFile with names and will find user in  AD and show SamAccountName and ExtentionAttribute1

	.DESCRIPTION
		man .\get-user-example.ps1

    .PARAMETER  UserName
		ActiveDirectory username 

    .PARAMETER  inFile
		text file with usernames - one name by line

    .PARAMETER  export
		switch -export  -> export results to csv file

	.EXAMPLE
        .\get-user-example.ps1 -UserName agolubenkov

	.EXAMPLE
        .\get-user-example.ps1 -inFile C:\scripts\myusers.txt

	.OUTPUTS
		CSV File on Desktop

	.NOTES
		Take It, Hold It, Love It

	.LINK
		Author : Andrew V. Golubenkoff (https://www.linkedin.com/in/golubenkoff/)

	.LINK
		GitHub : https://github.com/golubenkoff/CommunityScripts

    .VERSION 1  08-09-2020
#>



param(
[parameter(mandatory=$false, ParameterSetName = "User")][string]$UserName,
[parameter(mandatory=$false, ParameterSetName = "File")][string]$inFile,
[parameter(mandatory=$false)][switch]$Export

)


if ($UserName){
$Result = $null
try{
$Result = ([adsisearcher]"(&(objectCategory=person)(objectClass=user)(extensionAttribute1=*)(Name=${UserName}))").FindOne() | Select @{N="SamAccountName";E={$_.Properties['SamAccountName']}},@{N="extensionAttribute1";E={$_.Properties['extensionAttribute1']}}
}
catch{Write-Host "Error: " $_.Exception.Message -ForegroundColor Red}

}

if ($inFile){
[array]$Result = @()
    foreach($UserName in $(gc $inFile -ErrorAction Stop))
    {
        if (!([string]::IsNullOrEmpty($UserName))){
        $Result += ([adsisearcher]"(&(objectCategory=person)(objectClass=user)(extensionAttribute1=*)(Name=${UserName}))").FindOne() | Select @{N="SamAccountName";E={$_.Properties['SamAccountName']}},@{N="extensionAttribute1";E={$_.Properties['extensionAttribute1']}}

        }

    }

}




if($Export -and $Result)
{
$FileName = "~\Desktop\AD_Ext_Attribute_Report_$(get-date -f 'ddMMyyyyHHnmm').csv"
$Result | export-csv -NoTypeInformation -Encoding UTF8 -Path $FileName -Delimiter ";" -ErrorAction Stop
Write-host "Exported to FileName: " $FileName
}elseif($Result){
$Result 
}else{Write-host "no results..."}


