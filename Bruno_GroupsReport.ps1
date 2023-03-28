param(
    [parameter(Mandatory = $true,HelpMessage = 'Organizational Unit Name')]$OU
)

#requires -module ActiveDirectory

$ScriptPath = Split-Path $MyInvocation.MyCommand.Definition

[System.Collections.ArrayList]$Report = @()

if (!([string]::IsNullOrEmpty($OU))) {
    $OU_ID = $null ; $OU_ID = Get-ADOrganizationalUnit -Filter { Name -eq $OU }
    if ($OU_ID) {
            (Get-ADGroup -SearchBase $OU_ID -SearchScope Subtree -Filter *) | ForEach-Object {
            $GroupName = $null ; $GroupName = $_
            $GroupName | Get-ADGroupMember | ForEach-Object {
                [void]$Report.Add([PSCustomObject]@{
                        GroupName            = $GroupName.Name
                        MemberSamAccountName = $_.SamAccountName
                        MemberObjectClass    = $_.ObjectClass
                    })
            }
        }
    }
}

if ($Report) {
    $Report | Export-Csv -Path $(Join-Path $ScriptPath "Report_OU_${OU}_$(Get-Date -f 'ddMMyyyyHHmm').csv") -UseCulture -Encoding UTF8 -NoTypeInformation
}