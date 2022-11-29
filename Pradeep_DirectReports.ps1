param(
    [Parameter(Mandatory)] [string] $UserName,
    [Parameter(Mandatory = $false)] [switch] $Resolve
)

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

Function GetReportees {
    param(
        [string]$UserName,
        [string]$UserDN,
        [string]$ParentUserName
    )

    $Report = New-Object System.Collections.ArrayList

    $UserID = $null
    if ($UserName) {
        $UserID = ([adsisearcher]"(&(objectClass=user)(|(name=$UserName)(samaccountname=$UserName)))").FindOne()
    } elseif ($UserDN) {
        $UserID = [adsi]"LDAP://$UserDN"
    }

    if (!([string]::IsNullOrEmpty($UserID.Path))) {

        if (($UserID.Properties['directreports'])) {
            Write-Host `t"[$($UserID.Properties['name'])]" `t $DR -ForegroundColor Yellow
            foreach ($DR in $($UserID.Properties['directreports'])) {
                [void]$Report.Add(
                    [PsCustomObject]@{
                        UserName   = $Userid.Properties['name'][0]
                        ParentUser = $ParentUserName
                        Reportee   = $DR
                    })

                GetReportees -UserDN $DR -ParentUserName $Userid.Properties['name'][0] | ForEach-Object {
                    [void]$Report.Add($_)
                }
            }

        }

        return $Report

    } else {
        return $null
    }
}

$Report = $null ; $Report = GetReportees -UserName $UserName

if ($Resolve.IsPresent) {
    foreach ($item in $Report) {
        $Resolved = $null ; $Resolved = [adsi]"LDAP://$($item.Reportee)"
        if ($Resolved.Path) {
            $item | Add-Member -MemberType NoteProperty -Name DisplayName -Value $($Resolved.Properties['displayname'][0])
            $item | Add-Member -MemberType NoteProperty -Name Department -Value $($Resolved.Properties['Department'][0])
            $item | Add-Member -MemberType NoteProperty -Name Title -Value $($Resolved.Properties['Title'][0])
            $item | Add-Member -MemberType NoteProperty -Name Description -Value $($Resolved.Properties['Description'][0])
        } else {
            $item | Add-Member -MemberType NoteProperty -Name DisplayName -Value $('')
            $item | Add-Member -MemberType NoteProperty -Name Department -Value $('')
            $item | Add-Member -MemberType NoteProperty -Name Title -Value $('')
            $item | Add-Member -MemberType NoteProperty -Name Description -Value $('')
        }
    }
}

$Report | Sort-Object UserName | Export-Csv $(Join-Path $scriptPath "Report_$($UserName)_$(Get-Date -f 'ddMMyyyyHHmm').csv") -NoTypeInformation -Encoding utf8 -UseCulture

