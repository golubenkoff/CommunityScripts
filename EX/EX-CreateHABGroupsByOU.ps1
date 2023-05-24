<#
	.SYNOPSIS
Create Exchange HAB Groups by OU Structure

	.PARAMETER
    $HabRootName  - Root OU Name where to Start

	.PARAMETER
    $HabGroupsOuName - OU Name where to create HAB Groups

	.OUTPUTS
		Console

	.NOTES
		Take It, Hold It, Love It

	.LINK
		Author : Andrew V. Golubenkoff (andrew.golubenkoff@outlook.com)

#>



<#

    .NOTES
████████╗ █████╗ ██╗  ██╗███████╗    ██╗████████╗
╚══██╔══╝██╔══██╗██║ ██╔╝██╔════╝    ██║╚══██╔══╝
   ██║   ███████║█████╔╝ █████╗      ██║   ██║
   ██║   ██╔══██║██╔═██╗ ██╔══╝      ██║   ██║
   ██║   ██║  ██║██║  ██╗███████╗    ██║   ██║▄█╗
   ╚═╝   ╚═╝  ╚═╝╚═╝  ╚═╝╚══════╝    ╚═╝   ╚═╝╚═╝

██╗  ██╗ ██████╗ ██╗     ██████╗     ██╗████████╗
██║  ██║██╔═══██╗██║     ██╔══██╗    ██║╚══██╔══╝
███████║██║   ██║██║     ██║  ██║    ██║   ██║
██╔══██║██║   ██║██║     ██║  ██║    ██║   ██║
██║  ██║╚██████╔╝███████╗██████╔╝    ██║   ██║▄█╗
╚═╝  ╚═╝ ╚═════╝ ╚══════╝╚═════╝     ╚═╝   ╚═╝╚═╝

██╗      ██████╗ ██╗   ██╗███████╗    ██╗████████╗
██║     ██╔═══██╗██║   ██║██╔════╝    ██║╚══██╔══╝
██║     ██║   ██║██║   ██║█████╗      ██║   ██║
██║     ██║   ██║╚██╗ ██╔╝██╔══╝      ██║   ██║
███████╗╚██████╔╝ ╚████╔╝ ███████╗    ██║   ██║
╚══════╝ ╚═════╝   ╚═══╝  ╚══════╝    ╚═╝   ╚═╝

    .LINK
        Author : Andrew V. Golubenkoff (andrew.golubenkoff@outlook.com)


#>
param(
    $HabRootName = 'HAB_ROOT',
    $HabGroupsOuName = 'HAB_GROUPS',
    [switch]$StartProcessing
)


$DC = $env:LOGONSERVER -replace '\\\\'

#region Functions

function Trim-Length {
    param (
        [parameter(Mandatory = $True,ValueFromPipeline = $True)] [string]$Str
        , [parameter(Mandatory = $True,Position = 1)] [int] $Length
    )
    $Str[0..($Length - 1)] -join ''
}

function TranslitRU2LAT {
    param([string]$inString)
    $Translit = @{
        [char]'а' = 'a';[char]'А' = 'A';
        [char]'б' = 'b';[char]'Б' = 'B';
        [char]'в' = 'v';[char]'В' = 'V';
        [char]'г' = 'g';[char]'Г' = 'G';
        [char]'д' = 'd';[char]'Д' = 'D';
        [char]'е' = 'e';[char]'Е' = 'E';
        [char]'ё' = 'ye';[char]'Ё' = 'Ye';
        [char]'ж' = 'zh';[char]'Ж' = 'Zh';
        [char]'з' = 'z';[char]'З' = 'Z';
        [char]'и' = 'i';[char]'И' = 'I';
        [char]'й' = 'y';[char]'Й' = 'Y';
        [char]'к' = 'k';[char]'К' = 'K';
        [char]'л' = 'l';[char]'Л' = 'L';
        [char]'м' = 'm';[char]'М' = 'M';
        [char]'н' = 'n';[char]'Н' = 'N';
        [char]'о' = 'o';[char]'О' = 'O';
        [char]'п' = 'p';[char]'П' = 'P';
        [char]'р' = 'r';[char]'Р' = 'R';
        [char]'с' = 's';[char]'С' = 'S';
        [char]'т' = 't';[char]'Т' = 'T';
        [char]'у' = 'u';[char]'У' = 'U';
        [char]'ф' = 'f';[char]'Ф' = 'F';
        [char]'х' = 'kh';[char]'Х' = 'Kh';
        [char]'ц' = 'ts';[char]'Ц' = 'Ts';
        [char]'ч' = 'ch';[char]'Ч' = 'Ch';
        [char]'ш' = 'sh';[char]'Ш' = 'Sh';
        [char]'щ' = 'sch';[char]'Щ' = 'Sch';
        [char]'ъ' = '';[char]'Ъ' = '';
        [char]'ы' = 'y';[char]'Ы' = 'Y';
        [char]'ь' = '';[char]'Ь' = '';
        [char]'э' = 'e';[char]'Э' = 'E';
        [char]'ю' = 'yu';[char]'Ю' = 'Yu';
        [char]'я' = 'ya';[char]'Я' = 'Ya';
        [char]'ґ' = 'g';[char]'Ґ' = 'G';
        [char]'Є' = 'Ye';[char]'є' = 'Ye';
        [char]'і' = 'i'; [char]'І' = 'I';
        [char]'ї' = 'yi';[char]'Ї' = 'Yi'
        [char]'№' = '#';[char]'-' = '_';
        [char]'„' = '';[char]'”' = '';
        [char]'"' = '';[char]'–' = '_';
        [char]'`' = '';[char]"’" = '' # Апостроф должен быть в Двойных Кавычках
    }
    $outChars = ''
    $Result = New-Object PsCustomObject
    $Result | Add-Member -MemberType NoteProperty -Name Text -Value ''
    #$inString = $InputName.Text +"."+ $InputSName.Text
    foreach ($c in $inChars = $inString.ToCharArray()) {
        if ($Translit[$c] -cne $Null )
        { $outChars += $Translit[$c] }
        else
        { $outChars += $c }
    }
    $Result.Text = $outChars
    return $Result.Text
}

Function GetParent {
    param(
        [parameter(Mandatory)][string]$ouDistinguishedName
    )
    $ouComponents = $ouDistinguishedName -split ','
    return $(($ouComponents[1..($ouComponents.Count - 1)]) -join ',')
}

Function GetHabName {
    param([parameter(Mandatory)]$OU)

    return [PSCustomObject]@{
        DisplayName = $(Trim-Length -Str $( $('# ' + $($OU.Description)) ) -Length 250) -replace '^\s' -replace '\s$'
        Name        = $("HAB_$($OU.uSnCreated)_" + $(TranslitRU2LAT -inString $(Trim-Length -Str (($($OU.Name)) -replace ",",'_' -replace "'" -replace '\(' -replace '\)' -replace '"' -replace '”' -replace '„' ) -Length 44 -ErrorAction Stop) ) ) -replace '\s','_' -replace '\.$'
    }

}

Function CreateHabGroup {
    param([parameter(Mandatory)]$GroupObj)
    $ParentName = GetParent($GroupObj.DistinguishedName)
    $ParentOU = Get-ADOrganizationalUnit $ParentName -Properties Description,uSnCreated
    $ParentGroup = GetHabName $ParentOU

    $nameObj = $null ; $NameObj = GetHabName $OU
    Write-Host $GroupObj.Name `t "[$($nameObj.Name)] - [$($nameObj.displayName)]" `t $GroupObj.uSnCreated `t $ParentName

    $membersHAB = $null ; $membersHAB = try {
        Get-ADUser -SearchBase $GroupObj.DistinguishedName -SearchScope OneLevel -Server $DC -LDAPFilter '(&(objectClass=user)(!(useraccountcontrol:1.2.840.113556.1.4.803:=2))(Mail=*@*))' -ErrorAction Stop
    } catch {
        Write-Host 'Cannot Get Members:' $GroupObj.DistinguishedName $_.exception.messages
    }


    try {
        $CheckGroupID = Get-DistributionGroup $nameObj.Name -DomainController $DC -ErrorAction stop -Verbose
    } catch {
        try {
            if ($null -ne $membersHAB -and $StartProcessing.IsPresent ) {
                Write-Host `t`t 'Creating Group: ' -ForegroundColor Gray -NoNewline ; Write-Host $nameObj.Name -ForegroundColor Green
                $CheckGroupID = New-DistributionGroup -name $nameObj.Name -DisplayName $nameObj.DisplayName -Alias $nameObj.Name -OrganizationalUnit $groupsDN.DistinguishedName -DomainController $DC -ErrorAction Stop
                Set-Group $nameObj.Name -IsHierarchicalGroup:$true -DomainController $DC -WarningAction SilentlyContinue -ErrorAction Stop -BypassSecurityGroupManagerCheck
                Set-DistributionGroup $nameObj.Name -MaxSendSize 10MB -MaxReceiveSize 10MB -MailTip 'Внимание! Максимальный Размер сообщения 10Мб' -DomainController $DC
            }elseif($null -ne $membersHAB -and !($StartProcessing.IsPresent)){
                Write-Host `t`t '[Check Only ]Creating Group: ' -ForegroundColor Gray -NoNewline ; Write-Host $nameObj.Name -ForegroundColor Green
            }else{
                Write-Host `t`t '[Skipping ] Creating Group - No Members: ' -ForegroundColor Gray -NoNewline ; Write-Host $nameObj.Name -ForegroundColor Green
            }
        } catch { Write-Host "Cannot Create Group: $($nameObj.Name) $($_.Exception.ItemName) $($_.Exception.Message)" }
    } finally {

        if ($null -ne $CheckGroupID) {
            # ADD to Parent
            if ($GroupObj.DistinguishedName -ne $rootDN.DistinguishedName) {

                if (!( (Get-DistributionGroupMember $ParentGroup.Name -DomainController $DC).Name -ccontains $CheckGroupID.Name) -and $StartProcessing.IsPresent) {
                    Add-DistributionGroupMember $ParentGroup.Name -Member $CheckGroupID.DistinguishedName
                }
            }

            # Add Users to Group

            $CurrentGroupID = $null ; $CurrentGroupID = try {
                Get-ADGroup $CheckGroupID.DistinguishedName -Properties member -ErrorAction Stop
            } catch {
                Write-Host 'Cannot Find Group:' $CheckGroupID.DistinguishedName $_.exception.messages
            }

            if ($null -ne $membersHAB) {
                $membersHABToAdd = $null ;  $membersHABToAdd = $membersHAB | Where-Object DistinguishedName -NotIn $CurrentGroupID.member

                if ($membersHABToAdd -and $CurrentGroupID -and $StartProcessing.IsPresent) {
                    try {
                        Add-ADGroupMember $CurrentGroupID -Members $membersHABToAdd -ErrorAction Stop
                    } catch {
                        Write-Host "Cannot Add HAB Members [$($CurrentGroupID.Name)]: "$_.exception.Message
                    }
                }elseif($membersHABToAdd -and $CurrentGroupID -and !($StartProcessing.IsPresent)){
                    Write-Host `t`t '[Check Only ]Adding Members to Group: ' -ForegroundColor Gray -NoNewline ; Write-Host $nameObj.Name -ForegroundColor Green
                }
            }
        }
    }
}

#endregion Functions

$groupsDN = $null ; $groupsDN = Get-ADOrganizationalUnit -Filter { Name -eq $HabGroupsOuName }
$rootDN = $null ; $rootDN = Get-ADOrganizationalUnit -Filter { Name -eq $HabRootName } -Properties Description,uSnCreated


if ($null -ne $rootDN -and $null -ne $groupsDN) {

    $OU_TREE = $null ; $OU_TREE = try {
        Get-ADOrganizationalUnit -SearchBase $rootDn.DistinguishedName -Properties Description,uSnCreated -LDAPFilter '(&(ObjectClass=organizationalUnit)(Description=*))' -ErrorAction Stop
    } catch {
        Write-Host 'Cannot Create OU Tree: ' $_.exception.message -BackgroundColor Red;
        break
    }

    if ($null -ne $OU_TREE) {

        $c = 0 ; $max = $OU_TREE.Count
        foreach ($OU in $OU_TREE) {
            $c++
            $percentCompleted = $c * 100 / $max
            $message = '{0:p1} done, processing {1}' -f ( $percentCompleted / 100), $OU.Name
            Write-Progress -Activity 'Processing' -PercentComplete $($c * 100 / $max) -Status $message

            CreateHabGroup($OU)

        }


    }
} else {
    throw 'No root OU for Users or Groups. Please check Names'
}

# DONE