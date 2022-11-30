param(
    [Parameter(Mandatory = $false)]
    [string]
    $filePath
)

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

#region Functions
function Read-OpenFileDialog([string]$WindowTitle, [string]$InitialDirectory, [string]$Filter = 'All files (*.*)|*.*', [switch]$AllowMultiSelect) {
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = $WindowTitle
    if (![string]::IsNullOrWhiteSpace($InitialDirectory)) { $openFileDialog.InitialDirectory = $InitialDirectory }
    $openFileDialog.Filter = $Filter
    if ($AllowMultiSelect) { $openFileDialog.MultiSelect = $true }
    $openFileDialog.ShowHelp = $true
    $openFileDialog.ShowDialog() > $null
    if ($AllowMultiSelect) { return $openFileDialog.Filenames } else { return $openFileDialog.Filename }
}

Function Get-GroupMemberRaw {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true,ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [Alias('DN','DistinguishedName')]
        $Name,
        [uint32]$HighRange,
        [uint32]$LowRange = 0,
        [switch]$CountMembers,
        [string]$Server
    )

    begin {
        $RangeStep = 999
        $Filter = '(&(objectClass=Group)(objectCategory=Group))'
    }

    process {
        $LowRange = [int]$PSBoundParameters['LowRange']

        if ($LowRange -gt 0) { $LowRange -= 1 }
        if ($HighRange -gt 0) { $HighRange -= 1 }

        $Members = @()

        if ($Server) {
            $AdsPath = "LDAP://$Server/$Name"
        } else {
            $AdsPath = "LDAP://$Name"
        }

        $IsExists = $false

        if ([DirectoryServices.DirectoryEntry]::Exists($AdsPath)) {
            if ($HighRange) {
                while ($LowRange -lt $HighRange) {
                    $MiddleRange = $LowRange + $RangeStep - 1
                    if ($MiddleRange -gt $HighRange) {
                        $Properties = "member;range=$LowRange-$HighRange"
                    } else {
                        $Properties = "member;range=$LowRange-$MiddleRange"
                    }

                    $Searcher = New-Object DirectoryServices.DirectorySearcher(
                        $AdsPath, $Filter, $Properties, 'Base'
                    )

                    try {
                        $Group = $Searcher.FindOne().Properties

                        $Attribute = ($Group.GetEnumerator() |
                                Where-Object { $_.Name -match 'member' }).Name

                        Write-Verbose "$Name - $Attribute"
                        $Members += $Group[$Attribute]
                    }

                    catch {
                        break
                    }

                    $LowRange += $RangeStep
                }
            }

            else {
                while ($true) {
                    $HighRange = $LowRange + $RangeStep - 1
                    $Properties = "member;range=$LowRange-$HighRange"

                    $Searcher = New-Object DirectoryServices.DirectorySearcher(
                        $AdsPath, $Filter, $Properties, 'Base'
                    )

                    try {
                        $Group = $Searcher.FindOne().Properties

                        $Attribute = ($Group.GetEnumerator() |
                                Where-Object { $_.Name -match 'member' }).Name

                        Write-Verbose "$Name - $Attribute"
                        $Members += $Group[$Attribute]
                    }

                    catch {
                        break
                    }

                    $LowRange += $RangeStep
                }
            }

            $IsExists = $true
        }

        else {
            Write-Host "The path $Name is invalid" -ForegroundColor Yellow
        }

        if ($IsExists) {
            if ($CountMembers) {
                $Members.Count
            }

            else {
                $Members
            }
        }
    }
}

Function Get-GroupMemberFast {
    [cmdletbinding()]
    Param($distinguishedname,[switch]$Recursive)
    Write-Verbose "DN: $($distinguishedname)"
    Get-GroupMemberRaw -Name $distinguishedname | ForEach-Object {
        $c = $(New-Object System.DirectoryServices.DirectoryEntry('LDAP://' + $_)) ;
        if ($c.Properties.objectcategory -match 'group' -and $Recursive.IsPresent) {
            Get-GroupMemberFast $c.properties.distinguishedname -Recursive
        } else { $c }
    }
}
#endregion Functions


if ([string]::IsNullOrEmpty($filePath)) {
    $filePath = Read-OpenFileDialog -WindowTitle 'Select Source Txt File' -InitialDirectory $scriptPath -Filter 'Text Files (*.txt)|*.txt'
    if (![string]::IsNullOrEmpty($filePath)) { Write-Host "You selected the file: $filePath" -ForegroundColor Cyan }
    else { Write-Host 'You did not select a Text [txt] file.' -ForegroundColor White -BackgroundColor Red; break }
}

$SourceData = $null ; $SourceData = Get-Content $filePath | ForEach-Object { $_.Trim() }

if ($SourceData) {

    foreach ($Group in $SourceData) {
        Write-Host 'Processing: ' $Group

        $GroupADSPath = $null ; $GroupADSPath = ([adsisearcher]"(&(objectClass=group)(name=$Group))").FindOne()

        if ($GroupADSPath.Path) {

            $AdGroupMembers = $null ;    $AdGroupMembers = Get-GroupMemberFast -distinguishedname $GroupADSPath.Properties['distinguishedname'] -Recursive

            if ($AdGroupMembers) {

                $members = $null ; $members = $AdGroupMembers | Select-Object @{N = 'GroupName';E = { $Group } },@{N = 'MemberName';E = { $_.Properties['name'][0] } },@{N = 'MemberDisplayName';E = { $_.Properties['displayname'][0] } },@{N = 'Mail';E = { $_.Properties['mail'][0] } }

                if ($members) {
                    $members | Export-Csv "$($($filePath -replace '.txt',"_Results_$(Get-Date -f 'ddMMyyyyHHmm').csv"))" -UseCulture -Encoding UTF8 -Append -NoTypeInformation
                }
            }
        }

    }
}
