param(
[Parameter(Mandatory=$false)][string]$InFile = "C:\Scripts\Pradeep_ADGroups_SrcData.csv"
)

try{
$SourceData = $null ; $SourceData = import-csv $InFile -ErrorAction stop -UseCulture
}catch {Write-Host "Cannot Read source file: $InFile :  " $_.Exception.Message -ForegroundColor Red}

$DC = $env:LOGONSERVER -replace "\\\\"
if ($SourceData)
{

    foreach ($group in $SourceData)
    {
        try{
            New-ADGroup -Name $group.Name -SamAccountName $group.Name -GroupCategory Security -GroupScope Global -DisplayName $group.DisplayName -Path $group.OU -Description $group.Description -Server $DC -ErrorAction stop
            }catch {Write-Host "Cannot Create Group: $group :  " $_.Exception.Message}

        try{
            Set-ADGroup $group.Name -Add @{extensionAttribute1=$group.extensionAttribute1} -Server $DC -ErrorAction stop
            }catch {Write-Host "Cannot Set Attribute Value: $group :  " $_.Exception.Message -ForegroundColor red}

    }
}
