param(
$FileName = "C:\Scripts\Pradeep_Source.csv"
)


$Data = import-csv $FileName -Encoding UTF8
$Report = @()
foreach ($user in $Data){
$Name = $null; $Name = $user.Name
$UserID = $null; $UserID = Get-ADUser -Filter {Name -eq $Name} -Properties extensionattribute11,mail,DisplayName
    if ($UserID -and !([string]::IsNullOrEmpty($UserID.extensionattribute11)))
    {
        $Report += $UserID | Select Name,DisplayName,Mail,extensionattribute11
    }
}

$Report | export-csv -NoTypeInformation -Encoding UTF8 -Path $($FileName -replace "\.csv","_Report_$(get-date -f 'ddMMyyyyHHmm').csv")