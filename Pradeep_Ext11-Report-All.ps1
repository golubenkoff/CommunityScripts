param(
$FileName = "C:\Scripts\Pradeep_Source.csv"
)

$Report = import-csv $FileName -Encoding UTF8 | %{ $Name = $null ;  $Name = $_.Name
Get-ADUser -Filter {Name -eq $Name} -Properties extensionattribute11,mail,DisplayName | Select Name,DisplayName,Mail,extensionattribute11
}  | export-csv -NoTypeInformation -Encoding UTF8 -Path $($FileName -replace "\.csv","_Report_$(get-date -f 'ddMMyyyyHHmm').csv")
