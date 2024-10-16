$csvData = Import-Csv -Path 'C:\temp\groups.csv' -UseCulture

$groupCreated = @{}

foreach ($row in $csvData) {
    $dlName = $row.DLName
    $owner = $row.Owner
    $member = $row.Member

    if (-not $groupCreated[$dlName]) {
        try {
            New-DistributionGroup -Name $dlName -Alias $dlName -PrimarySmtpAddress "$dlName@yourdomain.com" -ManagedBy $owner -Members $member -ErrorAction Stop
            Write-Host "Created Distribution Group: $dlName with owner $owner and member $member"
        } catch {
            Write-Host "Error creating Distribution Group: $dlName. $_"
        }

        try {
            Set-DistributionGroup -Identity $dlName -ManagedBy @{Add = $owner }
            Write-Host "Added owner $owner to group $dlName"
        } catch {
            Write-Host "Error adding owner $owner to $dlName. $_"
        }

        $groupCreated[$dlName] = $true
    }

    try {
        Add-DistributionGroupMember -Identity $dlName -Member $member
        Write-Host "Added member $member to group $dlName"
    } catch {
        Write-Host "Error adding member $member to $dlName. $_"
    }
}

<# CSV Example
DLName,Owner,Member
DLNAMEONE,Bruno,John
DLNAMEONE,Andrew,David
DLTWONAME,Bruno,Mathew
DLNAMETHREE,Andrew,Bruno
DLNAMETHREE,John,Mathew
#>