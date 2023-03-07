<#
.SYNOPSIS
Retrieves information for Azure AD users whose display name starts with "Train." and exports it to a table.

.DESCRIPTION
This script retrieves information for Azure AD users whose display name starts with "Train.", including their user principal name, email address, last logon date, group membership, licenses assigned, mailbox size, and OneDrive size. The information is exported to a table.

.PARAMETER None
This script has no parameters.

.EXAMPLE
.\Get-TrainUsersReport.ps1
This command retrieves information for Azure AD users whose display name starts with "Train." and exports it to a table.

.NOTES
This script requires the AzureAD, ExchangeOnlineManagement, and Microsoft.Online.SharePoint.PowerShell modules. If any of these modules are missing, the script will exit and display a warning message.

References:

AzureAD module documentation: https://docs.microsoft.com/en-us/powershell/module/azuread/

Exchange Online Management module documentation: https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2/exchange-online-powershell-v2?view=exchange-ps

SharePoint Online Management Shell module documentation: https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online?view=sharepoint-ps

#>

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Check if required modules are installed and import them if necessary
$requiredModules = @('AzureADPreview', 'ExchangeOnlineManagement', 'Microsoft.Online.SharePoint.PowerShell')
$missingModules = $requiredModules | Where-Object { -not (Get-Module -ListAvailable -Name $_) }
if ($missingModules.Count -gt 0) {
    Write-Warning "The following required PowerShell modules are missing: $($missingModules -join ', '). Please install them before running this script."
    exit
} else {
    foreach ($module in $requiredModules) {
        Import-Module -Name $module -ErrorAction Stop
    }
}

# Connect to Azure Active Directory
Connect-AzureAD -ErrorAction Stop
$AllSKU = @{}
Get-AzureADSubscribedSku | ForEach-Object { $AllSKU.Add($_.SkuId,$_) }

# Connect to Exchange Online
Connect-ExchangeOnline -ErrorAction Stop

# Get SharePoint Online URL from Azure AD tenant
$tenant = Get-AzureADTenantDetail
$sharePointUrl = "https://$($($tenant.VerifiedDomains | Where-Object { $_.Initial -eq $true }).Name -replace '\..*')-admin.sharepoint.com"

# Connect to SharePoint Online
Connect-SPOService -Url $sharePointUrl -ErrorAction Stop
$Sites = Get-SPOSite -IncludePersonalSite $true -limit All

# Define variables
$users = Get-AzureADUser -All $true -Filter "startswith(DisplayName,'Train.')"

$results = @()

# Loop through each user and retrieve necessary information
foreach ($user in $users) {

    $userPrincipalName = $null ; $userPrincipalName = $user.UserPrincipalName
    $mail = $null ; $mail = $user.Mail

    # Get last logon date
    $AzureADAuditSignInLogs = $null ; $AzureADAuditSignInLogs = Get-AzureADAuditSignInLogs -Filter "startsWith(userPrincipalName,'$($user.UserPrincipalName)')" -Top 1

    $lastLogon = $lastLogonDate = $null ; $lastLogon = [datetime]::Parse($AzureADAuditSignInLogs.CreatedDateTime)

    if ($lastLogon) {
        $lastLogonDate = $lastLogon.ToString('dd.MM.yyyy HH:mm:ss')
    } else {
        $lastLogonDate = 'Never logged in'
    }

    $groups = $null ; $groups = (Get-AzureADUserMembership -ObjectId $user.ObjectId).DisplayName -join "`n"

    $licenses = $null ; $licenses = ($user.AssignedLicenses | ForEach-Object {
            $AllSku[$_.SkuId].SkuPartNumber
        }) -join "`n"

    # Get mailbox size
    try {
        $mailboxSize = [math]::Round(($((Get-MailboxStatistics -Identity $user.UserPrincipalName -ErrorAction Stop | Select-Object -ExpandProperty TotalItemSize).Value).ToString().Split('(')[1].Split(' ')[0].Replace(',','') / 1MB),2)
    } catch {
        $mailboxSize = 'Error retrieving mailbox size'
    }

    # Get OneDrive size
    try {
        $oneDrive = $null ; $oneDrive = $Sites | ? Owner -eq $user.UserPrincipalName | ? Template -eq 'SPSPERS#10'
        $oneDriveSize = $null ; $oneDriveSize = $oneDrive | Select-Object -ExpandProperty StorageUsageCurrent
    } catch {
        $oneDriveSize = 'Error retrieving OneDrive size'
    }

    # Create a custom object and add it to the results array
    $result = [PSCustomObject]@{
        UserPrincipalName = $userPrincipalName
        Mail              = $mail
        LastLogonDate     = $lastLogonDate
        Groups            = $groups
        Licenses          = $licenses
        MailboxSize       = $mailboxSize
        OneDriveSize      = $oneDriveSize
    }
    $results += $result
}


$results | Export-Csv $(join-path $scriptPath "TrainReport_$(get-date -f 'ddMMyyyyHHmm').csv") -NoTypeInformation -Encoding utf8 -UseCulture

