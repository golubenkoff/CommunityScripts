param(
    [parameter(Mandatory = $true)] [string]$UserPrincipalName ,
    [parameter(Mandatory = $true)][string]$AppName,
    [parameter(Mandatory = $true)][string]$TenantID,
    [parameter(Mandatory = $true)][string]$ClientID,
    [parameter(Mandatory = $false)][validateset('Cert:\CurrentUser\My','cert:\LocalMachine\My')]$certStoreLocation = 'Cert:\CurrentUser\My'
)

Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Calendar
Import-Module Microsoft.Graph.Users.Actions


if (!(Get-MgContext)) {

    $Cert = $null ; $Cert = Get-ChildItem $certStoreLocation | Where-Object FriendlyName -EQ $AppName

    if ($null -eq $Cert) {
        Write-Host "No Certificate for $AppName found. Aborting..." -ForegroundColor White -BackgroundColor DarkRed
        break;
    } else {
        Connect-MgGraph -ClientId $ClientID -TenantId $TenantID -CertificateThumbprint $Cert.Thumbprint #-Scopes $RequiredScopes
    }
}

Get-MgContext

$Report = New-Object System.Collections.ArrayList

$UserId = Get-MgUser -UserId $UserPrincipalName -ExpandProperty @("Calendar","Calendars","CalendarView","CalendarGroups")


if ($null -ne $userId.Id) {
    Write-Host 'Working with Mailbox: ' $userId.DisplayName `t $userid.Mail

    $Calendars = $null ;  $Calendars = Get-MgUserCalendar -UserId $userId.Id
    Write-Host 'Calendars Found: ' $Calendars.Count

    foreach ($CalendarId in $Calendars){
        Write-Host 'Working with Calendar: ' $CalendarId.Name

        Write-Host 'Searching Events. Please wait....'

        $Events = $null ; $Events = Get-MgUserCalendarEvent -UserId $userId.Id -CalendarId $CalendarId.Id -All

        if ($Events) {

            Write-Host 'Total Events      :' $Events.Count
            Write-Host 'isOrganizer Events:' ($Events | Where-Object IsOrganizer -EQ $true).count

            foreach ($Event in $($Events | Where-Object IsOrganizer -EQ $true)) {
                Write-Progress -CurrentOperation $Event.Subject -Activity "[$UserPrincipalName] - [$($CalendarId.Name)]  Events"

                [void]$Report.Add($($Event  | Select-Object @{N="UserPrincipalName";E={$UserPrincipalName}}, `
                @{N="DisplayName";E={$userid.DisplayName}}, `
                @{N="CalendarName";E={$CalendarId.Name}}, `
                IsOrganizer, `
                CreatedDateTime, `
                HasAttachments, `
                Subject,
                @{N="StartTime";E={(get-date $_.Start.DateTime).ToLocalTime().ToString('dd.MM.yyyy HH:mm')}},
                @{N="EndTime";E={(get-date $_.End.DateTime).ToLocalTime().ToString('dd.MM.yyyy HH:mm')}})
                );
            }
        }
    }

}

if ($Report.Count -gt 0){
    $Report | Export-csv -NoTypeInformation -Encoding utf8 -UseCulture -Path "~\Desktop\Events_$UserPrincipalName_$(get-date -f 'ddMMyyyyHHmm').csv"
}

Disconnect-MgGraph
