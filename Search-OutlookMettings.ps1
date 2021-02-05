<#
	.SYNOPSIS
	    Search User Mailboxes for Skype or Teams Meetings

	.DESCRIPTION
        Requires EWS DLL to be installed

	.PARAMETER    
		Start  - Start DateTime for Meetings

	.PARAMETER     
		End - End DateTime for Meetings

	.PARAMETER  
	    ImpersonateUserMail - User account with Impersonate AccessRighs

	.PARAMETER  
	    List - String array with  PrimarySmtpAddress for mailboxes.
        if not set - will search for all Enabled ActiveDirectory users with Mail attribute set

	.PARAMETER  
	    ReportPath - Output folder

	.EXAMPLE
		PS C:\> .\Search-OutlookMeetings.ps1 -list 'user@contoso.com'

	.OUTPUTS
		File in -ReportPath Directory

	.NOTES
		Take It, Hold It, Love It

	.LINK
		Author : Andrew V. Golubenkoff (andrew.golubenkoff@outlook.com)

	.LINK
		

#>

param(
#region params
[Parameter(Mandatory=$false)][string]$ImpersonateUserMail = "admin@contoso.com",         # Admin Account with Impersonate Rights
[Parameter(Mandatory=$false)][string[]]$List = @("Vasya.Pupkin@contoso.com"),      # List of Primary Emails for Mailboxes

[Parameter(Mandatory=$false)][datetime]$DateTimePicker_EndDate   = (Get-Date).Date,
[Parameter(Mandatory=$false)][datetime]$DateTimePicker_StartDate = (Get-Date).AddDays(-7).Date, # Start <= MAX 2 years => End
[Parameter(Mandatory=$false)][string]$ReportPath = "C:\Scripts"
#endregion params
)

if (!($List)){
$List = get-aduser -Filter {Enabled -eq $true -and Mail -like "*@*"} -Properties mail | Select -ExpandProperty Mail
}


#region EWS DLL
$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
		$EWSDLL = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
        if (Test-Path $EWSDLL)
		    {
		    Import-Module $EWSDLL
            #Add-Type -AssemblyName PresentationFramework
		    }
		else
		    {
		    "$(get-date -format yyyyMMddHHmmss):"
		    "This script requires the EWS Managed API 1.2 or later."
		    "Please download and install the current version of the EWS Managed API from"
		    "http://go.microsoft.com/fwlink/?LinkId=255472"
		    ""
		    "Exiting Script."
		    $exception = New-Object System.Exception ("Managed Api missing")
			throw $exception
		    } 

#endregion EWS DLL

$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016

#region Event Handlers



$Script:service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
if (!($Script:service))
{
Write-Host 'Cannot Create ExchangeService. ' 
break
}
$service.AutodiscoverUrl($ImpersonateUserMail)
$service.UseDefaultCredentials = $true

$Report = @()
$c = 0

[regex]$URL = @"
(?i)\b((?:[a-z][\w-]+:(?:\/{1,3}|[a-z0-9%])|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}\/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'".,<>?«»“”‘’]))
"@

$Props = new-object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$Props.add([Microsoft.Exchange.WebServices.Data.ItemSchema]::TextBody)

foreach($Mail in $List)
{
$c ++
Write-Progress -Activity "Processing Messages" -CurrentOperation $mail -PercentComplete $($c*100/$($List.count))

$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Mail);

$CalendarFolder= New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$null)
$Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$CalendarFolder)

#region Find Meetings

    $appointments  = $null
    if($Calendar.TotalCount -gt 0)
    {
    if ($Calendar.TotalCount -le 1000){
    $cView = New-Object  Microsoft.Exchange.WebServices.Data.CalendarView($DateTimePicker_StartDate, $DateTimePicker_EndDate, $Calendar.TotalCount);

        try{
        $appointments  = $calendar.FindAppointments($cView);
        }catch{
        Write-Host "Search Failed. `n $($_.Exception.Message)"
        }
    }
    else{
        
    $appointments = @()
    
    $Start = $DateTimePicker_StartDate
    $cView = New-Object  Microsoft.Exchange.WebServices.Data.CalendarView($Start, $DateTimePicker_EndDate, 1000);
    
        try{
            $app_offset = $calendar.FindAppointments($cView);
            }catch{
            }

    $appointments += ($app_offset)
    while($app_offset.MoreAvailable)
    {
    $start = ($app_offset.Items | Select-Object -Last 1).Start
    $cView.StartDate = $start
        try{
        $app_offset = $calendar.FindAppointments($cView);
        }catch{
        }
        $appointments += ($app_offset)
        
    }
}
    }
#endregion Find Meetings

    if ($appointments)
    {
    
        foreach ($a in $($appointments | ? IsOnlineMeeting -eq $true | ? {($_.JoinOnlineMeetingUrl -ne $null) -or ($_.Location -like "*Microsoft Teams*")}))
        {
    
        $a.Load($Props)
        $URLS = $null;$URLs = ($URL.Matches($a.TextBody.Text).value | ? {$_ -like "*meetup-join*"}) -join "`n"

            $Report += [PsCustomObject]@{
            Organizer = $a.Organizer -replace "<.*"
            Start = get-date $a.Start -f 'dd.MM.yyyy HH.mm'
            End = get-date $a.End -f 'dd.MM.yyyy HH.mm'
            Participants = ($a.DisplayTo -split ";").count + ($a.DisplayCC -split ";").count
            Subject = $a.Subject
            Size = $a.Size
            JoinUrlSkypeFB = $a.JoinOnlineMeetingUrl
            JoinUrlMsTeams = $URLS
            }
        }

    }

}



if($Report){
$Report | Export-csv $(Join-path $ReportPath "Outlook_meetings_$(get-date -Format 'ddMMyyyyHHmm').csv") -NoTypeInformation -Encoding UTF8 -Delimiter ";"
}


