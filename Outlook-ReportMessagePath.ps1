<#
	.SYNOPSIS
    Get Messages from Outlook and Create a Messages Report

	.DESCRIPTION
    Info about messages from Configured Outlook Account

	.PARAMETER
    -ArchivePath - Path for Export and Report

	.EXAMPLE
	PS C:\> .\Outlook-ReportMessagePath -ReportPath C:\temp

    .OUTPUTS
		CSV Report File in Selected Directory and MSG Files with  messages by Folders

	.NOTES
		Take It, Hold It, Love It

	.LINK
		Author : Andrew V. Golubenkoff (andrew.golubenkoff@outlook.com)

    .lINK

#>

param(
    [parameter(Mandatory = $false)][string]$ReportPath = 'C:\TEMP'
)

$LogFile = Join-Path $ReportPath "ExportLog_$(Get-Date -f 'ddMMyyyyHHmm').csv"

# Create class for Logger
# Connect to Outlook
$outlook = New-Object -com outlook.application;
$ns = $outlook.GetNameSpace('MAPI');

if ($ns.CurrentProfileName) {
    Write-Host $($ns.CurrentProfileName.ToString() + ': ' + $ns.CurrentUser.Name.ToString() ) -BackgroundColor DarkCyan -ForegroundColor Black

    # Select Mailbox or Connected PST File
    $SelectedProfileID = $ns.Folders | Where-Object FullFolderPath -EQ $($ns.Folders | Select-Object Name,FullFolderPath -Unique | Out-GridView -Title 'Select Source' -OutputMode Single).FullFolderPath

} else {
    Write-Host 'No Connection' -BackgroundColor red -ForegroundColor White
}

Write-Host "[$($SelectedProfileID.Name)] Folders count     : " $SelectedProfileID.Folders.Count

#region Functions

class LogMessage {
    [string]$MessagePath
    [string]$DateTime = ([datetime]::Now).ToString('dd.MM.yyyy HH:mm:ss')
    [string]$Subject
    [string]$FromAddress
    [string]$FromName
    [string]$To
    [string]$Attachments

    LogMessage() {
    }

    LogMessage([__ComObject]$item) {
        $this.Subject = $item.Subject
        $this.DateTime = if ($item.ReceivedTime) { $item.ReceivedTime.ToString('dd.MM.yyyy HH:mm:ss') }else { $item.CreationTime.ToString('dd.MM.yyyy HH:mm:ss') }
        $this.FromAddress = $item.Sender.Address
        $this.FromName = $item.Sender.Name
        $this.To = $(($item.Recipients | Select-Object -ExpandProperty Address) -join ",")
        $this.Attachments = $(($item.Attachments | Select-Object -ExpandProperty FileName) -join ',')
    }

    ExportCsv ($LogFile) {
        $this | Select-Object * | Export-Csv -Path $LogFile -Append -NoTypeInformation -Encoding UTF8 -UseCulture
    }

}


Function ProcessItems {
    param($sFolder,$ItemPath)

    Write-Host 'Processing Folder: ' $sFolder.Name `t $ItemPath
    # save current path items
    foreach ($item in $sFolder.Items) {

        $Log = [LogMessage]::New($item)
        $Log.MessagePath = Join-Path $ItemPath "$($item.Subject)"
        $Log.ExportCsv($LogFile)
    }
    # check if subFolder Exist
    if ($sFolder.Folders.Count -gt 0) {
        foreach ($subFolder in $sFolder.Folders) {
            ProcessItems -sFolder $subFolder -ItemPath $(Join-Path $ItemPath $subFolder.Name) -ArchivePath $ArchivePath
        }
    }
}
#endregion Functions

ProcessItems -sFolder $SelectedProfileID -ItemPath $SelectedProfileID.FolderPath
