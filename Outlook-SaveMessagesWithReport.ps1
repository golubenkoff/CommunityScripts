<#
	.SYNOPSIS
    Export Messages from Outlook to MSG Files or Just Create a Messages Report

	.DESCRIPTION
    Export messages from Configured Outlook Account

	.PARAMETER
    -ArchivePath - Path for Export and Report

	.PARAMETER
    -SaveToMsg - true - export and create report, false - only report

	.EXAMPLE
	PS C:\> .\Outlook-SaveMessagesWithReport.ps1 -ArchivePath C:\temp -SaveToMsg $true

    .OUTPUTS
		CSV Report File in Selected Directory and MSG Files with  messages by Folders

	.NOTES
		Take It, Hold It, Love It

	.LINK
		Author : Andrew V. Golubenkoff (andrew.golubenkoff@outlook.com)

    .lINK

#>

param(
    [parameter(Mandatory = $false)][string]$ArchivePath = 'C:\TEMP',
    [parameter(Mandatory = $false)][bool]$SaveToMsg = $false
)

$LogFile = Join-Path $ArchivePath "ExportLog_$(Get-Date -f 'ddMMyyyyHHmm').csv"

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

$StartFolder = $SelectedProfileID.Folders | Where-Object FolderPath -EQ $($SelectedProfileID.Folders | Select-Object Name,FolderPath | Out-GridView -Title 'Select Folder for Export' -OutputMode Single).FolderPath
Write-Host "[$($StartFolder.Name)] Items count     : " $StartFolder.Items.Count
Write-Host "[$($StartFolder.Name)] subFolders count: " $StartFolder.Folders.Count

#region Functions
Function FixFilename($string) {

    # Loop through each invalid character and replace it with an underscore
    foreach ($char in [IO.Path]::GetInvalidFileNameChars()) {
        $string = $string.Replace($char, '_')
    }
    return $string
}


class LogMessage {
    [string]$DateTime = ([datetime]::Now).ToString('dd.MM.yyyy HH:mm:ss')
    [string]$Subject
    [string]$FromAddress
    [string]$FromName
    [string]$To
    [string]$Attachments
    [string]$FilePath

    LogMessage() {
    }

    LogMessage([__ComObject]$item) {
        $this.Subject = $item.Subject
        $this.DateTime = if ($item.ReceivedTime) { $item.ReceivedTime.ToString('dd.MM.yyyy HH:mm:ss') }else { $item.CreationTime.ToString('dd.MM.yyyy HH:mm:ss') }
        $this.FromAddress = $item.Sender.Address
        $this.FromName = $item.Sender.Name
        $this.To = $($item.Recipients | Select-Object Name,Address | Format-Table -AutoSize -Wrap | Out-String)
        $this.Attachments = $(($item.Attachments | Select-Object -ExpandProperty FileName) -join ',')
    }

    ExportCsv ($LogFile) {
        $this | Select-Object * -ExcludeProperty ProxyAddresses | Export-Csv -Path $LogFile -Append -NoTypeInformation -Encoding UTF8 -UseCulture
    }

}


Function ProcessItems {
    param($sFolder,$ItemPath,$ArchivePath,[bool]$SaveToMsg)

    Write-Host 'Processing Folder: ' $sFolder.Name `t $ItemPath
    # save current path items
    foreach ($item in $sFolder.Items) {

        $Log = [LogMessage]::New($item)

        if ($SaveToMsg) {
            $targetPath = $null ; $targetPath = $(Join-Path $ArchivePath $($ItemPath -replace '\\\\'))
            if (!([System.IO.Directory]::Exists($targetPath))) {
                [void][System.IO.Directory]::CreateDirectory($targetPath)
            }

            $fileName = Join-Path $targetPath "$(FixFileName($item.Subject))_$($item.CreationTime.toString('ddMMyyyyHHmmss')).msg"
            $Log.FilePath = $fileName
            $item.SaveAs($fileName)
        }
        $Log.ExportCsv($LogFile)
    }
    # check if subFolder Exist
    if ($sFolder.Folders.Count -gt 0) {
        foreach ($subFolder in $sFolder.Folders) {
            ProcessItems -sFolder $subFolder -ItemPath $(Join-Path $ItemPath $subFolder.Name) -ArchivePath $ArchivePath -SaveToMsg $SaveToMsg
        }
    }
}
#endregion Functions

$sFolder = $StartFolder
$ItemPath = $StartFolder.FolderPath
ProcessItems -sFolder $StartFolder -ItemPath $StartFolder.FolderPath -ArchivePath $ArchivePath -SaveToMsg $SaveToMsg
