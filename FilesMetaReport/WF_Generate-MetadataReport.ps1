param($RootFolder,$LogFolder)

# To execute the script without agreeing with the execution policy
Set-ExecutionPolicy Bypass -Scope Process

if (!$LogFolder) {
    $ScriptPath = $MyInvocation.MyCommand.Path
    $LogFolder = Split-Path -Parent $ScriptPath
}

# Start Transcript
Start-Transcript $(Join-Path $LogFolder 'MEDIA_TRANSCRIPT.txt')

#region Functions
workflow WF_MetadataReportFile {
    param(
        [Parameter(Mandatory = $true)]$FolderPath,
        $ThrottleLimit = 50
    )

    sequence {

        $Files = Get-ChildItem -Path $FolderPath -File

        if ($Files.Count -lt $ThrottleLimit) { $ThrottleLimit = $Files.Count }

        Write-Verbose "ThrottleLimit: $ThrottleLimit"
        Write-Verbose "Files Count  : $(($Files | Measure-Object).Count)"

        foreach -Parallel -ThrottleLimit $ThrottleLimit ($file in $Files) {

            InlineScript {
                $shell = New-Object -ComObject Shell.Application
                $File = $using:file
                $folder = $shell.NameSpace($File.DirectoryName)
                $fileObject = $folder.ParseName($file.Name)

                [PSCustomObject]@{
                    Folder      = $File.DirectoryName
                    Name        = $file.Name
                    Path        = $file.FullName
                    Title       = $folder.GetDetailsOf($fileObject, 21)
                    ItemType    = $folder.GetDetailsOf($fileObject, 2)
                    Genre       = $folder.GetDetailsOf($fileObject, 16)
                    Length      = $folder.GetDetailsOf($fileObject, 27)
                    Album       = $folder.GetDetailsOf($fileObject, 14)
                    AlbumArtist = $folder.GetDetailsOf($fileObject, 13)
                    Size        = $file.Length
                }
            }
        }
    }
}
function Remove-InvalidFilePathCharacters {
    param(
        [string]$FilePath
    )

    # Get the invalid characters for a file path
    $invalidChars = [System.IO.Path]::GetInvalidPathChars() | ForEach-Object { [regex]::Escape($_) }

    # Remove the invalid characters from the file path
    $cleanFilePath = [regex]::Replace($FilePath, "[$invalidChars]", '')

    return $cleanFilePath
}

function Read-FolderBrowserDialog([string]$Message, [string]$InitialDirectory, [switch]$NoNewFolderButton)
{
    $browseForFolderOptions = 0
    if ($NoNewFolderButton) { $browseForFolderOptions += 512 }

    $app = New-Object -ComObject Shell.Application
    $folder = $app.BrowseForFolder(0, $Message, $browseForFolderOptions, $InitialDirectory)
    if ($folder) { $selectedDirectory = $folder.Self.Path } else { $selectedDirectory = '' }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) > $null
    return $selectedDirectory
}


#endregion Functions


#region  Defines the directory where the files will be, to get their metadata
Write-Host "[$(Get-Date -f 'dddd dd MMMM yyyy hh:mm:ss tt')] Select RootFolder: " -BackgroundColor cyan -ForegroundColor white

$FilePath = Read-FolderBrowserDialog -Message 'Select Root Folder' -InitialDirectory  $(if ($RootFolder) { $RootFolder }else { 'MyComputer' })

if ([string]::IsNullOrEmpty($FilePath)) {
    Write-Host 'No Root Folder Selected... Aborting' -BackgroundColor Red -ForegroundColor White
    break
}

Write-Host Selected folder: $FilePath -ForegroundColor Yellow

# Lets get the Volume Name
$FPSubString = $FilePath.Substring(0,2)
$VolumeName = Get-CimInstance CIM_LogicalDisk | Where-Object { $_.Name -eq $FPSubString } | Select-Object -ExpandProperty VolumeName -Last 1

# Lets create File Name
$Dir = (Remove-InvalidFilePathCharacters $FilePath) -replace ':','-' -replace '\\','-'
$Dir = Join-Path $LogFolder "LIST-${VolumeName}-${Dir}-$(get-date -f 'ddMMyyyyHHmm').xlsx"
Write-Host "Filename = $Dir" -ForegroundColor Yellow
#endregion  Defines the directory where the files will be, to get their metadata


#region Processing Data
Write-Host "[$(Get-Date -f 'dddd dd MMMM yyyy hh:mm:ss tt')] Processing Metadata for RootFolder: " $FilePath

$StartMetaRoot = Get-Date
$MetaRoot = WF_MetadataReportFile -FolderPath $FilePath
$MetaRoot.count
if (($MetaRoot | Measure-Object).count -gt 0) {
    $MetaRoot.psobject.Properties.Remove('PSComputerName')
    $MetaRoot.psobject.Properties.Remove('PSShowComputerName')
    $MetaRoot.psobject.Properties.Remove('PSSourceJobInstanceId')
}
Get-Job | Remove-Job
$EndMetaRoot = Get-Date
Write-Host 'MetaRoot Total Time: ' $($EndMetaRoot - $StartMetaRoot).totalSeconds


$StartMetaRecurse = Get-Date
(Get-ChildItem $FilePath -Recurse -Directory) | ForEach-Object {
    Write-Progress -Activity 'Starting Jobs for subFolders' -CurrentOperation $_.Name
    [void](WF_MetadataReportFile -FolderPath $_.FullName -AsJob)
}
Write-Host 'Please Wait for All Jobs to Complete' -back darkcyan -for white
$MetaRecurse = Get-Job | Wait-Job | Receive-Job
$MetaRecurse.Count
if (($MetaRecurse | Measure-Object).count -gt 0) {
    $MetaRecurse.psobject.Properties.Remove('PSComputerName')
    $MetaRecurse.psobject.Properties.Remove('PSShowComputerName')
    $MetaRecurse.psobject.Properties.Remove('PSSourceJobInstanceId')
}
Get-Job | Remove-Job
$EndMetaRecurse = Get-Date
Write-Host 'MetaRecurse Total Time: ' $($EndMetaRecurse - $StartMetaRecurse).totalSeconds

#endregion Processing Data


#region Export to Excel
[PSCustomObject]@{
    ReportCreated   = $(Get-Date -f 'dddd dd MMMM yyyy hh:mm:ss tt')
    VolumeName      = $VolumeName
    RootFolder      = $FilePath
    TotalFilesCount = ($MetaRoot | Measure-Object).count + ($MetaRecurse | Measure-Object).count
} | Export-Excel -Path $Dir -WorksheetName 'MEDIA STAT' -Title 'Media Statistics' -AutoSize -TableStyle Dark1

[PSCustomObject]@{
    StartTime   = $(Get-Date $StartMetaRoot -f 'dddd dd MMMM yyyy hh:mm:ss tt')
    EndTime      = $(Get-Date $EndMetaRecurse -f 'dddd dd MMMM yyyy hh:mm:ss tt')
    RootFolder      = $FilePath
    TotalTime  = $($EndMetaRecurse-$StartMetaRoot)
} | Export-Excel -Path $Dir -WorksheetName 'TIME STAT' -Title 'Time Statistics' -AutoSize -TableStyle Dark1 -Append


$MetaRoot | Where-Object { $_.'Item type' -ne 'ITC2 File' -and $_.'Item type' -ne 'iTunes Database File' } | Select-Object @{N = 'VolumeName';E = { $VolumeName } },* | Sort-Object Folder,Name | Export-Excel -Path $Dir -WorksheetName 'MEDIA METADATA' -TableStyle Light10
$MetaRecurse | Where-Object { $_.'Item type' -ne 'ITC2 File' -and $_.'Item type' -ne 'iTunes Database File' } | Select-Object @{N = 'VolumeName';E = { $VolumeName } },* | Sort-Object Folder,Name | Export-Excel -Path $Dir -WorksheetName 'MEDIA METADATA'  -Append

#endregion Export to Excel


# Fun Staff
1..3 | ForEach-Object { [console]::beep(600, 1000) ; Start-Sleep -Milliseconds 50 }

Stop-Transcript
