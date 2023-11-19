function Read-FolderBrowserDialog([string]$Message, [string]$InitialDirectory, [switch]$NoNewFolderButton) {
    $browseForFolderOptions = 0
    if ($NoNewFolderButton) { $browseForFolderOptions += 512 }

    $app = New-Object -ComObject Shell.Application
    $folder = $app.BrowseForFolder(0, $Message, $browseForFolderOptions, $InitialDirectory)
    if ($folder) { $selectedDirectory = $folder.Self.Path } else { $selectedDirectory = '' }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) > $null
    return $selectedDirectory
}



$FilePath = Read-FolderBrowserDialog -Message 'Select Folder to Get Volume Name' -InitialDirectory 'MyComputer'
Write-Host Selected folder: $FilePath

$DirInfo = [System.IO.DirectoryInfo]$FilePath

# Method 1
Write-Host 'Method 1: ' -ForegroundColor Cyan -NoNewline ; Write-Host ([System.IO.DriveInfo]::GetDrives() | Where-Object Name -EQ $DirInfo.Root.Name).VolumeLabel

# Method 2
Write-Host 'Method 2: ' -ForegroundColor Cyan -NoNewline ; Write-Host (Get-Partition -DriveLetter $DirInfo.Root.Name[0] | Get-Volume).FileSystemLabel

# Method 3
Write-Host 'Method 3: ' -ForegroundColor Cyan -NoNewline ; Write-Host (Get-WmiObject win32_volume | Where-Object { $_.DriveLetter } | Where-Object Name -EQ $DirInfo.Root.Name).Label

