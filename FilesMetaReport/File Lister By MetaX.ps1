#Get-ChildItem -Recurse -File | % fullname | Export-Excel D:\TEMP\MEDIA_ALL.xlsx -Append

# D:\Music (Full) 1 Hour 30 Minutes 13096
# D:\Music (MetaRoot) 1 Hour Count 5259
# D:\Music (MetaRecurse) 25 Minutes 3771
# R:\Music (Full) 6 Hours 40 Minutes Count 1802
# R:\Music (MetaRoot) 5 Hours 00 Minutes Count
# R:\Music (MetaRecurse) 1 Hour 35 Minutes Count
# MEDIA BACKUP1 23 Hours 20 Minutes
# MEDIA BACKUP2 25 Hours
# MEDIA5TB 15 Hours 30 Minutes

# Start Transcript
Start-Transcript D:\LOGS\MEDIA_TRANSCRIPT.txt

# To execute the script without agreeing with the execution policy
Set-ExecutionPolicy Bypass -Scope Process

# Import the PS1 file to access the functions
Import-module .\Get-FileMetaDataReturnObject.ps1 -Force

# Defines the directory where the files will be, to get their metadata
$FilePath = New-Object System.Windows.Forms.FolderBrowserDialog
$FilePath.RootFolder = "MyComputer"
$FilePath.ShowDialog()
Write-Host Selected folder: $FilePath.SelectedPath -ForegroundColor Yellow

# Lets get the Volume Name
$FPSubString = $FilePath.SelectedPath.Substring(0,2)
$VolumeName = Get-CimInstance CIM_LogicalDisk | Where-Object {$_.Name -eq $FPSubString} | Select-Object -ExpandProperty VolumeName -Last 1

# Lets create File Name
$Dir = $FilePath.SelectedPath.Replace(':\','-')
$Dir = $Dir.Replace('\','-')
$Dir = "D:\Temp\LIST-${VolumeName}-${Dir}.xlsx"
Write-Host Filename = $Dir -ForegroundColor Yellow
Start-Sleep -Seconds 30

#Loads metadata into the variable
Get-Date -Format "dddd dd MMMM yyyy hh:mm tt"
Write-Host Processing MetaRoot
$MetaRoot = Get-FileMetaData -Folder $FilePath.SelectedPath
Get-Date -Format "dddd dd MMMM yyyy hh:mm tt"
Write-Host Processing MetaRecurse
$MetaRecurse = Get-FileMetaData -Folder (Get-ChildItem $FilePath.SelectedPath -Recurse -Directory).FullName
Get-Date -Format "dddd dd MMMM yyyy hh:mm tt"

######## LETS EXPORT THE DATA TO AN EXCEL FILE ########
# Wild card selection
#$ExcelDataRoot = $MetaRoot | Where-Object {$_.Name -match 'Jessie'}
#$ExcelDataRecurse = $MetaRecurse | Where-Object {$_.Name -match 'Jessie'}

# Exclude Folder Names and iTunes type files
$ExcelDataRoot = $MetaRoot | Where-Object {$_.'Item type' -ne "File folder"}
$ExcelDataRecurse = $MetaRecurse | Where-Object {$_.'Item type' -ne "File folder" -and $_.'Item type' -ne "ITC2 File" -and
    $_.'Item type' -ne "iTunes Database File"}

$ExcelDataRoot    | Select-Object 'File Location', Name, Path, Title, 'Item Type', Genre, Length, Album, 'Album Artist', Size | Export-Excel -Path $Dir | Sort-Object 'File Location', Name
$ExcelDataRecurse | Select-Object 'File Location', Name, Path, Title, 'Item Type', Genre, Length, Album, 'Album Artist', Size | Export-Excel -Path $Dir -Append | Sort-Object 'File Location', Name

# Include everything
#$MetaRoot | Select-Object 'File Location', Name, Path, 'Item Type', Genre, Length, Album, 'Album Artist', Size | Export-Excel -Path $Dir | Sort-Object 'File Location', Name
#$MetaRecurse | Select-Object 'File Location', Name, Path, 'Item Type', Genre, Length, Album, 'Album Artist' | Export-Excel -Path $Dir | Sort-Object 'File Location', Name

######## LETS OPEN THE EXPORTED FILE ########
$Excel = New-Object -ComObject Excel.Application
$Excel.Windowstate = "xlMaximized"
$Book = $Excel.Workbooks.Open($Dir)
$Sheet = $Book.Worksheets.Item(1)
$Excel.Visible = $true

######## LETS ADD THE FILE COUNT ########
$RowCount = $Excel.ActiveSheet.UsedRange.Rows.Count - 1
$Excel.ActiveCell.Cells.Item(1,10) = 'File Count'
$Excel.ActiveCell.Cells.Item(1,11) = $RowCount
$Excel.ActiveCell.Cells.Item(1,11).Interior.ColorIndex = 6

######## LETS ADD VOLUME NAME COLUMN ########
$Range1 = $Sheet.Range("A:A")
$Range1.EntireColumn.Insert($ExcelShiftLeft)
$Excel.ActiveCell.Cells.Item(1,1) = 'Volume Name'
$Excel.ActiveCell.Cells.Item(2,1) = $VolumeName

$VolumeCount = $RowCount + 1
$Sheet = $Book.Worksheets.Item(1).Range("A2")
$Sheet.Copy()
$Excel.Range("A3:A$VolumeCount").Select()
$Excel.Worksheets.Item(1).paste()

######## LETS FORMAT THE EXPORTED FILE ########
$Excel.Rows.Item(1).font.bold=$true
$Excel.ActiveSheet.Cells.EntireColumn.autofit()
$Excel.Rows.Item(2).Select()
$Excel.ActiveWindow.FreezePanes = $true
$Excel.Range("A1").Select()
$Excel.Worksheets.Item(1).name = $VolumeName
$Book.Save()
Beep 600 1000 /s 50 /r 3
Stop-Transcript