<#
	.SYNOPSIS
        Get ACL Report for folder

	.DESCRIPTION
        Generates .CSV file with ACL Report for one or many (Recursive) folers for provided path

	.PARAMETER
        -FolderName - folder to get ACL from

	.PARAMETER
        -Recursive - do a recursive check

	.PARAMETER
        -Depth - depth for recursive check

    .PARAMETER
        -GUI - show GUI for options selection

	.PARAMETER
        -ExportPath - path where save a csv file. If  not provided = current script directory

	.EXAMPLE
		PS C:\> .\Get-FolderACL.ps1 -FolderName C:\TEMP\ -Recursive -Depth 1

    .EXAMPLE
		PS C:\> .\Get-FolderACL.ps1 -GUI

	.OUTPUTS
        CSV File

	.NOTES
		Take It, Hold It, Love It

	.LINK
		Author : Andrew V. Golubenkoff (andrew.golubenkoff@outlook.com)

	.LINK
        https://github.com/golubenkoff/CommunityScripts/blob/master/Get-FolderACL.ps1


#>

[CmdletBinding(DefaultParameterSetName = 'Path')]
param(
    [Parameter(Mandatory = $false,
        ParameterSetName = 'Path',
        HelpMessage = 'Enter Folder Path',
        Position = 0)]
    [ValidateScript( {
            if (-Not (Test-Path $_) ) {
                throw 'Path does not exist'
            }
            Test-Path $_ -PathType Container

        })]
    [string]$FolderName,
    [Parameter(Mandatory = $false,
        ParameterSetName = 'Path',
        HelpMessage = 'switch for Recursive processing',
        Position = 1)][switch]$Recursive,
    [Parameter(Mandatory = $false,
        ParameterSetName = 'Path',
        HelpMessage = 'Depth for Recursive',
        Position = 2)][int]$Depth,
    [Parameter(Mandatory = $false,
        ParameterSetName = 'GUI',
        HelpMessage = 'switch for GUI Interface',
        Position = 2)][switch]$GUI,
    [Parameter(Mandatory = $false,Position = 3)][string]$ExportPath
)

[array]$Report = @()

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

function Read-FolderBrowserDialog([string]$Message, [string]$InitialDirectory, [switch]$NoNewFolderButton) {
    $browseForFolderOptions = 0
    if ($NoNewFolderButton) { $browseForFolderOptions += 512 }

    $app = New-Object -ComObject Shell.Application
    $folder = $app.BrowseForFolder(0, $Message, $browseForFolderOptions, $InitialDirectory)
    if ($folder) { $selectedDirectory = $folder.Self.Path } else { $selectedDirectory = '' }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) > $null
    return $selectedDirectory
}

function get-sid {
    Param (
        $DSIdentity
    )
    $ID = New-Object System.Security.Principal.NTAccount($DSIdentity)
    return $ID.Translate( [System.Security.Principal.SecurityIdentifier] ).toString()
}

Function Convert-SDDLToACL {
    <#
    .Synopsis
    Convert SDDL String to ACL Object
    .DESCRIPTION
    Converts one or more SDDL Strings to a human readable format.
    .EXAMPLE
    Convert-SDDLToACL -SDDLString (get-acl .\path).sddl
    .EXAMPLE
    Convert-SDDLToACL -SDDLString “O:S-1-5-21-1559460989-2589464504-629046386-3966G:DUD:(A;OICIID;FA;;;SY)(A;OICIID;FA;;;BA)(A;OICIID;FA;;;S-1-5-21-1559460989-2589464504-629046386-3966)”
    .NOTES
    Robert Amartinesei
    #>
    [Cmdletbinding()]

    param (
        #One or more strings of SDDL syntax.
        [string[]]$SDDLString
    )
    foreach ($SDDL in $SDDLString) {

        $ACLObject = New-Object -TypeName System.Security.AccessControl.DirectorySecurity
        $ACLObject.SetSecurityDescriptorSddlForm($SDDL)

        $ACLObject.Access
    }
}

Function GetFolderACL {
    param(
        [parameter(mandatory = $true)]$FolderName
    )
    $FolderACL = $null ; $FolderACL = Get-Acl $FolderName
    $ReportACL = @()
    if ($FolderACL) {
        $SDDL = $null ; $SDDL = Convert-SDDLToACL $FolderACL.Sddl
        if ($SDDL) {
            $SDDL | ForEach-Object {
                $ReportACL += [PSCustomObject]@{
                    FolderName        = $FolderName
                    Owner             = $FolderACL.Owner
                    AccessControlType = $_.AccessControlType
                    IdentityReference = $_.IdentityReference
                    IsInherited       = $_.IsInherited
                    InheritanceFlags  = $_.InheritanceFlags
                }
            }
        }
    }
    return $ReportACL
}

if ($GUI) {


    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')

    #~~< Form1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    $Form1 = New-Object System.Windows.Forms.Form
    $Form1.ClientSize = New-Object System.Drawing.Size(425, 209)
    $Form1.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
    $Form1.Text = 'Get Folder ACL'
    #~~< CheckBox1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    $CheckBox1 = New-Object System.Windows.Forms.CheckBox
    $CheckBox1.Location = New-Object System.Drawing.Point(8, 128)
    $CheckBox1.Size = New-Object System.Drawing.Size(104, 24)
    $CheckBox1.TabIndex = 7
    $CheckBox1.Text = 'Recursive'
    $CheckBox1.UseVisualStyleBackColor = $true
    #~~< Button2 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    $Button2 = New-Object System.Windows.Forms.Button
    $Button2.Location = New-Object System.Drawing.Point(280, 80)
    $Button2.Size = New-Object System.Drawing.Size(139, 23)
    $Button2.TabIndex = 6
    $Button2.Text = 'Select ExportPath'
    $Button2.UseVisualStyleBackColor = $true
    $Button2.add_Click( { Button2Click($Button2) })
    #~~< TextBox1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    $TextBox1 = New-Object System.Windows.Forms.TextBox
    $TextBox1.Location = New-Object System.Drawing.Point(152, 128)
    $TextBox1.Size = New-Object System.Drawing.Size(120, 20)
    $TextBox1.TabIndex = 5
    $TextBox1.Text = '1'
    #~~< TextBox_exportPath >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    $TextBox_exportPath = New-Object System.Windows.Forms.TextBox
    $TextBox_exportPath.Location = New-Object System.Drawing.Point(8, 80)
    $TextBox_exportPath.Size = New-Object System.Drawing.Size(264, 20)
    $TextBox_exportPath.TabIndex = 5
    $TextBox_exportPath.Text = ''
    #~~< Label3 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    $Label3 = New-Object System.Windows.Forms.Label
    $Label3.Location = New-Object System.Drawing.Point(152, 112)
    $Label3.Size = New-Object System.Drawing.Size(100, 16)
    $Label3.TabIndex = 4
    $Label3.Text = 'Depth'
    #~~< Label2 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    $Label2 = New-Object System.Windows.Forms.Label
    $Label2.Location = New-Object System.Drawing.Point(8, 64)
    $Label2.Size = New-Object System.Drawing.Size(100, 16)
    $Label2.TabIndex = 4
    $Label2.Text = 'Export Path:'
    #~~< Button_selectFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    $Button_selectFolder = New-Object System.Windows.Forms.Button
    $Button_selectFolder.Location = New-Object System.Drawing.Point(280, 32)
    $Button_selectFolder.Size = New-Object System.Drawing.Size(139, 23)
    $Button_selectFolder.TabIndex = 3
    $Button_selectFolder.Text = 'Select Folder'
    $Button_selectFolder.UseVisualStyleBackColor = $true
    $Button_selectFolder.add_Click( { Button_selectFolderClick($Button_selectFolder) })
    #~~< Label1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    $Label1 = New-Object System.Windows.Forms.Label
    $Label1.Location = New-Object System.Drawing.Point(8, 16)
    $Label1.Size = New-Object System.Drawing.Size(100, 16)
    $Label1.TabIndex = 2
    $Label1.Text = 'FolderName:'
    #~~< TextBox_folderName >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    $TextBox_folderName = New-Object System.Windows.Forms.TextBox
    $TextBox_folderName.Location = New-Object System.Drawing.Point(8, 32)
    $TextBox_folderName.Size = New-Object System.Drawing.Size(264, 20)
    $TextBox_folderName.TabIndex = 1
    $TextBox_folderName.Text = ''
    #~~< Button1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    $Button1 = New-Object System.Windows.Forms.Button
    $Button1.Location = New-Object System.Drawing.Point(336, 168)
    $Button1.Size = New-Object System.Drawing.Size(75, 23)
    $Button1.TabIndex = 0
    $Button1.Text = 'Start'
    $Button1.UseVisualStyleBackColor = $true
    $Button1.add_Click( { Button1Click($Button1) })
    $Form1.Controls.Add($CheckBox1)
    $Form1.Controls.Add($Button2)
    $Form1.Controls.Add($TextBox1)
    $Form1.Controls.Add($TextBox_exportPath)
    $Form1.Controls.Add($Label3)
    $Form1.Controls.Add($Label2)
    $Form1.Controls.Add($Button_selectFolder)
    $Form1.Controls.Add($Label1)
    $Form1.Controls.Add($TextBox_folderName)
    $Form1.Controls.Add($Button1)
    $Form1.add_Activated( { Form1Activated($Form1) })
    #~~< FolderBrowserDialog1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    $FolderBrowserDialog1 = New-Object System.Windows.Forms.FolderBrowserDialog
    $TextBox_exportPath.Text = $scriptPath

    function Main {
        [System.Windows.Forms.Application]::EnableVisualStyles()
        [System.Windows.Forms.Application]::Run($Form1)
    }



    function Button2Click( $object ) {
        $FolderBrowserDialog1.ShowNewFolderButton = $false
        $Text = $null ; $text = $FolderBrowserDialog1.ShowDialog()
        $TextBox_exportPath.Text = $FolderBrowserDialog1.SelectedPath
    }

    function Button_selectFolderClick( $object ) {
        $FolderBrowserDialog1.ShowNewFolderButton = $false
        $Text = $null ; $text = $FolderBrowserDialog1.ShowDialog()
        $TextBox_folderName.Text = $FolderBrowserDialog1.SelectedPath
    }
    function Read-MessageBoxDialog([string]$Message, [string]$WindowTitle, [System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::None) {
        Add-Type -AssemblyName System.Windows.Forms
        return [System.Windows.Forms.MessageBox]::Show($Message, $WindowTitle, $Buttons, $Icon)
    }


    function Button1Click( $object ) {
        $Report = @()
        if ($CheckBox1.Checked) {
            $Report += GetFolderACL $TextBox_folderName.Text
            if (!([string]::IsNullOrEmpty($TextBox1.Text))) {
                $Report += Get-ChildItem $TextBox_folderName.Text -Recurse -Directory -Depth $Depth | ForEach-Object { GetFolderACL $_.FullName }
            } else {
                $Report += Get-ChildItem $TextBox_folderName.Text -Recurse -Directory | ForEach-Object { GetFolderACL $_.FullName }
            }
        } else {
            $Report += GetFolderACL $TextBox_folderName.Text
        }
        if ($TextBox_exportPath.Text.Length -gt 3) {
            $filePath = Join-Path $TextBox_exportPath.Text "Report_FolderACL_$(Get-Date -f 'ddMMyyyyHHmmss').csv"
        } else {
            $filePath = Join-Path $scriptPath "Report_FolderACL_$(Get-Date -f 'ddMMyyyyHHmmss').csv"
        }

        if ($Report) {
            Write-Host "Exporting Report for folder [$($TextBox_folderName.Text))] to file: " -ForegroundColor DarkGray -NoNewline
            Write-Host $filePath -ForegroundColor Cyan

            try {
                $Report | Export-Csv -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Path $filePath -ErrorAction stop
            } catch {
                Write-Host 'Cannot export: ' $_.Exception.Message
            }
            Read-MessageBoxDialog -Message "Exporting ACL Report for folder: [$($TextBox_folderName.Text))] to file: $filepath" -WindowTitle 'Export CSV' -Buttons OK
        }
    }

    function Form1Activated( $object ) {

    }

    Main


} else {
    if ([string]::IsNullOrEmpty($FolderName)) {
        $FolderName = Read-FolderBrowserDialog -InitialDirectory '~' -NoNewFolderButton
    }


    if ($Recursive) {
        $Report += GetFolderACL $FolderName
        if ($Depth) {
            $Report += Get-ChildItem $FolderName -Recurse -Directory -Depth $Depth | ForEach-Object { GetFolderACL $_.FullName }
        } else {
            $Report += Get-ChildItem $FolderName -Recurse -Directory | ForEach-Object { GetFolderACL $_.FullName }
        }
    } else {
        $Report += GetFolderACL $FolderName
    }


    if ($Report) {

        if ($ExportPath) {
            $filePath = Join-Path $ExportPath "Report_FolderACL_$(Get-Date -f 'ddMMyyyyHHmmss').csv"
        } else {
            $filePath = Join-Path $scriptPath "Report_FolderACL_$(Get-Date -f 'ddMMyyyyHHmmss').csv"
        }

        Write-Host "Exporting Report for folder [$FolderName] to file: " -ForegroundColor DarkGray -NoNewline
        Write-Host $filePath -ForegroundColor Cyan

        try {
            $Report | Export-Csv -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Path $filePath -ErrorAction stop
        } catch {
            Write-Host 'Cannot export: ' $_.Exception.Message
        }
    }

}


