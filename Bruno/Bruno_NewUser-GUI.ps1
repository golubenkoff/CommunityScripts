<#
	.SYNOPSIS
		

	.DESCRIPTION
        The specs for the script I have to write are; 
        Create a new user copying from a current user. 

        Step 1 Input New User First Name, New User Last Name. 

        This will propagate Username, Username (Pre-Windows 2000), First Name, Last Name & Display Name 

        Step 2 Input Current user’s logon name to copy the following fields from; 
        Description 
        Job Title 
        Department 
        Company 
        Manager 
        All the AD Groups in Member of 
        Move the newly created account to the same OU as the user the fields were copied from  
        Default Password = Welcome1  

	.PARAMETER    
		

	.PARAMETER     
		

	.PARAMETER  
	

	.EXAMPLE
		PS C:\> 

	.OUTPUTS
		

	.NOTES
		Take It, Hold It, Love It

	.LINK
		Author : Andrew V. Golubenkoff (andrew.golubenkoff@outlook.com)

	.LINK
		

#>
Function CreateCompanyUser {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,Position = 0,HelpMessage = "New userName ie 'John Smith'")]
        [ValidatePattern(
            '^.+ .+$'
        )
        ] 
        [string]
        $newUser,

        [Parameter(Mandatory = $true,Position = 1,HelpMessage = 'Source oldUser name ie tanderson')]
        [string]
        $oldUser,

        [Parameter(Mandatory = $false)]
        [string]
        [ValidateNotNullOrEmpty()]
        $DefaultPassword = 'Welcome1'
    )


    $DC = $env:LOGONSERVER -replace '\\\\'

    $givenName = $newuser.split()[0] 
    $surname = $newuser.split()[1] 
    $username = $($givenName.Substring(0,1) + $surname.Split(0,19)).ToLower() 

    if ( [System.String]::IsNullOrEmpty($username)) { throw 'Error: No UserName, please check paramers' }

    $oldUserParams = $null
    $oldUserParams = try { Get-ADUser -filter {SamAccountName -eq $oldUser} -Properties * -ErrorAction Stop }
    catch {
        Write-Host "Error: [$oldUser]" -ForegroundColor red -NoNewline ; Write-Host $_.Exception.Message -ForegroundColor Yellow ; return 1
    }

    if (!($oldUserParams)){Write-Host "Old User Not Found..." -ForegroundColor Red ;return 1}

    Write-Host 'Creating User:' -ForegroundColor DarkGray


    Write-Host `t 'UserName : ' -ForegroundColor DarkGray -NoNewline;  Write-Host $username -ForegroundColor Cyan
    Write-Host `t 'givenName: ' -ForegroundColor DarkGray -NoNewline;  Write-Host $givenName -ForegroundColor Cyan
    Write-Host `t 'surName  : ' -ForegroundColor DarkGray -NoNewline;  Write-Host $surname -ForegroundColor Cyan

    Write-Host `t 'OldUsere : ' -ForegroundColor DarkGray -NoNewline;  Write-Host $oldUserParams.name -ForegroundColor Cyan
    #region Create new userName
    $check = $null ; $check = try { Get-ADUser $username -Server $DC -ErrorAction Stop }catch {}
    if (!$check) {
        New-ADUser -Path $($oldUserParams.DistinguishedName -replace 'CN=.+?,') -AccountPassword (ConvertTo-SecureString -AsPlainText $DefaultPassword -Force) `
            -GivenName "$givenName" `
            -Surname "$surname" `
            -DisplayName "$givenName $surname" `
            -Name "$givenName $surname" `
            -Enabled $true `
            -SamAccountName $username `
            -UserPrincipalName $($oldUserParams.UserPrincipalName -replace $oldUserParams.SamAccountName,$username) `
            -Server $DC `
            -ErrorAction Stop
    } else {
        Write-Host "[$($username)]" -ForegroundColor Green -NoNewline; Write-Host ' Already exist' -ForegroundColor red; return 1
    }
    #endregion Create new userName

    #region set newUser Params
    $newUserID = $null ; $newUserID = Get-ADUser $username -Server $DC 
    if ($null -ne $newUserID) {

        $newUserParameters = [hashtable]@{}

        if ($oldUserParams.Description) { $newUserParameters.add('Description',$oldUserParams.Description) }
        if ($oldUserParams.Title) { $newUserParameters.add('Title',$oldUserParams.Title) }
        if ($oldUserParams.Department) { $newUserParameters.add('Department',$oldUserParams.Department) }
        if ($oldUserParams.Company) { $newUserParameters.add('Company',$oldUserParams.Company) }
        if ($oldUserParams.Manager) { $newUserParameters.add('Manager',$oldUserParams.Manager) }

    
        $newUserID | Set-ADUser @newUserParameters -Server $DC

        $oldUserParams.MemberOf | ForEach-Object { Write-Host "Adding [$username] to $($_):" ; Add-ADGroupMember -Identity $_ -Members $newUserID -Server $DC }
        return 0
    } else {
        Write-Host "[$($username)]" -ForegroundColor Green -NoNewline; Write-Host ' user not found. Please check if created before.' -ForegroundColor red; return 1
    }
    #endregion set newUser Params
}

#region GUI
#region Script Settings
#<ScriptSettings xmlns="http://tempuri.org/ScriptSettings.xsd">
#  <ScriptPackager>
#    <process>powershell.exe</process>
#    <arguments />
#    <extractdir>%TEMP%</extractdir>
#    <files />
#    <usedefaulticon>true</usedefaulticon>
#    <showinsystray>false</showinsystray>
#    <altcreds>false</altcreds>
#    <efs>true</efs>
#    <ntfs>true</ntfs>
#    <local>false</local>
#    <abortonfail>true</abortonfail>
#    <product />
#    <version>1.0.0.1</version>
#    <versionstring />
#    <comments />
#    <company />
#    <includeinterpreter>false</includeinterpreter>
#    <forcecomregistration>false</forcecomregistration>
#    <consolemode>false</consolemode>
#    <EnableChangelog>false</EnableChangelog>
#    <AutoBackup>false</AutoBackup>
#    <snapinforce>false</snapinforce>
#    <snapinshowprogress>false</snapinshowprogress>
#    <snapinautoadd>2</snapinautoadd>
#    <snapinpermanentpath />
#    <cpumode>1</cpumode>
#    <hidepsconsole>false</hidepsconsole>
#  </ScriptPackager>
#</ScriptSettings>
#endregion

#region ScriptForm Designer

#region Constructor

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

#endregion

#region Post-Constructor Custom Code

#endregion

#region Form Creation
#Warning: It is recommended that changes inside this region be handled using the ScriptForm Designer.
#When working with the ScriptForm designer this region and any changes within may be overwritten.
#~~< Form1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Form1 = New-Object System.Windows.Forms.Form
$Form1.ClientSize = New-Object System.Drawing.Size(367, 212)
$Form1.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
$Form1.Text = "Create Company User"
#~~< Button_Create >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button_Create = New-Object System.Windows.Forms.Button
$Button_Create.Location = New-Object System.Drawing.Point(199, 173)
$Button_Create.Size = New-Object System.Drawing.Size(75, 23)
$Button_Create.TabIndex = 0
$Button_Create.Text = "Create"
$Button_Create.UseVisualStyleBackColor = $true
$Button_Create.add_Click({Button_CreateClick($Button_Create)})
$Form1.AcceptButton = $Button_Create
#~~< Button_Cancel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button_Cancel = New-Object System.Windows.Forms.Button
$Button_Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$Button_Cancel.Location = New-Object System.Drawing.Point(280, 173)
$Button_Cancel.Size = New-Object System.Drawing.Size(75, 23)
$Button_Cancel.TabIndex = 1
$Button_Cancel.Text = "Cancel"
$Button_Cancel.UseVisualStyleBackColor = $true
$Button_Cancel.add_Click({Button_CancelClick($Button_Cancel)})
$Form1.CancelButton = $Button_Cancel
#~~< ProgressBar >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(12, 172)
$ProgressBar.Size = New-Object System.Drawing.Size(181, 23)
$ProgressBar.TabIndex = 8
$ProgressBar.Text = ""
#~~< TextBox_Password >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TextBox_Password = New-Object System.Windows.Forms.TextBox
$TextBox_Password.Font = New-Object System.Drawing.Font("Tahoma", 8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](204)))
$TextBox_Password.Location = New-Object System.Drawing.Point(12, 128)
$TextBox_Password.Size = New-Object System.Drawing.Size(343, 21)
$TextBox_Password.TabIndex = 7
$TextBox_Password.Text = "Welcome1"
#~~< TextBox_oldUser >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TextBox_oldUser = New-Object System.Windows.Forms.TextBox
$TextBox_oldUser.Font = New-Object System.Drawing.Font("Tahoma", 8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](204)))
$TextBox_oldUser.Location = New-Object System.Drawing.Point(12, 77)
$TextBox_oldUser.Size = New-Object System.Drawing.Size(343, 21)
$TextBox_oldUser.TabIndex = 6
$TextBox_oldUser.Text = ""
#~~< TextBox_newUser >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TextBox_newUser = New-Object System.Windows.Forms.TextBox
$TextBox_newUser.Font = New-Object System.Drawing.Font("Tahoma", 8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](204)))
$TextBox_newUser.Location = New-Object System.Drawing.Point(13, 28)
$TextBox_newUser.Size = New-Object System.Drawing.Size(342, 21)
$TextBox_newUser.TabIndex = 5
$TextBox_newUser.Text = ""
#~~< Label3 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label3 = New-Object System.Windows.Forms.Label
$Label3.Location = New-Object System.Drawing.Point(12, 113)
$Label3.Size = New-Object System.Drawing.Size(343, 23)
$Label3.TabIndex = 4
$Label3.Text = "Enter Default User Password"
#~~< Label2 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label2 = New-Object System.Windows.Forms.Label
$Label2.Location = New-Object System.Drawing.Point(12, 61)
$Label2.Size = New-Object System.Drawing.Size(332, 23)
$Label2.TabIndex = 3
$Label2.Text = "Enter Current Usercode to duplicate (eg jsmith)"
#~~< Label1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label1 = New-Object System.Windows.Forms.Label
$Label1.Location = New-Object System.Drawing.Point(13, 13)
$Label1.Size = New-Object System.Drawing.Size(342, 23)
$Label1.TabIndex = 2
$Label1.Text = "Enter Name of  the New User (eg John Smith)"
$Form1.Controls.Add($ProgressBar)
$Form1.Controls.Add($TextBox_Password)
$Form1.Controls.Add($TextBox_oldUser)
$Form1.Controls.Add($TextBox_newUser)
$Form1.Controls.Add($Label3)
$Form1.Controls.Add($Label2)
$Form1.Controls.Add($Label1)
$Form1.Controls.Add($Button_Cancel)
$Form1.Controls.Add($Button_Create)

#endregion

#region Custom Code

#endregion

#region Event Loop

function Main{
	[System.Windows.Forms.Application]::EnableVisualStyles()
	[System.Windows.Forms.Application]::Run($Form1)
}

#endregion

#endregion

#region Event Handlers

function Button_CreateClick( $object ){
if([string]::IsNullOrEmpty($TextBox_newUser.Text)){}
elseif([string]::IsNullOrEmpty($TextBox_oldUser.Text)){}
elseif([string]::IsNullOrEmpty($TextBox_Password.Text)){}
else{
$ProgressBar.Value = 20
$Result = CreateCompanyUser -newUser $TextBox_newUser.Text -oldUser $TextBox_oldUser.Text -DefaultPassword $TextBox_Password.Text
if ($Result -eq 0){$ProgressBar.Value = 100 ; $ProgressBar.ForeColor = "green"}else{$ProgressBar.Value = 0 ;$ProgressBar.ForeColor = "red" }
}
}

function Button_CancelClick( $object ){
$Form1.Close()
$Form1.Dispose()
exit
}

Main # This call must remain below all other event functions

#endregion

#endregion GUI