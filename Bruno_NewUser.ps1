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

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true,Position=0,HelpMessage="New userName ie 'John Smith'")]
    [ValidatePattern(
			'^.+ .+$'
			)
	] 
    [string]
    $newUser,

    [Parameter(Mandatory=$true,Position=1,HelpMessage="Source oldUser name ie tanderson")]
    [string]
    $oldUser,

    [Parameter(Mandatory=$false)]
    [string]
    [ValidateNotNullOrEmpty()]
    $DefaultPassword = "Welcome1"
)


$DC =  $env:LOGONSERVER -replace "\\\\"

$givenName = $newuser.split()[0] 
$surname = $newuser.split()[1] 
$username = $($givenName.Substring(0,1)+$surname.Split(0,19)).ToLower() 

if( [System.String]::IsNullOrEmpty($username)){throw "Error: No UserName, please check paramers"}

$oldUserParams = $null
$oldUserParams = try{ Get-ADUser $oldUser -Properties * -ErrorAction Stop}
catch{
Write-Host "Error: [$oldUser]" -ForegroundColor red -NoNewline ; Write-Host  $_.Exception.Message -ForegroundColor Yellow ; Break
}

Write-Host "Creating User:" -ForegroundColor DarkGray
Write-Host `t "UserName : " -ForegroundColor DarkGray -NoNewline;  Write-Host $username -ForegroundColor Cyan
Write-Host `t "givenName: " -ForegroundColor DarkGray -NoNewline;  Write-Host $givenName  -ForegroundColor Cyan
Write-Host `t "surName  : " -ForegroundColor DarkGray -NoNewline;  Write-Host $surname  -ForegroundColor Cyan


#region Create new userName
$check = $null ; $check = try{ Get-ADUser $username -Server $DC -ErrorAction Stop}catch{}
if(!$check){
New-ADUser -path $($oldUserParams.DistinguishedName -replace "CN=.+?,") -AccountPassword (ConvertTo-SecureString -AsPlainText $DefaultPassword -Force) `
-GivenName "$givenName" `
-Surname "$surname" `
-DisplayName "$givenName $surname" `
-Name "$givenName $surname" `
-Enabled $true `
-SamAccountName $username `
-UserPrincipalName $($oldUserParams.UserPrincipalName -replace $oldUserParams.SamAccountName,$username) `
-Server $DC `
-ErrorAction Stop
}else{
    Write-Host "[$($username)]" -ForegroundColor Green -NoNewline; Write-Host " Already exist" -ForegroundColor red; break
}
#endregion Create new userName

#region set newUser Params
$newUserID = $null ; $newUserID = Get-ADUser $username -Server $DC 
if($null -ne $newUserID){

    $newUserParameters = [hashtable]@{}

    if($oldUserParams.Description){$newUserParameters.add('Description',$oldUserParams.Description)}
    if($oldUserParams.Title){$newUserParameters.add('Title',$oldUserParams.Title)}
    if($oldUserParams.Department){$newUserParameters.add('Department',$oldUserParams.Department)}
    if($oldUserParams.Company){$newUserParameters.add('Company',$oldUserParams.Company)}
    if($oldUserParams.Manager){$newUserParameters.add('Manager',$oldUserParams.Manager)}

    
    $newUserID |  Set-ADUser @newUserParameters -Server $DC

    $oldUserParams.MemberOf | %{Write-Host "Adding [$username] to $($_):" ; Add-ADGroupMember -Identity $_ -Members $newUserID -Server $DC}

    }else{
        Write-Host "[$($username)]" -ForegroundColor Green -NoNewline; Write-Host " user not found. Please check if created before." -ForegroundColor red; break
    }
#endregion set newUser Params

