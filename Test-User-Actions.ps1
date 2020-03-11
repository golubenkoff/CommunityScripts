# Set SearchBase - for your environment

$Password = "NewPassw0rd!"

#region ActiveDirectory


# User1..User20
    # Password Reset
    1..20 | %{try{get-aduser -Identity "User$($_)" -ea stop | Set-ADAccountPassword -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force) -Verbose}catch{Write-Host $($_)} }
    
    # Change Company Name & Description ( for example )
    1..20 | %{try{get-aduser -Identity "User$($_)" -ea stop | Set-ADUser -Description "Test User Description" -Company "MY Test Company" -Verbose}catch{Write-Host $($_)} }

# Searching by User Name - User* in specific OU
    # Password Reset
    try{get-aduser -filter {Name -like "User*"}  -SearchBase "OU=Test,DC=testad,DC=com" -SearchScope Subtree -ea stop | Set-ADAccountPassword -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force) -Verbose}catch{Write-Host $($_)} 

    # Set User Password never expire (for example)
    try{get-aduser -filter {Name -like "User*"}  -SearchBase "OU=Test,DC=testad,DC=com" -SearchScope Subtree -ea stop | set-aduser -PasswordNeverExpires $true -ChangePasswordAtLogon $false -Verbose}catch{Write-Host $($_)} 
#endregion ActiveDirectory

#region ADSI

    # Searching users without 'ActiveDirectory' module for PowerShell - pure .Net 
    ([adsisearcher]"(&(objectClass=user)(name=User*))").FindAll() | %{ $_.GetDirectoryEntry()} | Select Name,Company,Description 

    # Resetting Passwords ( for example )
    # Working on any computer - you need to run this command with user with appropriate permissions in ActiveDirectory
    ([adsisearcher]"(&(objectClass=user)(name=User*))").FindAll() | %{ $Entry = $_.GetDirectoryEntry() ; $Entry.Invoke("SetPassword",$Password) ; $Entry.CommitChanges() }

#endregion ADSI