param(
[string]$ComputerName,
[string]$inFile,
[int]$Days = 300
)

$List = gc $inFile -ErrorAction stop




function get-logonhistory{
Param (
 [string]$ComputerName,
 [int]$Days = 365
 )
 Write-Host "Processing: " -ForegroundColor DarkGray -NoNewline ; Write-Host $ComputerName -ForegroundColor green
 $Result = @()
 $ELogs = Get-EventLog System -Source Microsoft-Windows-WinLogon -After (Get-Date).AddDays(-$Days) -ComputerName $ComputerName
 If ($ELogs)
 { 
 
     ForEach ($Log in $ELogs)
     { If ($Log.InstanceId -eq 7001)
       { $ET = "Logon"
       }
       ElseIf ($Log.InstanceId -eq 7002)
       { $ET = "Logoff"
       }
       Else
       { Continue
       }
       $Result += New-Object PSObject -Property @{
        Time = $Log.TimeWritten
        'Event Type' = $ET
        User = (New-Object System.Security.Principal.SecurityIdentifier $Log.ReplacementStrings[1]).Translate([System.Security.Principal.NTAccount])
       }
     }
    return $($Result | Select @{N="DateTime";E={Get-date $_.Time -f 'dd.MM.yyyy HH:mm:ss'}},"Event Type",User,@{N="ComputerName";E={$ComputerName}} | Sort DateTime -Descending) # | Out-GridView
 }
 Else
 { 
 Write-Warning "$ComputerName - Check RemoteRegystry"
 }
}


$Report = @()
if ($ComputerName){$List = $ComputerName}
foreach ($ComputerName in $List){
$Data = $null ; $Data = get-logonhistory -ComputerName $ComputerName -Days $Days
if ($Data){
$Report += $Data
}
}


$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
$Report | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $(Join-path $ScriptDir "Logon_Report_for_${Days}_days.csv") -Delimiter ";"

