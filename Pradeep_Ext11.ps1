param(
$FileName = "C:\Scripts\Pradeep_Source.csv",
$SmtpServer = "smtp.mydomain.com",
$From = "HR@mydomain.com"
)

[string]$Body = @"
<span style="font-family: Calibri;">Hi, <span
 style="font-weight: bold;">__USERNAME__<br>
<br style="font-family: Calibri;">
</span></span><span style="font-family: Calibri;">You
are fired.</span><br style="font-family: Calibri;">
<br style="font-family: Calibri;">
<span style="font-family: Calibri;">Joking :). <span
 style="font-weight: bold; color: rgb(255, 0, 0);">Happy
birthday!</span></span><br style="font-family: Calibri;">
<br style="font-family: Calibri;">
<span style="font-family: Calibri; color: rgb(41, 182, 55);">--</span><br
 style="font-family: Calibri; color: rgb(41, 182, 55);">
<span style="font-family: Calibri; color: rgb(41, 182, 55);">HR</span>
"@

$Data = import-csv $FileName -Encoding UTF8

foreach ($user in $Data){
$Name = $null; $Name = $user.Name
$UserID = $null; $UserID = Get-ADUser -Filter {Name -eq $Name} -Properties extensionattribute11,mail,DisplayName
    if ($UserID -and !([string]::IsNullOrEmpty($UserID.extensionattribute11)))
    {
        [string]$uBody = $Body -replace '__USERNAME__',$($UserID.DisplayName)
        Send-MailMessage -From $From -to $UserID.mail -Body $uBody -BodyAsHtml -Subject "Message from HR" -SmtpServer $SmtpServer -Encoding UTF8
    }
}