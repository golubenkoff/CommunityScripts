$a = @'
<style>
BODY{
    font-family: Arial;
    font-size: 8pt;
}
H1{
    font-family: "Courier New", Courier, monospace;
    font-size: 16px;
    font-weight: bold;
}
H2{
    font-size: 14px;
    font-weight: bold;
}
H3{
    font-size: 12px;
    font-weight: bold;
}
TABLE{
    font-family: "Courier New", Courier, monospace;
    border: 1px solid #1C6EA4;
    border-collapse: collapse;
    font-size: 10pt;
    text-align: left;
}
TH{
    font-size: 14px;
    font-weight: bold;
    color: #FFFFFF;
    border-left: 2px solid #D0E4F5;
    border: 1px solid #DADADA;
    padding: 5px;
    background: #414141;
    background: -moz-linear-gradient(top, #707070 0%, #545454 66%, #414141 100%);
    background: -webkit-linear-gradient(top, #707070 0%, #545454 66%, #414141 100%);
    background: linear-gradient(to bottom, #707070 0%, #545454 66%, #414141 100%);
    border-bottom: 2px solid #444444;
}
TD{border: 1px solid #DADADA; padding: 3px; }
td.pass{background: #7FFF00;}
td.warn{background: #FFE600;}
td.fail{background: #FF0000; color: #ffffff;}
td.info{background: #85D4FF;}

</style>
'@
$ScriptPath = $MyInvocation.MyCommand.Path
$Folder = Split-Path -Parent $ScriptPath

$GP_Links = (Get-ADOrganizationalUnit -Filter * | Get-GPInheritance).GpoLinks | Select-Object -Property Target,DisplayName,Enabled,Enforced,Order
$GP_LinksGroup = $GP_Links | Group-Object -Property DisplayName


$GPOs = Get-GPO -Domain $env:USERDNSDOMAIN -All | Select-Object DisplayName,Description,Domain,Owner,GpoStatus,*Time,
@{N = 'WmiFilter';E = { if ($_.WmiFilter -ne $null) { $_.wmifilter.Name }else { '-' } } },
@{N = 'GPO Links';E = {
        $P = $null ; $P = $_.DisplayName ; 
($GP_LinksGroup | Where-Object Name -EQ $P | Select-Object -ExpandProperty Group | Where-Object Enabled -EQ $true | ForEach-Object {
            $_.Target + " | Enforced: $($_.Enforced) | Order: $($_.Order)"
        }) -join '||'

    }
}



$Body = '<br>'
$Body += "<h1>GPO Report $(Get-Date -f 'dd.MM.yyyy')</h1>"

$body += '<br><hr><br>'

$Body += $GPOs | ConvertTo-Html -Head $a
$Body -replace '\|\|','<br>' -replace '<table>','<table id="myTable" class="table w-auto table-condensed table-sm  table-bordered table-striped table-hover dataTable no-footer display nowrap">' | Out-File $(Join-Path $Folder "GPO_Report_$(Get-Date -f 'ddMMyyyyHHmm').html") -Encoding utf8


