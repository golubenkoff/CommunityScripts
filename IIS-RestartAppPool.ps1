[CmdletBinding()]
    param
    (
        [Parameter(Position = 1, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ServerName,

        [Parameter(Position = 2, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ApplicationPool
    )

invoke-command -ComputerName $ServerName -ArgumentList $ApplicationPool -ScriptBlock {
    param($PoolName)
    import-module WebAdministration
    $pool = get-item “IIS:\Sites\$PoolName” | Select-Object applicationPool
    $pool
    Restart-WebAppPool $pool.applicationPool -Verbose
}


