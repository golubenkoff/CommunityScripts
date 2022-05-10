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
    $pool = get-item IIS:\AppPools\$PoolName
    $pool
    Restart-WebAppPool $pool.Name -Verbose
}
