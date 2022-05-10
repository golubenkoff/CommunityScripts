<#
Requires on target server:
IIS Management Scripts and Tools            Web-Scripting-Tools
Install-WindowsFeature Web-Scripting-Tools  -ComputerName $ServerName
#>
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

function ConvertAppPoolState {
    param(
        [int] $value
    )
    switch($value)
    {
        0 { "Starting"; break }
        1 { "Started"; break }
        2 { "Stopping"; break }
        3 { "Stopped"; break }
        default { "Unknown"; break }
    }
}

#
# Stops and restarts the specified application pool on the specified machine:
function RecycleAppPool
{
    param (
        [string] $serverName,
        [string] $applicationPool
    )

    # Return an app pool object
    $appPool = Get-WmiObject -Authentication PacketPrivacy -Impersonation Impersonate -ComputerName `
        $serverName -Namespace "root\WebAdministration" -Class "ApplicationPool" `
            | Where-Object { $_.Name -eq "$applicationPool" }

    # Check to make sure we actually have an object
    if ($appPool -ne $null)
    {
        if ((ConvertAppPoolState ($appPool.GetState() | Select-Object -ExpandProperty ReturnValue)) -eq "Started")
        {
            Write-Host -ForegroundColor Green "Attempting to Stop Application Pool: $applicationPool on Server: $server ..."
            $appPool.Stop()

            $stopAttempts = 0
            while (((ConvertAppPoolState ($appPool.GetState() | Select-Object -ExpandProperty ReturnValue)) -eq "Stopping"))
            {
                Write-Host -ForegroundColor Yellow "Stopping..."
                $stopAttempts++
                Start-Sleep 5

                if ($stopAttempts -eq 10)
                {
                    Write-Host -ForegroundColor Red "There was an issue with stopping the Application Pool $applicationPool."
                    exit 1
                }
            }

            if ((ConvertAppPoolState ($appPool.GetState() | Select-Object -ExpandProperty ReturnValue)) -eq "Stopped")
            {
                Write-Host -ForegroundColor Green "Application Pool: $applicationPool on Server: $server stopped."
            }
        }

        if ((ConvertAppPoolState ($appPool.GetState() | Select-Object -ExpandProperty ReturnValue)) -eq "Stopped")
        {
            Write-Host -ForegroundColor Green "Attempting to Start Application Pool: $applicationPool on Server: $server ...."
            $appPool.Start()

            if ((ConvertAppPoolState ($appPool.GetState() | Select-Object -ExpandProperty ReturnValue)) -eq "Started")
            {
                Write-Host -ForegroundColor Green "Application Pool: $applicationPool on Server: $server started."
            }
        }
    }
    else
    {
        Write-Host -ForegroundColor Red "Error occurred while attempting to recycle Application Pool: $applicationPool on Server: $serverName."
        exit 1
    }
}

RecycleAppPool -serverName $ServerName -applicationPool $ApplicationPool