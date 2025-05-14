<#
.SYNOPSIS
    Checks the current Exchange-related Active Directory schema versions against a specified Exchange 2019 CU version.

.DESCRIPTION
    This function compares the following Active Directory components against the expected values for a specified
    Exchange 2019 Cumulative Update (CU):
    - Schema version (rangeUpper)
    - Default context objectVersion (Microsoft Exchange System Objects)
    - Configuration context objectVersion (Exchange Organization container)

.PARAMETER TargetVersion
    The target Exchange 2019 CU version to compare against (e.g., "CU15", "CU13").

.EXAMPLE
    Check-ExchangeADSchema -TargetVersion "CU15"

    Checks if the current Active Directory schema versions match those required for Exchange 2019 CU15.

.LINK
    https://learn.microsoft.com/en-us/exchange/plan-and-deploy/prepare-ad-and-domains

.NOTES
    Author: Andrew V. Golubenkoff
    Version: 1.0
    Last Updated: 2025-05-14
#>

param (
    [string]$TargetVersion = "CU15"
)

# Define a hashtable mapping Exchange 2019 CU versions to their corresponding AD schema versions
$ExchangeSchemaVersions = @{
    "CU15" = @{ rangeUpper = 17003; objectVersionDefault = 13243; objectVersionConfig = 16763 }
    "CU14" = @{ rangeUpper = 17003; objectVersionDefault = 13243; objectVersionConfig = 16762 }
    "CU13" = @{ rangeUpper = 17003; objectVersionDefault = 13243; objectVersionConfig = 16761 }
    "CU12" = @{ rangeUpper = 17003; objectVersionDefault = 13243; objectVersionConfig = 16760 }
    "CU11" = @{ rangeUpper = 17003; objectVersionDefault = 13242; objectVersionConfig = 16759 }
    "CU10" = @{ rangeUpper = 17003; objectVersionDefault = 13241; objectVersionConfig = 16758 }
    "CU9"  = @{ rangeUpper = 17002; objectVersionDefault = 13240; objectVersionConfig = 16757 }
    "CU8"  = @{ rangeUpper = 17002; objectVersionDefault = 13239; objectVersionConfig = 16756 }
    "CU7"  = @{ rangeUpper = 17001; objectVersionDefault = 13238; objectVersionConfig = 16755 }
    "CU6"  = @{ rangeUpper = 17001; objectVersionDefault = 13237; objectVersionConfig = 16754 }
    "CU5"  = @{ rangeUpper = 17001; objectVersionDefault = 13237; objectVersionConfig = 16754 }
    "CU4"  = @{ rangeUpper = 17001; objectVersionDefault = 13237; objectVersionConfig = 16754 }
    "CU3"  = @{ rangeUpper = 17001; objectVersionDefault = 13237; objectVersionConfig = 16754 }
    "CU2"  = @{ rangeUpper = 17001; objectVersionDefault = 13237; objectVersionConfig = 16754 }
    "CU1"  = @{ rangeUpper = 17000; objectVersionDefault = 13236; objectVersionConfig = 16752 }
    "RTM"  = @{ rangeUpper = 17000; objectVersionDefault = 13236; objectVersionConfig = 16751 }
}

# Validate the provided TargetVersion
if (-not $ExchangeSchemaVersions.ContainsKey($TargetVersion)) {
    Write-Host "Invalid TargetVersion specified. Available versions are:"
    $ExchangeSchemaVersions.Keys | Sort-Object
    exit
}

# Retrieve the target schema versions
$target = $ExchangeSchemaVersions[$TargetVersion]

Write-Host "=== Checking Exchange AD Versions against $TargetVersion ===`n"

# 1. rangeUpper (Schema)
$schemaPath = "CN=ms-Exch-Schema-Version-Pt,CN=Schema,CN=Configuration," + (Get-ADRootDSE).rootDomainNamingContext
$rangeUpper = (Get-ADObject $schemaPath -Properties rangeUpper).rangeUpper
Write-Host "Schema version (rangeUpper): $rangeUpper" -ForegroundColor Cyan
if ($rangeUpper -eq $target.rangeUpper) {
    Write-Host "✓ Schema is up to date for $TargetVersion." -ForegroundColor Green
} else {
    Write-Host "✗ Schema is outdated. Expected: $($target.rangeUpper)" -ForegroundColor Red
}

# 2. objectVersion (Default) - Microsoft Exchange System Objects
$defaultPath = "CN=Microsoft Exchange System Objects," + (Get-ADRootDSE).defaultNamingContext
$objectVersionDefault = (Get-ADObject $defaultPath -Properties objectVersion).objectVersion
Write-Host "`nObjectVersion (Default naming context): $objectVersionDefault" -ForegroundColor Cyan
if ($objectVersionDefault -eq $target.objectVersionDefault) {
    Write-Host "✓ Default context is up to date for $TargetVersion." -ForegroundColor Green
} else {
    Write-Host "✗ Default context is outdated. Expected: $($target.objectVersionDefault)" -ForegroundColor Red
}

# 3. objectVersion (Configuration) - Auto-detect Exchange Organization container(s)
$orgBase = "CN=Microsoft Exchange,CN=Services,CN=Configuration," + (Get-ADRootDSE).rootDomainNamingContext
$orgContainers = Get-ADObject -Filter * -SearchBase $orgBase -SearchScope OneLevel -Properties objectVersion

Write-Host "`nObjectVersion (Configuration naming context):" -ForegroundColor Cyan

foreach ($org in $orgContainers) {
    $name = $org.Name
    $version = $org.objectVersion
    Write-Host "• $name : $version"
    if ($version -eq $target.objectVersionConfig) {
        Write-Host "  ✓ Configuration context is up to date for $TargetVersion." -ForegroundColor Green
    } else {
        Write-Host "  ✗ Configuration context is outdated. Expected: $($target.objectVersionConfig)" -ForegroundColor Red
    }
}


Write-Host "`n=== Check Complete ==="
