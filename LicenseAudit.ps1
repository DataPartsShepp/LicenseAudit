param (
    [string]$Tenant,
    [string]$File,
    [string]$OutputFile = ".\TenantLicenseAudit.xlsx"
)

# Ensure required modules are installed
$requiredModules = @('ImportExcel', 'Microsoft.Graph')

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Warning "The '$module' module is required."

        $install = Read-Host "Install it now? (Y/N)"
        if ($install -match '^[yY]') {
            Install-Module -Name $module -Scope CurrentUser -Force
        } else {
            Write-Error "Install the module first: Install-Module $module"
            exit
        }
    }
}

# Validate parameters
if ($Tenant -and $File) {
    Write-Error "Use either -Tenant or -File, not both."
    exit
}

# Build tenant list
if ($File) {
    $tenantList = Import-Csv $File
}
elseif ($Tenant) {
    $tenantList = @(
        [PSCustomObject]@{
            Company = "Default"
            Tenant  = $Tenant
        }
    )
}
else {
    $inputTenant = Read-Host "Enter tenant domain or tenant ID"
    $tenantList = @(
        [PSCustomObject]@{
            Company = "Default"
            Tenant  = $inputTenant
        }
    )
}

# Download SKU mapping
$csvUrl = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
$skuMapPath = "$env:TEMP\LicenseSkuMap.csv"

try {
    Invoke-WebRequest -Uri $csvUrl -OutFile $skuMapPath -UseBasicParsing -ErrorAction Stop
}
catch {
    Write-Warning "Could not download SKU map, using cached version if available."
}

# Build SKU lookup
$skuLookup = @{}
if (Test-Path $skuMapPath) {
    Import-Csv $skuMapPath | ForEach-Object {
        if ($_.GUID -and $_.'Product_Display_Name') {
            $skuLookup[$_.GUID] = $_.'Product_Display_Name'
        }
    }
}

# Load exclusions
$excludePath = Join-Path $PSScriptRoot "exclude.txt"
$excludedUsers = if (Test-Path $excludePath) {
    Get-Content $excludePath | Where-Object { $_.Trim() -ne "" }
} else { @() }

$licenseExcludePath = Join-Path $PSScriptRoot "licenseExclude.txt"
$excludedLicenses = if (Test-Path $licenseExcludePath) {
    Get-Content $licenseExcludePath | Where-Object { $_.Trim() -ne "" }
} else { @() }

# Prepare Excel output
if (Test-Path $OutputFile) {
    Remove-Item $OutputFile
}

# Fix auth instability
$env:AZURE_IDENTITY_DISABLE_WAM = "true"

foreach ($entry in $tenantList) {

    $company = $entry.Company
    $tenant = $entry.Tenant

    Write-Host "`nConnecting to $company ($tenant)..."

    # Clean previous session
    Disconnect-MgGraph -ErrorAction SilentlyContinue

    try {
        Connect-MgGraph `
            -TenantId $tenant `
            -Scopes "User.Read.All","Organization.Read.All" `
            -UseDeviceAuthentication `
            -ContextScope Process `
            -NoWelcome
    }
    catch {
        Write-Error "Login failed for $tenant"
        Write-Error $_
        continue
    }

    # Validate connection
    $context = Get-MgContext
    if (-not $context -or -not $context.Account) {
        Write-Warning "Authentication did not complete for $tenant"
        continue
    }

    try {
        $users = Get-MgUser -All `
            -Property "DisplayName,UserPrincipalName,AssignedLicenses" `
            -ConsistencyLevel eventual
    }
    catch {
        Write-Error "Failed to retrieve users for $tenant"
        Write-Error $_
        continue
    }

    $licensedUsers = $users | Where-Object {
        $_.AssignedLicenses.Count -gt 0 -and
        $_.UserPrincipalName -notin $excludedUsers
    }

    $exportData = $licensedUsers | ForEach-Object {

        $validLicenses = $_.AssignedLicenses | ForEach-Object {
            $skuId = $_.SkuId.ToString()

            $name = if ($skuLookup.ContainsKey($skuId)) {
                $skuLookup[$skuId]
            } else {
                "Unknown License ($skuId)"
            }

            if ($name -notin $excludedLicenses) { $name }
        }

        if ($validLicenses) {
            [PSCustomObject]@{
                DisplayName       = $_.DisplayName
                UserPrincipalName = $_.UserPrincipalName
                Licenses          = ($validLicenses -join "; ")
            }
        }
    }

    if ($exportData) {

        $exportData | Export-Excel `
            -Path $OutputFile `
            -WorksheetName $company `
            -AutoSize

        $summary = $exportData.Licenses -split "; " |
            Group-Object |
            Select-Object @{
                Name = 'License Product'
                Expression = { $_.Name }
            }, @{
                Name = 'Count'
                Expression = { $_.Count }
            } |
            Sort-Object Count -Descending

        $summary | Export-Excel `
            -Path $OutputFile `
            -WorksheetName $company `
            -StartColumn 5 `
            -AutoSize
    }
    else {
        Write-Host "No licensed users found for $company."
    }

    Disconnect-MgGraph
    Write-Host "Finished processing $company"
}

Write-Host "`nAll tenants processed. Report saved to: $OutputFile"