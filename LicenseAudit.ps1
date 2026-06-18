param (
    [string]$Tenant,
    [string]$File,
    [string]$OutputFile = ".\TenantLicenseAudit.xlsx",
    [string]$ClientId,
    [string]$ClientSecret,
    [string]$AppTenantId,
    [string]$CredentialFile = ".\app-credentials.json",
    [string]$AppName
)

# Load app credentials from JSON file if it exists and parameters not provided
$configPath = Join-Path $PSScriptRoot $CredentialFile
$appConfig = @{}

if (Test-Path $configPath) {
    try {
        $config = Get-Content $configPath | ConvertFrom-Json
        $appConfig = $config
        Write-Host "Loaded app credentials from $configPath" -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Could not load credentials from $configPath: $_"
    }
}

# Helper function to resolve credentials for a tenant
function Get-AppCredentials {
    param(
        [string]$TenantId,
        [string]$AppName,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$AppTenantId,
        [object]$AppConfig
    )
    
    $creds = @{
        ClientId = $ClientId
        ClientSecret = $ClientSecret
        AppTenantId = $AppTenantId
    }
    
    # If command-line parameters are provided, use them
    if ($ClientId -and $ClientSecret -and $AppTenantId) {
        return $creds
    }
    
    # If AppName is specified, use that app
    if ($AppName -and $AppConfig.apps -and $AppConfig.apps.$AppName) {
        $appCreds = $AppConfig.apps.$AppName
        return @{
            ClientId = $appCreds.ClientId
            ClientSecret = $appCreds.ClientSecret
            AppTenantId = $appCreds.AppTenantId
        }
    }
    
    # Check if there's a mapping for this tenant
    if ($AppConfig.tenantMappings -and $AppConfig.tenantMappings.$TenantId) {
        $mappedAppName = $AppConfig.tenantMappings.$TenantId
        if ($AppConfig.apps -and $AppConfig.apps.$mappedAppName) {
            $appCreds = $AppConfig.apps.$mappedAppName
            return @{
                ClientId = $appCreds.ClientId
                ClientSecret = $appCreds.ClientSecret
                AppTenantId = $appCreds.AppTenantId
            }
        }
    }
    
    # Fall back to default credentials
    if ($AppConfig.default) {
        return @{
            ClientId = $AppConfig.default.ClientId
            ClientSecret = $AppConfig.default.ClientSecret
            AppTenantId = $AppConfig.default.AppTenantId
        }
    }
    
    # Return empty if no credentials found
    return $creds
}

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

# Helper function to authenticate with fallback
function Connect-ToTenant {
    param(
        [string]$TenantId,
        [string]$AppName,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$AppTenantId,
        [object]$AppConfig
    )
    
    # Resolve which credentials to use
    $creds = Get-AppCredentials -TenantId $TenantId -AppName $AppName -ClientId $ClientId `
        -ClientSecret $ClientSecret -AppTenantId $AppTenantId -AppConfig $AppConfig
    
    # Try app credentials first if available
    if ($creds.ClientId -and $creds.ClientSecret -and $creds.AppTenantId) {
        try {
            Write-Host "Attempting to connect using app credentials..."
            $credential = New-Object System.Management.Automation.PSCredential(
                $creds.ClientId,
                (ConvertTo-SecureString -String $creds.ClientSecret -AsPlainText -Force)
            )
            
            Connect-MgGraph `
                -TenantId $TenantId `
                -ClientSecretCredential $credential `
                -Scopes "User.Read.All","Organization.Read.All" `
                -ContextScope Process `
                -NoWelcome
            
            Write-Host "Connected using app credentials." -ForegroundColor Green
            return $true
        }
        catch {
            Write-Warning "App credential authentication failed: $_"
            Write-Host "Falling back to device authentication..."
        }
    }
    
    # Fall back to device authentication
    try {
        Connect-MgGraph `
            -TenantId $TenantId `
            -Scopes "User.Read.All","Organization.Read.All" `
            -UseDeviceAuthentication `
            -ContextScope Process `
            -NoWelcome
        
        Write-Host "Connected using device authentication." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Device authentication failed: $_"
        return $false
    }
}

foreach ($entry in $tenantList) {

    $company = $entry.Company
    $tenant = $entry.Tenant
    $entryAppName = $entry.AppName  # Optional column in CSV

    Write-Host "`nConnecting to $company ($tenant)..."

    # Clean previous session
    Disconnect-MgGraph -ErrorAction SilentlyContinue

    # Connect with fallback logic - use AppName from CSV if available, otherwise use script parameter
    $selectedAppName = if ($entryAppName) { $entryAppName } else { $AppName }
    $connectResult = Connect-ToTenant -TenantId $tenant -AppName $selectedAppName -ClientId $ClientId `
        -ClientSecret $ClientSecret -AppTenantId $AppTenantId -AppConfig $appConfig
    
    if (-not $connectResult) {
        Write-Error "Failed to authenticate for $tenant"
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