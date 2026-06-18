# Microsoft 365 License Audit Script

This PowerShell script audits Microsoft 365 licenses for one or multiple tenants. It retrieves licensed users, translates technical License SKU IDs into human-readable product names (e.g., "Office 365 E3"), and exports the data to an Excel workbook with a summary tally of licenses used.

## Features

- **Dynamic SKU Mapping**: Automatically downloads the latest Microsoft Product Name to GUID mapping CSV to ensure license names are up to date.
- **Exclusion List**: Supports an `exclude.txt` file to omit specific service accounts or users from the report.
- **License Exclusion**: Supports a `licenseExclude.txt` file to omit specific license types (e.g., "Microsoft Power Automate Free") from the report.
- **Multi-Tenant Support**: Can process a single tenant or a batch of tenants via CSV.
- **Excel Reporting**:
  - Creates a separate worksheet for each tenant.
  - Lists individual users and their assigned licenses.
  - Generates a summary table (Count of each license type) on the same sheet (starting at Column E).

## Prerequisites

Ensure the following PowerShell modules are installed:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser
```

You will need an account with permissions to read user and organization data (specifically `User.Read.All` and `Organization.Read.All` scopes) for the tenants you are auditing.

## Usage

### 1. Interactive Mode
Run the script without parameters. You will be prompted to enter a Tenant Domain or ID.

```powershell
.\LicenseAudit.ps1
```

### 2. Single Tenant Mode
Specify the tenant directly via the command line.

```powershell
.\LicenseAudit.ps1 -Tenant "contoso.onmicrosoft.com"
```

### 3. Bulk Mode (Multiple Tenants)
Provide a CSV file containing a list of tenants to process.

```powershell
.\LicenseAudit.ps1 -File ".\tenants.csv"
```

**CSV Format for `-File`:**
The CSV must have headers `Company` and `Tenant`. Optionally, add an `AppName` column to specify which app credential to use per tenant.

```csv
Company,Tenant,AppName
Contoso,contoso.onmicrosoft.com,app1
Fabrikam,fabrikam.onmicrosoft.com,app2
Adventure Works,adventureworks.onmicrosoft.com,
```

**Tip:** A template file `tenants.csv.example` is provided. Copy it to `tenants.csv` and customize it with your tenant information.

## Authentication Methods

### Device Authentication (Default)
If no app credentials are provided, the script falls back to device authentication (interactive login).

```powershell
.\LicenseAudit.ps1 -Tenant "contoso.onmicrosoft.com"
```

You will be prompted to authenticate via a device code in your browser.

### Application Authentication (App Keys)
For automated or unattended runs, use an Azure AD application with client credentials. This method takes precedence if provided and falls back to device authentication if it fails.

#### Prerequisites
1. Create an Azure AD Application in your tenant
2. Grant it `User.Read.All` and `Organization.Read.All` application permissions
3. Create a client secret

#### Option 1: Command-Line Parameters

```powershell
.\LicenseAudit.ps1 `
  -Tenant "contoso.onmicrosoft.com" `
  -ClientId "<app-client-id>" `
  -ClientSecret "<app-client-secret>" `
  -AppTenantId "<app-tenant-id>"
```

#### Option 2: Single App Configuration File

Create an `app-credentials.json` file with a single app (simple setup):

```json
{
  "default": {
    "ClientId": "12345678-1234-1234-1234-123456789abc",
    "ClientSecret": "abc~defghijklmnopqrstuvwxyz",
    "AppTenantId": "abcdef12-3456-7890-abcd-ef1234567890"
  }
}
```

Then run the script:

```powershell
.\LicenseAudit.ps1 -Tenant "contoso.onmicrosoft.com"
```

#### Option 3: Multiple Apps Configuration File (Recommended)

For multi-tenant audits with different apps per tenant, create an `app-credentials.json` file:

```json
{
  "default": {
    "ClientId": "default-app-id",
    "ClientSecret": "default-secret",
    "AppTenantId": "default-tenant-id"
  },
  "apps": {
    "app1": {
      "ClientId": "app1-client-id",
      "ClientSecret": "app1-secret",
      "AppTenantId": "app1-tenant-id"
    },
    "app2": {
      "ClientId": "app2-client-id",
      "ClientSecret": "app2-secret",
      "AppTenantId": "app2-tenant-id"
    }
  },
  "tenantMappings": {
    "contoso.onmicrosoft.com": "app1",
    "fabrikam.onmicrosoft.com": "app2"
  }
}
```

**How it works:**
- `default`: Used if no mapping or app is specified
- `apps`: Named app credential sets
- `tenantMappings`: Maps tenant domains to app names (automatic app selection)

**Run with automatic app selection:**

```powershell
.\LicenseAudit.ps1 -File ".\tenants.csv"
```

The script automatically selects the correct app for each tenant based on `tenantMappings`.

#### Option 4: Specify App via CSV Column

For even more flexibility, add an `AppName` column to your CSV:

```csv
Company,Tenant,AppName
Contoso,contoso.onmicrosoft.com,app1
Fabrikam,fabrikam.onmicrosoft.com,app2
```

Run the script:

```powershell
.\LicenseAudit.ps1 -File ".\tenants.csv"
```

The script uses the `AppName` from each row if provided, otherwise falls back to `tenantMappings`.

#### Option 5: Override via Command-Line

Command-line parameters override everything:

```powershell
# Use specific app from config
.\LicenseAudit.ps1 -Tenant "contoso.onmicrosoft.com" -AppName "app1"

# Use command-line credentials (highest priority)
.\LicenseAudit.ps1 `
  -Tenant "contoso.onmicrosoft.com" `
  -ClientId "override-id" `
  -ClientSecret "override-secret" `
  -AppTenantId "override-tenant-id"
```

**Credential Resolution Order:**
1. Command-line parameters (`-ClientId`, `-ClientSecret`, `-AppTenantId`)
2. CSV `AppName` column (if processing multiple tenants from file)
3. Command-line `-AppName` parameter
4. `tenantMappings` lookup
5. `default` credentials from config file
6. Fall back to device authentication

**Security:** The actual `app-credentials.json` file is excluded from version control (see `.gitignore`). Keep your credentials secure and never commit them to the repository.

## Configuration

### Exclusions (`exclude.txt`)
To exclude specific users (e.g., admin accounts or test users) from the audit, add their **UserPrincipalName** (email) to a file named `exclude.txt` located in the same folder as the script.

Example `exclude.txt`:
```
admin@contoso.onmicrosoft.com
serviceprincipal@contoso.onmicrosoft.com
test.user@contoso.onmicrosoft.com
```

**Tip:** A template file `exclude.txt.example` is provided. Copy it to `exclude.txt` and customize it with your excluded users.

### License Exclusions (`licenseExclude.txt`)
To exclude specific license products (e.g., free trials or zero-cost licenses) from the report, add the exact **Product Name** to a file named `licenseExclude.txt` located in the same folder as the script.

## Output

The script generates an Excel file (default: `.\TenantLicenseAudit.xlsx`).

- **Columns A-C**: User Details (Display Name, UPN, Assigned Licenses).
- **Columns E-F**: License Summary (License Product Name, Count).