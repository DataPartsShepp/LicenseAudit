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
The CSV must have headers `Company` and `Tenant`.

```csv
Company,Tenant
Contoso,contoso.onmicrosoft.com
Fabrikam,fabrikam.onmicrosoft.com
```

## Configuration

### Exclusions (`exclude.txt`)
To exclude specific users (e.g., admin accounts or test users) from the audit, add their **UserPrincipalName** (email) to a file named `exclude.txt` located in the same folder as the script.

### License Exclusions (`licenseExclude.txt`)
To exclude specific license products (e.g., free trials or zero-cost licenses) from the report, add the exact **Product Name** to a file named `licenseExclude.txt` located in the same folder as the script.

## Output

The script generates an Excel file (default: `.\TenantLicenseAudit.xlsx`).

- **Columns A-C**: User Details (Display Name, UPN, Assigned Licenses).
- **Columns E-F**: License Summary (License Product Name, Count).