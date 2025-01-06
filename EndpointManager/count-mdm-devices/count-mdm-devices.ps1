<#
.SYNOPSIS
  Counts, per tenant (domain-based), the fully Intune-managed devices (MDM) and exports results to CSV.

.DESCRIPTION
  - Iterates through a list of tenant domains (e.g., contoso.onmicrosoft.com).
  - Uses an interactive MFA login (Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All").
  - Queries /deviceManagement/managedDevices where "managementAgent eq 'mdm'".
  - Exports:
      1) A full device list (AllDevices.csv).
      2) A per-tenant summary (TenantSummary.csv).
  - Provides verbose feedback in the console.

.NOTES
  - Make sure you have the Microsoft.Graph.Intune module (and its dependencies) properly installed.
  - If you still run into "[Assembly with same name is already loaded]" errors, try:
    Get-Module Microsoft.Graph.* | Remove-Module -Force
    or run PowerShell with -NoProfile.
#>

# ---------------------------------
# (A) Attempt to remove Graph modules from the current session
# ---------------------------------

Write-Host "Attempting to remove any loaded Microsoft.Graph modules to avoid assembly conflicts..."
try {
    Get-Module Microsoft.Graph.* | ForEach-Object {
        Remove-Module $_.Name -Force -ErrorAction SilentlyContinue
    }
    Write-Host "Successfully removed previously loaded Microsoft.Graph modules from this session."
}
catch {
    Write-Warning "Could not remove some loaded Graph modules. Continuing anyway. Error: $($_.Exception.Message)"
}

# ---------------------------------
# 1) Variables / Parameters
# ---------------------------------

# List of your tenant domains - adapt as needed
$tenantDomains = @(
    "aaaaaa.com",
    "bbbbbb.com",
    "cccccc.com"
    # add more if needed...
)

# CSV export paths
$csvPathDevices  = ".\AllDevices.csv"
$csvPathSummary  = ".\TenantSummary.csv"

# ---------------------------------
# 2) Module check & Import
# ---------------------------------

Write-Host "Checking if 'Microsoft.Graph.Intune' is installed..."

try {
    $moduleCheck = Get-Module -ListAvailable -Name "Microsoft.Graph.Intune"
    if (-not $moduleCheck) {
        Write-Warning "'Microsoft.Graph.Intune' is not installed. Attempting installation..."
        Install-Module Microsoft.Graph.Intune -Scope CurrentUser -Force
        Write-Host "Module installed successfully."
    } 
    else {
        Write-Host "'Microsoft.Graph.Intune' is already installed."
    }
}
catch {
    Write-Error "Error checking/installing the module: $($_.Exception.Message)"
    return
}

try {
    Write-Host "Importing 'Microsoft.Graph.Intune'..."
    Import-Module Microsoft.Graph.Intune -ErrorAction Stop
    Write-Host "'Microsoft.Graph.Intune' imported successfully."
}
catch {
    Write-Error "Error importing 'Microsoft.Graph.Intune': $($_.Exception.Message)"
    return
}

# ---------------------------------
# 3) Main logic
# ---------------------------------

# Prepare a list to store devices from all tenants
$allDevices = New-Object System.Collections.Generic.List[System.Object]

Write-Host "Starting query across $($tenantDomains.Count) tenant(s)..."

# Initialize progress
$tenantIndex = 0
$totalTenants = $tenantDomains.Count

foreach ($domain in $tenantDomains) {

    $tenantIndex++
    $percentComplete = [int](($tenantIndex / $totalTenants) * 100)

    Write-Progress -Activity "Connecting to Tenant $domain ($tenantIndex of $totalTenants)" `
                   -Status "Signing in & retrieving 'mdm' devices" `
                   -PercentComplete $percentComplete

    Write-Host "`n-------------------------------------------"
    Write-Host "Processing Tenant: $domain"

    # Disconnect any existing session before connecting again
    Disconnect-MgGraph 2>$null

    Write-Host "Initiating interactive MFA login..."
    try {
        Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All" -TenantId $domain | Out-Null
        Write-Host "Successfully connected."
    }
    catch {
        Write-Error "Error connecting to Microsoft Graph for Tenant '${domain}': $($_.Exception.Message)"
        continue
    }

    # Retrieve only fully managed (MDM) devices
    $filterString = "managementAgent eq 'mdm'"
    Write-Host "Retrieving devices for tenant '$domain' with filter: $filterString ..."

    try {
        # -All fetches all pages of results
        $tenantDevices = Get-MgDeviceManagementManagedDevice -Filter $filterString -All
    }
    catch {
        Write-Error "Error retrieving devices for '${domain}': $($_.Exception.Message)"
        continue
    }

    if (-not $tenantDevices) {
        Write-Warning "No devices found for '$domain' or empty results."
    }
    else {
        Write-Host "Found $($tenantDevices.Count) 'mdm' devices in tenant '$domain'."

        foreach ($device in $tenantDevices) {
            $obj = [PSCustomObject]@{
                Tenant               = $domain
                DeviceId            = $device.Id
                DeviceName          = $device.DeviceName
                OperatingSystem     = $device.OperatingSystem
                OsVersion           = $device.OsVersion
                Manufacturer        = $device.Manufacturer
                Model               = $device.Model
                ManagementAgent     = $device.ManagementAgent
                ComplianceState     = $device.ComplianceState
                LastCheckInDateTime = $device.LastSyncDateTime
            }
            $allDevices.Add($obj) | Out-Null
        }
    }

    # Optionally disconnect after each tenant
    Disconnect-MgGraph 2>$null
    Write-Host "Tenant '$domain' processing finished."
    Write-Host "-------------------------------------------`n"
}

Write-Host "Finished querying all tenants. Total number of fully managed devices (MDM): $($allDevices.Count)"

# ---------------------------------
# 4) CSV Output
# ---------------------------------

Write-Host "Exporting device list to: $csvPathDevices"
try {
    $allDevices | Export-Csv -Path $csvPathDevices -NoTypeInformation -Encoding UTF8
    Write-Host "Device list successfully exported to '$csvPathDevices'."
}
catch {
    Write-Error "Error exporting the device list: $($_.Exception.Message)"
}

Write-Host "Creating per-tenant summary..."

$summary = $allDevices | Group-Object -Property Tenant | Select-Object Name, Count

Write-Host "Exporting summary to: $csvPathSummary"
try {
    $summary | Export-Csv -Path $csvPathSummary -NoTypeInformation -Encoding UTF8
    Write-Host "Summary successfully exported to '$csvPathSummary'."
}
catch {
    Write-Error "Error exporting the summary: $($_.Exception.Message)"
}

Write-Host "Script execution finished. Have a nice day!"
