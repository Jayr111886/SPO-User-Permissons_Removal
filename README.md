# SPO-User-Permissons_Removal
# SharePoint Online User Permission Removal Script

This PowerShell script is designed to remove permissions for a specified user from all folders within the "Shared Documents" library across all SharePoint Online sites in a tenant. The script leverages the PnP PowerShell module to interact with SharePoint Online.

## Features

- **Remove Permissions for One User**: The script only removes permissions for the user specified by their email.
- **Report-Only Mode**: When set to `reportOnly = $true`, the script will generate a CSV report of what permissions **would** be removed without making any actual changes.
- **Commit Mode**: When set to `reportOnly = $false`, the script will commits changes and removes permissions for the specified user
- **Error Logging**: Any sites that fail during the process are logged in a separate CSV file with the reason for failure.
- **Supports Recursive Folder Scanning**: The script will search for unique permissions across all folders within the `Shared Documents` library.
  
## Prerequisites

- **PnP PowerShell**: Install the PnP PowerShell module for SharePoint Online version 1.3.0 or later.
  
  ```powershell
  Install-Module -Name PnP.PowerShell

## Running the Script

- Run in path
  ```powershell
  .\SharePointUserPermissionRemoval.ps1
