# Install or upgrade with PnP Powershell

This setup package is designed to either create a fresh new install or upgrade existing setup.

## Prerequisites

A recent version of PnP Powershell is needed to be installed
(It should support PnP Tenant Templates)

To install the latest version of PnP PowerShell, the following command can be used

First make sure to uninstall any edition of PnP PowerShell that runs on the Windows PowerShell version (the one with the blue background).

```powershell
Uninstall-Module -Name "SharePointPnPPowerShellOnline" -AllVersions -Force
```
```powershell
Install-Module -Name "PnP.PowerShell" -Force
```

## Run the setup script

The script `setup.ps1` is located [here](./PnPPowershell/setup.ps1)

### Setup script parameters

- tenantName: *REQUIRED* The name of the target tenant (e.g. `contoso` if the tenant URL is https://contoso.sharepoint.com)
- sitePath: *OPTIONAL* The URL path of the Site Designs Studio site (e.g. `SiteDesignsStudio` in https://contoso.sharepoint.com/sites/SiteDesignsStudio)
- credentialsName: *OPTIONAL* The name of the Generic credentials to used (if saved in Windows Credential Manager)

### Examples
```powershell
# Basic command
.\setup.ps1 -tenantName contoso
# Configuring the path of the site in the URL
.\setup.ps1 -tenantName contoso -sitePath SDStudio
# Using generic credentials registered in Windows Credential Manager
.\setup.ps1 -tenantName contoso -credentialsName MyCreds
```

>
> [!NOTE]
> Support for MFA using PnP PowerShell -WebLogin parameter has been retired
> since it was causing issues when applying the PnP templates.
> an alternative solution will be provided soon.
>

# Install with Office 365 CLI

__Coming soon__
