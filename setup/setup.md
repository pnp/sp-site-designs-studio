# Install or upgrade with PnP Powershell

This setup package is designed to either create a fresh new install or upgrade existing setup.

## Prerequisites

A recent version of PnP Powershell is needed to be installed
(It should support PnP Tenant Templates)

To install the latest version of PnP PowerShell, the following command can be used

```powershell
Install-Module SharePointPnPPowershellOnline -Force
```

## Run the setup script

The script `setup.ps1` is located [here](./PnPPowershell/setup.ps1)

### Setup script parameters

- tenantName: *REQUIRED* The name of the target tenant (e.g. `contoso` if the tenant URL is https://contoso.sharepoint.com)
- sitePath: *OPTIONAL* The URL path of the Site Designs Studio site (e.g. `SiteDesignsStudio` in https://contoso.sharepoint.com/sites/SiteDesignsStudio)
- credentialsName: *OPTIONAL* The name of the Generic credentials to used (if saved in Windows Credential Manager)
- useMFA: *OPTIONAL* A flag indicating whether to use web login in the case MFA is enabled for the used account. Will be ignored if `credentialsName` is used

### Examples
```powershell
# Basic command
.\setup.ps1 -tenantName contoso
# Configuring the path of the site in the URL
.\setup.ps1 -tenantName contoso -sitePath SDStudio
# Using MFA
.\setup.ps1 -tenantName contoso -useMFA
# Using generic credentials registered in Windows Credential Manager
.\setup.ps1 -tenantName contoso -credentialsName MyCreds
```

# Install with Office 365 CLI

__Coming soon__