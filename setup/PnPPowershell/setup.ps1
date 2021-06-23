#####################################################################
#                   Site Designs Studio V2                          #
#                 Setup Script (PnP Powershell)                     #
#####################################################################
param(
    [string] $tenantName,
    [string] $sitePath = "SiteDesignsStudio",
    [string] $credentialsName = ""
)

$url = "https://$tenantName.sharepoint.com"
if ($credentialsName -ne "") {
    Write-Host "Using credentials $credentialsName"
    Connect-PnPOnline -Url $url -Credentials $credentialsName
} else {
    Connect-PnPOnline -Url $url -UseWebLogin
}

$tenantId = Get-PnPTenantId
Register-PnPManagementShellAccess -ShowConsentUrl -TenantName "$tenantName.sharepoint.com"
Invoke-PnPTenantTemplate -Path .\SiteDesignsStudio.pnp -Parameter @{TenantName=$tenantName; SitePath=$sitePath; TenantId=$tenantId}