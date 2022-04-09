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
    Connect-PnPOnline $url -Credentials $credentialsName
} else {
    Connect-PnPOnline $url
}

$tenantId = Get-PnPTenantId
Invoke-PnPTenantTemplate -Path .\SiteDesignsStudio.pnp -Parameter @{TenantName=$tenantName; SitePath=$sitePath; TenantId=$tenantId}