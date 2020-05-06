#####################################################################
#                   Site Designs Studio V2                          #
#                 Setup Script (PnP Powershell)                     #
#####################################################################
param(
    [string] $tenantName,
    [string] $sitePath = "SiteDesignsStudio",
    [string] $credentialsName = "",
    [boolean] $useMFA= $false
)

$url = "https://$tenantName.sharepoint.com"
if ($credentialName -ne "") {
    Write-Host "Using credentials $credentialsName"
    Connect-PnPOnline $url -Credentials $credentialsName
} elseif ($useMFA) {
    Write-Host "Using WebLogin"
    Connect-PnPOnline $url -UseWebLogin
} else {
    Connect-PnPOnline $url
}

$tenantId = Get-PnPTenantId
Apply-PnPTenantTemplate -Path .\SiteDesignsStudio.pnp -Parameter @{TenantName=$tenantName; SitePath=$sitePath; TenantId=$tenantId}