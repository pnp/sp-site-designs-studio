#####################################################################
#                   Site Designs Studio V2                          #
#                 Setup Script (PnP Powershell)                     #
#####################################################################
#                       !!!!IMPORTANT!!!!
#####################################################################
# This script has not been tested yet
# It is not ready for usage just yet
#####################################################################
param(
    [string] $tenantName,
    [string] $sitePath = "sds-cli2",
    [string] $credentialsName = ""
)

$tenantSpUrl = "https://$tenantName.sharepoint.com"
$sdsSiteUrl = "$tenantSpUrl/sites/$sitePath"

# o365 login

$status = (o365 status -o json) | ConvertFrom-Json
# $currentUser is kept to add permission for it
$currentUser = $status.connectedAs

# Create a modern communication site
# o365 spo site add --type CommunicationSite -u $sdsSiteUrl -t "Site Designs Studio" -l 1033 --owners $currentUser
o365 spo site add --type CommunicationSite -u $sdsSiteUrl -t "Site Designs Studio" -l 1033

# TODO Add owner to created site

# Create a library in it to store site designs preview images.
o365 spo list add --baseTemplate DocumentLibrary -t "Site Design - Preview Images" --description "A library to store the Site Designs preview images" 

# Upload and deploy the SPFx application in the app catalog
o365 spo app add --filePath ..\package\site-designs-studio-v2.sppkg
spo app deploy --name site-designs-studio-v2.sppkg --skipFeatureDeployment

# TODO Approve the Microsoft Graph permission

# Add the page programmatically to the new site
$webPartDataJson = '{
    "id": "e164cc97-dcae-4a4e-a899-67ebb916207e",
    "instanceId": "c2a07e0a-2080-4912-a065-e957377cbc3b",
    "title": "Site Designs Studio",
    "description": "Site Designs Studio",
    "dataVersion": "2.0",
    "properties": {
      "description": "Site Designs Studio"
    },
    "dynamicDataPaths": {},
    "dynamicDataValues": {},
    "serverProcessedContent": {
      "htmlStrings": {},
      "searchablePlainTexts": {},
      "imageSources": {},
      "links": {}
    }
  }'
$webPartDataEscapedJson = '`"{0}"`' -f $webPartDataJson.Replace('\', '\\').Replace('"', '""')   
spo apppage add --title "Site Design Studio" --webUrl $sdsSiteUrl --webPartData $webPartDataEscapedJson

# Set as home page
# MISSING, possibility to change the URL (to be consistent with the PnP PowerShell setup => page name is sds2.aspx)
spo page set --name page.aspx --webUrl $sdsSiteUrl --promoteAs HomePage