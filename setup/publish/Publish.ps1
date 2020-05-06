#####################################################################
#                   Site Designs Studio V2                          #
#                    Publish Setup Package                          #
#####################################################################
# NOTES:                                                            #
# ------------------------------------------------------------------#
# This script has to be executed from the publish folder            #
#####################################################################

# Build the sppkg
Write-Host "Building SPFx solution package..." -ForegroundColor Yellow
gulp clean
gulp bundle --ship
gulp package-solution --ship
Write-Host "SPFx solution package built." -ForegroundColor Green


$pnpPackageFileName = "SiteDesignsStudio.pnp"
$pnpTemplateXmlFileName = "sdsv2-template.xml"

Write-Host "Building PnP setup package $pnpPackageFileName..." -ForegroundColor Yellow
# Copy the sppkg to /package folder
Copy-Item ../../sharepoint/solution/site-designs-studio-v2.sppkg ../package/site-designs-studio-v2.sppkg

# Copy the sppkg temporarily to the PnPPowershell/debug/package to rebuild .pnp file
if (!(test-path "../PnPPowershell/debug/package")) {
    mkdir ../PnPPowershell/debug/package | Out-Null
}
Copy-Item ../../sharepoint/solution/site-designs-studio-v2.sppkg ../PnPPowershell/debug/package/site-designs-studio-v2.sppkg | Out-Null

# Rebuild the .pnp file with the latest sppkg and tenant template
Set-Location ../PnPPowershell/debug/
$pnpTemplate = Read-PnPTenantTemplate -Path $pnpTemplateXmlFileName
Save-PnPTenantTemplate -Template $pnpTemplate -Out $pnpPackageFileName -Force
Move-Item $pnpPackageFileName .. -force | Out-Null
# Delete the temp package file
Remove-Item -force -recurse package
Write-Host "PnP setup package built." -ForegroundColor Green

Set-Location ../../publish