# Site Designs Studio V2

The Site Design Studio v2 is a new version of the solution.
It is designed to be used as an entire solution on a tenant to provision the customization and configuration to the sites allowing its users to create and manage Site Designs and Site Scripts without the need to write any PowerShell nor JSON.

## Setup

The solution is shipped as an application hosted in its own site collection.
It can be installed using a setup PowerShell script (it requires PnP PowerShell to be installed)

The setup guide is [here](./setup/setup.md)

> A setup script using Office 365 CLI will be released as well soon

## Features

The Site Designs Studio V2 comes with the following features
- Create, Update and Delete Site Designs on the tenant.
- Grant and revoke rights to Site Designs to specific principals (users or mail-enabled groups).
- Upload and select a preview image of Site Designs, when uploaded, the image will be stored in the Site Designs Studio site collection.
- Create, Update and Delete Site Scripts on the tenant.
- Edit the content of Site Scripts from a graphical user interface without the need of writing JSON. The JSON is automatically generated and displayed side by side with the graphical representation of the actions in a script.
- Site Actions and Sub Site Actions are sortable by drag and drop interaction in the user interface.
- For several actions that require more technical arguments (like IDs), the UI offers a user-friendly picker
  - Themes can be selected from the tenant available themes.
  - Apps can be selected from the available apps in the tenant app catalog.
  - Hub Sites can be selected from a list of available hub sites.
  - List base templates can be selected from a list of the supported list base templates.
- A Site Script can be created from scratch. (`Blank` option)
- A Site Script can be created from an existing site in the tenant. (`From Site` option)
  - The site to create the script from can be picked through a search box or directly specified by its URL.
  - The options to generate the script (e.g. The lists to include, the branding settings, sharing options... are configurable through the UI)
- A Site Script can be created from an existing list in the tenant. (`From List` option)
  - The site in which the list is located can be picked through a search box or directly specified by its URL.
  - When the site is selected, a dropdown allows the select the list to generated the Site Script from.
- When a new Site Script is created, the UI allows the user to directly associate it to either an existing Site Design or to a new Site Design.
- From the action button, a Site Script can be exported as package to be deployed to another tenant
  - The export package comes as several flavors (JSON, PnP PowerShell, O365 CLI PowerShell, O365 CLI Bash)

### Create a Site Script from an existing site
![Create from Site](docs/sdsv2_demo01.gif)

### Save a Site Script and associate it to a new Site Design
![Save a Site Script and associate it to a new Site Design](docs/sdsv2_demo03.gif)

### Export a Site Script to deploy it on another tenant
![Export Site Script](docs/sdsv2_export.gif)

## Important to know

The application will work in "read-only" for any users having access to the site. However, to be able to save and edit Site Designs and Site Scripts, the users **must** have the SharePoint global administrator privileges !

## API Permissions

In order to allow the lookup of users and groups in the current tenant, The permission `Directory.AccessAsUser.All` to Microsoft Graph API is requested.
CAUTION: 
The solution is installed globally on the tenant, it means this permission, when granted, will be granted to the whole tenant.
However, it will allow the application to only see the groups and users the current user is allowed to see. Thus, it should not cause any security breach.
It will only be needed for the ability to grant Site Designs to specific principals. The other features will be working seamlessly without this granted permission 
