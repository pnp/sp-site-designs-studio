
# Site Designs Studio V2 Changelog

## v2.0.1

- New Feature - Create a site script from PnP sample
- Improvement - Updated built-in Site Script schema
- Improvement - polishing Setup docs

## v2.0.0

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