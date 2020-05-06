import { ISiteScript } from "../models/ISiteScript";

const SiteScriptsMockData: ISiteScript[] = [
	{
		Id: '0001',
		Title: 'Add common libraries',
		Description: 'Add the needed common libraries in the target site',
		Content: {
			"$schema": "schema.json",
			actions: [],
			bindata: {},
			version: 1
		},
		Version: 1
	},
	{
		Id: '0002',
		Title: 'Set Team Site Branding',
		Description: 'Customize the site branding specific for Team Sites',
		Content: {
			"$schema": "schema.json",
			actions: [
				{
					verb: 'addNavLink',
					url: 'http://localhost',
					displayName: 'Localhost',
					isWebRelative: null
				},
				{
					verb: 'applyTheme',
					themeName: 'My Theme'
				}
			],
			bindata: {},
			version: 1
		},
		Version: 1
	},
	{
		Id: '0003',
		Title: 'Set Default Team Site Navigation',
		Description: 'Customize the site branding specific for common Team Sites',
		Content: {
			"$schema": "schema.json",
			actions: [
				{
					verb: 'addNavLink',
					url: 'http://localhost',
					displayName: 'Localhost',
					isWebRelative: null
				},
				{
					verb: 'applyTheme',
					themeName: 'My Theme'
				}
			],
			bindata: {},
			version: 1
		},
		Version: 1
	},
	{
		Id: '0004',
		Title: 'Set Accounting Team Site Navigation',
		Description: 'Customize the site branding specific for Accounting Team Sites',
		Content: {
			"$schema":"schema.json",
			actions: [],
			bindata: {},
			version: 1
		},
		Version: 1
	},
	{
		Id: '0005',
		Title: 'Set Communication sites Navigation',
		Description: 'Customize the site navigation specific for Communication Sites',
		Content: {
			"$schema":"schema.json",
			actions: [],
			bindata: {},
			version: 1
		},
		Version: 1
	},
	{
		Id: '0006',
		Title: 'Set Communication Sites Branding',
		Description: 'Customize the site branding specific for Communication Sites',
		Content: {
			"$schema":"schema.json",
			actions: [],
			bindata: {},
			version: 1
		},
		Version: 1
	}
];

export default SiteScriptsMockData;