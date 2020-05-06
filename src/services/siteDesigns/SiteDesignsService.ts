import { ISiteScript, ISiteScriptContent } from '../../models/ISiteScript';
import { ISiteDesign, WebTemplate, ISiteDesignWithGrantedRights, ISiteDesignGrantedPrincipal } from '../../models/ISiteDesign';
import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration } from '@microsoft/sp-http';
import { assign, find } from '@microsoft/sp-lodash-subset';
import { ISiteScriptSchemaService, SiteScriptSchemaServiceKey } from '../siteScriptSchema/SiteScriptSchemaService';
import { getPrincipalTypeFromName } from '../../utils/spUtils';

export interface IGetSiteScriptFromWebOptions {
	includeBranding?: boolean;
	includeLists?: string[];
	includeRegionalSettings?: boolean;
	includeSiteExternalSharingCapability?: boolean;
	includeTheme?: boolean;
	includeLinksToExportedItems?: boolean;
}

export interface IGetSiteScriptFromExistingResourceResult {
	JSON: ISiteScriptContent;
	Warnings: string[];
}

export interface ISiteDesignsService {
	baseUrl: string;
	getSiteDesigns(): Promise<ISiteDesign[]>;
	getSiteDesign(siteDesignId: string): Promise<ISiteDesignWithGrantedRights>;
	saveSiteDesign(siteDesign: ISiteDesign): Promise<ISiteDesignWithGrantedRights>;
	deleteSiteDesign(siteDesign: ISiteDesign | string): Promise<void>;
	getSiteScripts(): Promise<ISiteScript[]>;
	getSiteScript(siteScriptId: string): Promise<ISiteScript>;
	saveSiteScript(siteScript: ISiteScript): Promise<ISiteScript>;
	deleteSiteScript(siteScript: ISiteScript | string): Promise<void>;
	applySiteDesign(siteDesignId: string, webUrl: string): Promise<void>;
	getSiteScriptFromList(listUrl: string): Promise<IGetSiteScriptFromExistingResourceResult>;
	getSiteScriptFromWeb(webUrl: string, options?: IGetSiteScriptFromWebOptions): Promise<IGetSiteScriptFromExistingResourceResult>;
}

export class SiteDesignsService implements ISiteDesignsService {
	private spHttpClient: SPHttpClient;
	private schemaService: ISiteScriptSchemaService;

	constructor(serviceScope: ServiceScope) {
		serviceScope.whenFinished(() => {
			this.spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
			this.schemaService = serviceScope.consume(SiteScriptSchemaServiceKey);
		});
	}

	public baseUrl: string = '/';

	private _getEffectiveUrl(relativeUrl: string): string {
		return (this.baseUrl + '//' + relativeUrl).replace(/\/{2,}/, '/');
	}

	private async _restRequest<TResponse>(url: string, params: any = null): Promise<TResponse> {
		const restUrl = this._getEffectiveUrl(url);
		const options: ISPHttpClientOptions = {
			body: JSON.stringify(params),
			headers: {
				'Content-Type': 'application/json;charset=utf-8',
				ACCEPT: 'application/json; odata.metadata=minimal',
				'ODATA-VERSION': '4.0'
			}
		};
		const httpResponse = await this.spHttpClient.post(restUrl, SPHttpClient.configurations.v1, options);
		if (httpResponse.status == 204) {
			return {} as TResponse;
		} else {
			return httpResponse.json() as any as TResponse;
		}
	}

	public async getSiteDesigns(): Promise<ISiteDesign[]> {
		try {
			const response = await this._restRequest<{ value: any[] }>(
				'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns'
			);
			return response.value as ISiteDesign[];
		} catch (error) {
			console.error("An error occured while trying to get site designs", error);
			return [];
		}
	}

	public async getSiteDesign(siteDesignId: string): Promise<ISiteDesignWithGrantedRights> {

		try {
			const siteDesign = await this._restRequest<ISiteDesignWithGrantedRights>(
				'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignMetadata',
				{ id: siteDesignId }
			);

			const rights = await this._restRequest<{ value: { PrincipalName: string; DisplayName: string; }[] }>("/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRights", { id: siteDesignId });
			const existingRightPrincipalNames: ISiteDesignGrantedPrincipal[] = rights.value.map(r => ({
				id: null,
				displayName: r.DisplayName,
				principalName: r.PrincipalName,
				type: getPrincipalTypeFromName(r.PrincipalName)
			})
			);
			siteDesign.grantedRightsPrincipals = existingRightPrincipalNames;
			return siteDesign;
		} catch (error) {
			console.error(`An error occured while trying to get site design ${siteDesignId}`, error);
			return null;
		}
	}

	public async deleteSiteDesign(siteDesign: ISiteDesign | string): Promise<void> {
		let id = typeof siteDesign === 'string' ? siteDesign as string : (siteDesign as ISiteDesign).Id;
		return this._restRequest<void>(
			'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteDesign',
			{ id: id }
		);
	}

	public async saveSiteDesign(siteDesign: ISiteDesign): Promise<ISiteDesignWithGrantedRights> {
		if (siteDesign.Id) {
			// Update
			await this._restRequest<ISiteDesign>(
				'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign',
				{
					updateInfo: {
						Id: siteDesign.Id,
						Title: siteDesign.Title,
						Description: siteDesign.Description,
						SiteScriptIds: siteDesign.SiteScriptIds,
						WebTemplate: siteDesign.WebTemplate.toString(),
						PreviewImageUrl: siteDesign.PreviewImageUrl,
						PreviewImageAltText: siteDesign.PreviewImageAltText,
						Version: siteDesign.Version,
						IsDefault: siteDesign.IsDefault
					}
				}
			);
		} else {
			// Create
			const created = await this._restRequest<ISiteDesign>(
				'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign',
				{
					info: {
						Title: siteDesign.Title,
						Description: siteDesign.Description,
						SiteScriptIds: siteDesign.SiteScriptIds,
						WebTemplate: siteDesign.WebTemplate.toString(),
						PreviewImageUrl: siteDesign.PreviewImageUrl,
						PreviewImageAltText: siteDesign.PreviewImageAltText
					}
				}
			);
			siteDesign.Id = created.Id;
		}

		const withGrantedRights = (siteDesign as ISiteDesignWithGrantedRights);
		if (withGrantedRights.grantedRightsPrincipals) {
			await this._setSiteDesignRights(siteDesign.Id, withGrantedRights.grantedRightsPrincipals);
		}

		return siteDesign;
	}

	private async _setSiteDesignRights(siteDesignId: string, principals: ISiteDesignGrantedPrincipal[]): Promise<void> {
		// Get the current rights of the site design
		const existingRights = await this._restRequest<{ value: { PrincipalName: string }[] }>("/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRights", { id: siteDesignId });
		const existingRightPrincipalNames: string[] = existingRights.value.map(r => r.PrincipalName);
		// Remove the ones not included in specified principalNames
		const toRevokePrincipalNames: string[] = existingRightPrincipalNames.filter(r => !find(principals, p => p.principalName == r));
		// Add the ones from principalNames not included in existing
		// Aliases must be set to be granted...
		const toGrantPrincipalNames: string[] = principals
			.filter(p => !!p.alias)
			.filter(p => existingRightPrincipalNames.indexOf(p.principalName) < 0)
			.map(p => p.alias);
		await this._restRequest("/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights",
			{
				id: siteDesignId,
				principalNames: toRevokePrincipalNames
			});
		await this._restRequest("/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GrantSiteDesignRights",
			{
				id: siteDesignId,
				principalNames: toGrantPrincipalNames,
				grantedRights: 1, // Means "View" , only supported value currently
			});
	}

	public async getSiteScripts(): Promise<ISiteScript[]> {
		const response = await this._restRequest<{ value: ISiteScript[] }>(
			'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts'
		);
		return response.value;
	}

	public async getSiteScript(siteScriptId: string): Promise<ISiteScript> {
		const siteScript = await this._restRequest<ISiteScript>(
			'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptMetadata',
			{ id: siteScriptId }
		);
		siteScript.Content = JSON.parse(siteScript.Content as any);
		return siteScript;
	}

	public async saveSiteScript(siteScript: ISiteScript): Promise<ISiteScript> {
		if (siteScript.Id) {
			await this._restRequest(
				'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteScript',
				{
					updateInfo: {
						Id: siteScript.Id,
						Title: siteScript.Title,
						Description: siteScript.Description,
						Version: siteScript.Version,
						Content: JSON.stringify(siteScript.Content)
					}
				}
			);
		} else {
			const created = await this._restRequest<ISiteScript>(
				`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title)?@title='${siteScript.Title}'`,
				siteScript.Content
			);
			siteScript.Id = created.Id;
		}
		return siteScript;
	}

	public deleteSiteScript(siteScript: ISiteScript | string): Promise<void> {
		let id = typeof siteScript === 'string' ? siteScript as string : (siteScript as ISiteScript).Id;
		return this._restRequest(
			'/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteScript',
			{ id: id }
		);
	}

	public applySiteDesign(siteDesignId: string, webUrl: string): Promise<void> {
		// TODO Implement
		return null;
	}

	public async getSiteScriptFromList(listUrl: string): Promise<IGetSiteScriptFromExistingResourceResult> {
		const response = await this._restRequest<{ value: string }>('/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptFromList', {
			listUrl
		});

		const defaultContent = this.schemaService.getNewSiteScript();
		const siteScriptContent = assign(defaultContent, JSON.parse(response.value)) as ISiteScriptContent;
		siteScriptContent.$schema = "schema.json";
		return { Warnings: [], JSON: siteScriptContent };
	}

	public async getSiteScriptFromWeb(webUrl: string, options?: IGetSiteScriptFromWebOptions): Promise<IGetSiteScriptFromExistingResourceResult> {
		const info = {};
		if (options) {
			if (options.includeBranding === true || options.includeBranding === false) {
				info["IncludeBranding"] = options.includeBranding;
			}
			if (options.includeLists !== null || typeof options.includeLists !== "undefined") {
				info["IncludedLists"] = options.includeLists;
			}
			if (options.includeRegionalSettings === true || options.includeRegionalSettings === false) {
				info["IncludeRegionalSettings"] = options.includeRegionalSettings;
			}
			if (options.includeSiteExternalSharingCapability === true || options.includeSiteExternalSharingCapability === false) {
				info["IncludeSiteExternalSharingCapability"] = options.includeSiteExternalSharingCapability;
			}
			if (options.includeTheme === true || options.includeTheme === false) {
				info["IncludeTheme"] = options.includeTheme;
			}
			if (options.includeLinksToExportedItems === true || options.includeLinksToExportedItems === false) {
				info["IncludeLinksToExportedItems"] = options.includeLinksToExportedItems;
			}
		}
		const response = await this._restRequest<{ JSON: string, Warnings: string[] }>('/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptFromWeb', {
			webUrl,
			info
		});

		const defaultContent = this.schemaService.getNewSiteScript();
		const siteScriptContent = assign(defaultContent, JSON.parse(response.JSON)) as ISiteScriptContent;
		siteScriptContent.$schema = "schema.json";
		return { Warnings: response.Warnings, JSON: siteScriptContent };
	}
}

export const SiteDesignsServiceKey = ServiceKey.create<ISiteDesignsService>(
	'YPCODE:SDSv2:SiteDesignsService',
	SiteDesignsService
);
