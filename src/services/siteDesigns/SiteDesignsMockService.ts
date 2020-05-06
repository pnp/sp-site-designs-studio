import { ISiteScript, ISiteScriptContent } from "../../models/ISiteScript";
import { ISiteDesign } from "../../models/ISiteDesign";
import { assign, findIndex } from "@microsoft/sp-lodash-subset";
import SiteDesignsMockData from "../../mock/SiteDesignsMock";
import SiteScriptsMockData from "../../mock/SiteScriptsMock";
import { ServiceScope } from "@microsoft/sp-core-library";
import { ISiteDesignsService, IGetSiteScriptFromWebOptions, IGetSiteScriptFromExistingResourceResult } from "./SiteDesignsService";

const MockServiceDelay = 20;

export class MockSiteDesignsService implements ISiteDesignsService {
	constructor(serviceScope: ServiceScope) { }

	public baseUrl: string = '';

	private _getSiteDesign(siteDesignId: string): ISiteDesign {
		let index = findIndex(SiteDesignsMockData, (sd) => sd.Id == siteDesignId);
		return index >= 0 ? SiteDesignsMockData[index] : null;
	}

	private _getSiteScript(siteScriptId: string): ISiteScript {
		let index = findIndex(SiteScriptsMockData, (sd) => sd.Id == siteScriptId);
		return index >= 0 ? SiteScriptsMockData[index] : null;
	}

	private _dumpInMemoryDatabase() {
		console.log("Site Designs DB =", SiteDesignsMockData);
		console.log("Site Scripts DB =", SiteScriptsMockData);
	}

	private _mockServiceCall<T>(action: () => T): Promise<T> {
		return new Promise<T>((resolve, reject) => {
			setTimeout(() => {
				let result: T = action();
				this._dumpInMemoryDatabase();
				resolve(result);
			}, MockServiceDelay);
		});
	}

	public getSiteDesigns(): Promise<ISiteDesign[]> {
		return this._mockServiceCall(() => SiteDesignsMockData);
	}
	public getSiteDesign(siteDesignId: string): Promise<ISiteDesign> {
		return this._mockServiceCall(() => this._getSiteDesign(siteDesignId));
	}
	public saveSiteDesign(siteDesign: ISiteDesign): Promise<ISiteDesign> {
		const action = () => {
			if (siteDesign.Id) {
				// Update
				const existing = this._getSiteDesign(siteDesign.Id);
				assign(existing, siteDesign);
				return siteDesign;
			} else {
				// Create
				siteDesign.Id = (+new Date()).toString();
				SiteDesignsMockData.push(siteDesign);
				return siteDesign;
			}
		};

		return this._mockServiceCall(action);
	}
	public deleteSiteDesign(siteDesign: ISiteDesign | string): Promise<void> {
		const action = () => {
			let id = typeof siteDesign === 'string' ? siteDesign as string : (siteDesign as ISiteDesign).Id;
			const existing = this._getSiteDesign(id);
			let index = SiteDesignsMockData.indexOf(existing);
			SiteDesignsMockData.splice(index, 1);
		};
		return this._mockServiceCall(action);
	}
	public getSiteScripts(): Promise<ISiteScript[]> {
		return this._mockServiceCall(() => SiteScriptsMockData);
	}
	public getSiteScript(siteScriptId: string): Promise<ISiteScript> {
		return this._mockServiceCall(() => this._getSiteScript(siteScriptId));
	}

	public saveSiteScript(siteScript: ISiteScript): Promise<ISiteScript> {
		const action = () => {
			if (siteScript.Id) {
				// Update
				const existing = this._getSiteScript(siteScript.Id);
				assign(existing, siteScript);
				return siteScript;
			} else {
				// Create
				siteScript.Id = (+new Date()).toString();
				SiteScriptsMockData.push(siteScript);
				return siteScript;
			}
		};
		return this._mockServiceCall(action);
	}
	public deleteSiteScript(siteScript: ISiteScript | string): Promise<void> {
		const action = () => {
			let id = typeof siteScript === 'string' ? siteScript as string : (siteScript as ISiteScript).Id;
			const existing = this._getSiteScript(id);
			let index = SiteScriptsMockData.indexOf(existing);
			SiteScriptsMockData.splice(index, 1);
		};
		return this._mockServiceCall(action);
	}

	public applySiteDesign(siteDesignId: string, webUrl: string): Promise<void> {
		return Promise.resolve();
	}
	public getSiteScriptFromList(listUrl: string): Promise<IGetSiteScriptFromExistingResourceResult> {
		return Promise.resolve({
			JSON: SiteScriptsMockData[0].Content,
			Warnings: ["This is a fake exported site script from mock data"]
		});
	}
	public getSiteScriptFromWeb(webUrl: string, options?: IGetSiteScriptFromWebOptions): Promise<IGetSiteScriptFromExistingResourceResult> {
		return Promise.resolve({
			JSON: SiteScriptsMockData[0].Content,
			Warnings: ["This is a fake exported site script from mock data"]
		});
	}
}
