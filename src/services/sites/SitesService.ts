import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import { PageContext } from "@microsoft/sp-page-context";
import { sp, ISearchQuery, SearchResults, Web } from "@pnp/sp/presets/all";
import { ISPSite } from '../../models/ISPSite';
import { IList } from '../../models/IList';

export interface ISitesService {
    getSiteByNameOrUrl(nameOrUrl: string): Promise<ISPSite[]>;
    getSiteLists(siteUrl: string): Promise<IList[]>;
}

class SitesService implements ISitesService {

    private pageContext: PageContext = null;
    constructor(serviceScope: ServiceScope) {

        serviceScope.whenFinished(() => {
            this.pageContext = serviceScope.consume(PageContext.serviceKey);
            sp.setup({
                sp: {
                    baseUrl: this.pageContext.web.absoluteUrl
                }
            });
        });
    }

    public async getSiteByNameOrUrl(nameOrUrl: string): Promise<ISPSite[]> {
        const searchResults: SearchResults = await sp.search(<ISearchQuery>{
            Querytext: `contentclass:STS_Site AND (Title:${nameOrUrl}* OR SPSiteUrl:*${nameOrUrl}*)`,
            SelectProperties: ["Title", "SiteId", "SPSiteUrl", "WebTemplate"],
            RowLimit: 500,
            TrimDuplicates: false
        });

        return searchResults.PrimarySearchResults.map((value) => ({
            // NOTE SiteId is not in the interface => PR PnP JS
            id: value["SiteId"],
            url: value["SPSiteUrl"],
            title: value.Title
        } as ISPSite));
    }

    public async getSiteLists(webUrl: string): Promise<IList[]> {
        const serverUrl = `${document.location.protocol}//${document.location.host}`;
        const web = Web(webUrl);
        const lists = await web.lists.expand("RootFolder").select("Title", "Id", "RootFolder/ServerRelativeUrl").get();
        return lists.map(l => {
            const url = `${serverUrl}${l.RootFolder.ServerRelativeUrl}`;
            const webRelativeUrl = url.toLowerCase().replace(webUrl.toLowerCase(), "");
            return {
                title: l.Title,
                url,
                webRelativeUrl
            };
        });
    }
    
}

export const SitesServiceKey = ServiceKey.create<ISitesService>('YPCODE:SDSv2:SitesService', SitesService);
