import { IHubSite } from '../../models/IHubSite';
import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';

export interface IHubSitesService {
    baseUrl: string;
    getHubSites(): Promise<IHubSite[]>;
}

class HubSitesService implements IHubSitesService {
    private spHttpClient: SPHttpClient;
    private cachedResults: IHubSite[];

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            this.spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
        });
    }

    public baseUrl: string = '/';

    private _getEffectiveUrl(relativeUrl: string): string {
        return (this.baseUrl + '//' + relativeUrl).replace(/\/{2,}/, '/');
    }

    private _restRequest(url: string): Promise<any> {
        const restUrl = this._getEffectiveUrl(url);
        const options: ISPHttpClientOptions = {
            headers: {
                'Content-Type': 'application/json;charset=utf-8',
                ACCEPT: 'application/json; odata.metadata=minimal',
                'ODATA-VERSION': '4.0'
            }
        };
        return this.spHttpClient.get(restUrl, SPHttpClient.configurations.v1, options).then((response) => {
            if (response.status == 204) {
                return {};
            } else {
                return response.json();
            }
        });
    }

    public async getHubSites(): Promise<IHubSite[]> {
        if (this.cachedResults && this.cachedResults.length > 0) {
            return Promise.resolve(this.cachedResults);
        } else {
            return this._restRequest('/_api/hubSites').then((resp) => {
                this.cachedResults = resp.value.map((hubSite) => ({
                    title: hubSite.Title,
                    id: hubSite.ID
                }));
                return this.cachedResults;
            });
        }
        // return Promise.resolve([
        //     {
        //         id: "1",
        //         title: "Hub site 1"
        //     },
        //     {
        //         id: "2",
        //         title: "HubSite 2"
        //     }
        // ]);
    }
}

export const HubSitesServiceKey = ServiceKey.create<IHubSitesService>('YPCODE:SDSv2:HubSitesService', HubSitesService);
