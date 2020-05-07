import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import { HttpClient } from "@microsoft/sp-http";
import { ISiteScriptSamplesRepository } from "../../models/ISiteScriptSamplesRepository";
import { ISiteScriptSample } from "../../models/ISiteScriptSample";

export interface ISiteScriptSamplesService {
    getAvailableRepositories(): Promise<ISiteScriptSamplesRepository[]>;
    getSamples(repository: ISiteScriptSamplesRepository): Promise<ISiteScriptSample[]>;
    getSample(repository: ISiteScriptSamplesRepository, key: string): Promise<ISiteScriptSample>;
}

class SiteScriptSamplesService implements ISiteScriptSamplesService {

    constructor(private serviceScope: ServiceScope, private availableRepositories?: ISiteScriptSamplesRepository[]) {

    }

    private _httpClient: HttpClient;
    private get httpClient(): HttpClient {
        return this._httpClient || (this._httpClient = this.serviceScope.consume(HttpClient.serviceKey));
    }

    public async getAvailableRepositories(): Promise<ISiteScriptSamplesRepository[]> {
        return this.availableRepositories;
    }

    public async getSamples(repository: ISiteScriptSamplesRepository): Promise<ISiteScriptSample[]> {
        const responseContent: {name:string; type:string;}[] = await (this.httpClient.get(`https://api.github.com/repo/${repository.owner}/${repository.repository}/contents/${repository.samplesFolderPath}`, HttpClient.configurations.v1).then(response => response.json()));
        const samples = responseContent.filter(s => s.type == "dir");
        return samples.map(s => ({key: s.name, contentJson: null, readmeHtml: null}));
    }
    getSample(repository: ISiteScriptSamplesRepository, key: string): Promise<ISiteScriptSample> {
        throw new Error("Method not implemented.");
    }

}

export const SiteScriptSamplesServiceKey = ServiceKey.create<ISiteScriptSamplesService>('YPCODE:SDSv2:SiteScriptSamplesService', SiteScriptSamplesService);
