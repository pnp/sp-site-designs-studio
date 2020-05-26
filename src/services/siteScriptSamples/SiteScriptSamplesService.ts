import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import { HttpClient } from "@microsoft/sp-http";
import { ISiteScriptSamplesRepository } from "../../models/ISiteScriptSamplesRepository";
import { ISiteScriptSample } from "../../models/ISiteScriptSample";
import { find } from '@microsoft/sp-lodash-subset';
import { ICacheService, CacheServiceKey } from '../cacheService';
import * as MarkdownIt from 'markdown-it';
import * as BuiltInPnPSamples from '../../assets/pnp-samples.20200526.json';
import { ISiteScriptSchemaService, SiteScriptSchemaServiceKey } from '../siteScriptSchema/SiteScriptSchemaService';

export interface ISiteScriptSamplesService {
    getAvailableRepositories(): Promise<ISiteScriptSamplesRepository[]>;
    getSamples(repository: ISiteScriptSamplesRepository): Promise<ISiteScriptSample[]>;
    getSample(sample: ISiteScriptSample): Promise<ISiteScriptSample>;
}

interface IRepositoryFileInfo {
    name: string;
    type: string;
    path: string;
    html_url: string;
    content: string;
}

const SAMPLES_REPOSITORY_CACHE_KEY = "YPC::SDSv2_GITHUB_SAMPLES_REPO";
// Default expiration of the cache is 2 hours
const DEFAULT_CACHE_EXPIRATION_DELAY_IN_MSECS = 7200000;

const DEFAULT_RATE_LIMIT = 60;
interface IGitHubRateLimit {
    initial: number;
    remaining: number;
    retryAfter: number;
}

class SiteScriptGitHubSamplesService implements ISiteScriptSamplesService {
    private _availableRepositories: ISiteScriptSamplesRepository[] = [];
    private _cacheExpirationDelayInMsecs: number = DEFAULT_CACHE_EXPIRATION_DELAY_IN_MSECS;
    private _gitHubAPIRateLimit: IGitHubRateLimit = {
        initial: DEFAULT_RATE_LIMIT,
        remaining: DEFAULT_RATE_LIMIT,
        retryAfter: null
    };

    constructor(private serviceScope: ServiceScope) {

    }

    private _httpClient: HttpClient;
    private get httpClient(): HttpClient {
        return this._httpClient || (this._httpClient = this.serviceScope.consume(HttpClient.serviceKey));
    }

    private _cacheService: ICacheService;
    private get cache(): ICacheService {
        return this._cacheService || (this._cacheService = this.serviceScope.consume(CacheServiceKey));
    }

    private _siteScriptSchemaService: ISiteScriptSchemaService;
    private get siteScriptSchema(): ISiteScriptSchemaService {
        return this._siteScriptSchemaService || (this._siteScriptSchemaService = this.serviceScope.consume(SiteScriptSchemaServiceKey));
    }

    public async getAvailableRepositories(): Promise<ISiteScriptSamplesRepository[]> {
        return this._availableRepositories;
    }

    private async _getSampleFoldersInfo(repository: ISiteScriptSamplesRepository): Promise<IRepositoryFileInfo[]> {
        if (this._gitHubAPIRateLimit.remaining === 0) {
            throw {
                apiError: {
                    message: "The maximum number of requests to GitHub API has been reached",
                    retryAfter: new Date(this._gitHubAPIRateLimit.retryAfter * 1000)
                }
            };
        }

        return await (this.httpClient.get(`https://api.github.com/repos/${repository.owner}/${repository.repository}/contents/${repository.samplesFolderPath}`,
            HttpClient.configurations.v1).then(response => {
                this._gitHubAPIRateLimit.initial = parseInt(response.headers["X-RateLimit-Limit"]);
                this._gitHubAPIRateLimit.remaining = parseInt(response.headers["X-Ratelimit-Remaining"]);
                this._gitHubAPIRateLimit.retryAfter = parseInt(response.headers["X-RateLimit-Reset"]);
                return response.json();
            }));
    }

    public async getSamples(repository: ISiteScriptSamplesRepository): Promise<ISiteScriptSample[]> {
        try {
            const fromCache = await this.cache.get<ISiteScriptSample[]>(SAMPLES_REPOSITORY_CACHE_KEY, "persisted");
            if (fromCache) {
                return fromCache;
            }

            // Get the info of all files or folders in the root of the repository
            const responseContent = await this._getSampleFoldersInfo(repository);

            // Each sample is a directory
            const samples = responseContent.filter(s => s.type == "dir");
            const result = samples.map(s => ({
                key: s.name,
                path: s.path,
                jsonContent: null,
                readmeHtml: null,
                webSite: s.html_url,
                repository
            } as ISiteScriptSample));

            // Push the result to cache
            // Compute the expiration time from configured delay added from now
            const expires = new Date(Date.now() + this._cacheExpirationDelayInMsecs).toISOString();
            this.cache.set<ISiteScriptSample[]>(SAMPLES_REPOSITORY_CACHE_KEY, {
                content: result,
                repository: "persisted",
                expires
            });

            return result;
        } catch (error) {
            console.error("An error occured while trying to load all samples...", error);
            throw error;
        }
    }

    private async _getSampleFilesInfo(sample: ISiteScriptSample): Promise<IRepositoryFileInfo[]> {
        if (this._gitHubAPIRateLimit.remaining === 0) {
            throw {
                apiError: {
                    message: "The maximum number of requests to GitHub API has been reached",
                    retryAfter: new Date(this._gitHubAPIRateLimit.retryAfter * 1000)
                }
            };
        }

        const { repository } = sample;
        return await (this.httpClient.get(`https://api.github.com/repos/${repository.owner}/${repository.repository}/contents/${sample.path}`,
            HttpClient.configurations.v1).then(response => {
                this._gitHubAPIRateLimit.initial = parseInt(response.headers["X-RateLimit-Limit"]);
                this._gitHubAPIRateLimit.remaining = parseInt(response.headers["X-Ratelimit-Remaining"]);
                this._gitHubAPIRateLimit.retryAfter = parseInt(response.headers["X-RateLimit-Reset"]);
                return response.json();
            }));
    }

    private async _getSampleJsonContent(sample: ISiteScriptSample, jsonFilePath: string): Promise<string> {
        const { repository } = sample;
        return await this.httpClient.get(`https://raw.githubusercontent.com/${repository.owner}/${repository.repository}/${repository.branch}/${jsonFilePath}`, HttpClient.configurations.v1).then(r => r.text());
    }

    private async _getSampleReadmeContent(sample: ISiteScriptSample, readmeFilePath: string): Promise<string> {
        const { repository } = sample;
        return await this.httpClient.get(`https://raw.githubusercontent.com/${repository.owner}/${repository.repository}/${repository.branch}/${readmeFilePath}`, HttpClient.configurations.v1).then(r => r.text());
    }

    private async _renderMarkdown(markdown: string): Promise<string> {

        //#region  Initial draft using GitHub API
        // not used to spare unauthenticated GitHub API allowed requests
        // const renderMarkdownRequestBody = {
        //     "text": markdown,
        // };
        // return await this.httpClient.post(`https://api.github.com/markdown`, HttpClient.configurations.v1, {
        //     body: JSON.stringify(renderMarkdownRequestBody)
        // })
        //     .then(r => r.text());
        ////#endregion

        const md = new MarkdownIt({ html: false });
        const html = md.render(markdown);
        return html;
    }

    private _preprocessRelativeUrlInHtml(html: string, sample: ISiteScriptSample): string {
        const { repository } = sample;
        // Preprocess relative path in href="" and in src=""
        const regexRelativeUrl = /(?<attr>href|src)=\"(?!http|https)(?<relativePath>[^\"]*)\"/g;
        const absoluteUrl = `https://raw.githubusercontent.com/${repository.owner}/${repository.repository}/${repository.branch}/${sample.path}`;
        return html.replace(regexRelativeUrl, `$<attr>="${absoluteUrl}/$<relativePath>"`);
    }

    private _tryGetSampleFromBuiltinSamples(key: string): ISiteScriptSample {
        const builtInSamples = BuiltInPnPSamples.content as ISiteScriptSample[];
        const foundSample = find(builtInSamples, s => s.key == key);
        return foundSample;
    }

    public async getSample(sample: ISiteScriptSample): Promise<ISiteScriptSample> {
        let allSamplesFromCache = await this.cache.get<ISiteScriptSample[]>(SAMPLES_REPOSITORY_CACHE_KEY, "persisted");
        if (allSamplesFromCache) {
            // Try to find the specific sample from cache with set jsonContent
            const sampleFromCache = find(allSamplesFromCache, s => s.key == sample.key && !!s.jsonContent);
            if (sampleFromCache) {
                return sampleFromCache;
            }
        }

        try {
            const repository = sample.repository;
            if (!repository) {
                throw new Error("The repository is not set for the specified sample");
            }

            const sampleFilesInfo = await this._getSampleFilesInfo(sample);

            // Try get JSON file
            const jsonFile = find(sampleFilesInfo, f => f.name.endsWith('.json'));

            // Get the content of the JSON file
            let jsonContent = jsonFile ? await this._getSampleJsonContent(sample, jsonFile.path) : '';
            let hasPreprocessedJsonContent = false;
            if (jsonContent) {
                const jsonContentWithoutComment = jsonContent.replace(/\/\*(.*)\*\//g,'');
                if (!this.siteScriptSchema.validateSiteScriptJson(jsonContentWithoutComment)) {
                    jsonContent = null;
                } else {
                    hasPreprocessedJsonContent = jsonContentWithoutComment != jsonContent;
                    jsonContent = jsonContentWithoutComment;
                }
            }

            let readmeHtml: string = null;
            try {
                // Try get README file
                const readmeFile = find(sampleFilesInfo, f => f.name.toLowerCase() == 'readme.md');
                if (!readmeFile) {
                    throw new Error("README file cannot be found");
                }
                const readmeFileMarkdownContent = await this._getSampleReadmeContent(sample, readmeFile.path);

                // Render the README file markdown
                readmeHtml = await this._renderMarkdown(readmeFileMarkdownContent);
                // Replace relative URLs by absolute URLs to GitHub in rendered HTML
                readmeHtml = this._preprocessRelativeUrlInHtml(readmeHtml, sample);
            } catch (error) {
                console.debug("No README file found for this sample or it cannot be processed...");
            }

            // Set the loaded information to original sample
            const loadedSample = {
                ...sample,
                jsonContent,
                hasPreprocessedJsonContent,
                hasUsableJsonSample: !!jsonContent,
                readmeHtml
            } as ISiteScriptSample;

            // Push the loaded sample to cache
            allSamplesFromCache = allSamplesFromCache || [];
            if (allSamplesFromCache.length > 0) {
                // Try to find the sample to cache the extra data for
                const foundSampleFromCache = find(allSamplesFromCache, s => s.key == sample.key);
                if (foundSampleFromCache) {
                    foundSampleFromCache.jsonContent = jsonContent;
                    foundSampleFromCache.readmeHtml = readmeHtml;
                }
            } else {
                allSamplesFromCache.push(loadedSample);
            }
            // Ensure the new loaded sample is stored in cached with all its loaded data
            // Compute the expiration time from configured delay added from now
            const expires = new Date(Date.now() + this._cacheExpirationDelayInMsecs).toISOString();
            this.cache.set<ISiteScriptSample[]>(SAMPLES_REPOSITORY_CACHE_KEY, {
                content: allSamplesFromCache,
                repository: "persisted",
                expires
            });

            return loadedSample;
        } catch (error) {
            // If any error occurs while trying to fetch the sample from GitHub, fallback to built-in samples
            const builtInSample = this._tryGetSampleFromBuiltinSamples(sample.key);
            if (builtInSample) {
                return builtInSample;
            }

            // If not found from fallback built-in repository, raise an error
            console.error(`An error occured while trying to load sample ${sample.key}`, error);
            throw error;
        }
    }
}

export const SiteScriptSamplesServiceKey = ServiceKey.create<ISiteScriptSamplesService>('YPCODE:SDSv2:SiteScriptSamplesService', SiteScriptGitHubSamplesService);
