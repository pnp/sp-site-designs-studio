import { ISiteScriptSamplesRepository } from "./ISiteScriptSamplesRepository";

export interface ISiteScriptSample {
    key: string;
    path: string;
    readmeHtml: string;
    jsonContent: string;
    hasUsableJsonSample?: boolean;
    hasPreprocessedJsonContent?: boolean;
    webSite: string;
    repository: ISiteScriptSamplesRepository;
}