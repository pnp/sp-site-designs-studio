import { ServiceScope, ServiceKey } from "@microsoft/sp-core-library";


export type CacheRepository = "inMemory" | "session" | "persisted";

export interface ICacheItem<T> {
    /**
     * The content to cache
     */
    content: T;
    /**
     * The expiration time as ISO Date string
     * Is only applicable for "persisted" repository
     */
    expires?: string;
    /**
     * The repository to use to cache the item 
     * Available values : "inMemory"|"session"|"persisted" (default: inMemory)
     */
    repository?: CacheRepository;
}

export interface ICacheService {
    get<T>(key: string, repository?: CacheRepository): Promise<T>;
    set<T>(key: string, item: ICacheItem<T>): Promise<void>;
}

interface IInMemoryCacheRepository {
    [key: string]: any;
}

class CacheService implements ICacheService {
    private inMemoryRepository: IInMemoryCacheRepository = {};

    constructor(serviceScope: ServiceScope) {

    }

    public async get<T>(key: string, repository?: CacheRepository): Promise<T> {
        console.log("Get from cache: ", key);
        if (repository) {
            switch (repository) {
                case "inMemory":
                    return this._tryGetFromMemory<T>(key);
                case "session":
                    return this._tryGetFromSessionStorage<T>(key);
                case "persisted":
                    return this._tryGetFromLocalStorage<T>(key);
                default:
                    // If repository has invalid value, return nothing
                    return null;
            }
        }

        // If no repository is specified, lookup in each repository
        let value = this._tryGetFromMemory<T>(key);
        if (!value) {
            value = this._tryGetFromSessionStorage<T>(key);
            if (!value) {
                value = this._tryGetFromLocalStorage<T>(key);
            }
        }
        return value;
    }

    private _tryGetFromMemory<T>(key: string): T {
        return this.inMemoryRepository[key] || null;
    }

    private _tryGetFromSessionStorage<T>(key: string): T {
        if (!sessionStorage) {
            return null;
        }

        const itemStr = sessionStorage.getItem(key);
        const cacheItem = JSON.parse(itemStr) as ICacheItem<T>;
        return cacheItem ? cacheItem.content : null;
    }

    private _tryGetFromLocalStorage<T>(key: string): T {
        if (!localStorage) {
            return null;
        }

        const itemStr = localStorage.getItem(key);
        const cacheItem = JSON.parse(itemStr) as ICacheItem<T>;
        if (!cacheItem) {
            return null;
        }

        // Check for expiration
        if (cacheItem.expires) {
            const expirationDateTime = +new Date(cacheItem.expires);
            if (expirationDateTime < Date.now()) {
                return null;
            }
        }

        return cacheItem.content;
    }

    public async set<T>(itemKey: string, item: ICacheItem<T>): Promise<void> {
        console.log("Put to cache: ", item);
        if (!item || !itemKey) {
            return;
        }

        const repository = item.repository || "inMemory";
        const effectiveEntry = { ...item };
        // Remove unecessary info from cache item
        delete effectiveEntry.repository;

        switch (repository) {
            case "inMemory":
                this._setInMemory(itemKey, effectiveEntry);
                break;
            case "session":
                this._setInSessionStorage(itemKey, effectiveEntry);
                break;
            case "persisted":
                this._setInLocalStorage(itemKey, effectiveEntry);
                break;
        }
    }

    private _setInMemory<T>(itemKey: string, item: ICacheItem<T>): void {
        if (!item || !itemKey) {
            return;
        }

        if (!item.content) {
            delete this.inMemoryRepository[itemKey];
        } else {
            this.inMemoryRepository[itemKey] = item.content;
        }
    }

    private _setInSessionStorage<T>(itemKey: string, item: ICacheItem<T>): void {
        if (!sessionStorage) {
            return;
        }

        if (!item || !itemKey) {
            return;
        }

        if (!item.content) {
            sessionStorage.removeItem(itemKey);
        } else {
            sessionStorage.setItem(itemKey, JSON.stringify(item));
        }
    }

    private _setInLocalStorage<T>(itemKey: string, item: ICacheItem<T>) {
        if (!localStorage) {
            return;
        }

        if (!item || !itemKey) {
            return;
        }

        if (!item.content) {
            localStorage.removeItem(itemKey);
        } else {
            localStorage.setItem(itemKey, JSON.stringify(item));
        }
    }
}


export const CacheServiceKey = ServiceKey.create<ICacheService>('YPCODE:SDSv2:CacheService', CacheService);
