import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { sp } from "@pnp/sp/presets/all";

const PREVIEW_IMAGES_LIBRARY_WEBRELATIVE_URL = "SiteDesignPreviewImages";

export interface ISiteDesignPreviewImageService {
    /**
     * Upload preview image to current site
     * @param file The file to upload
     * @returns Promise<string> The absolute URL of the uploaded file
     */
    uploadPreviewImageToCurrentSite(file: File): Promise<string>;
}

class SiteDesignPreviewImageService implements ISiteDesignPreviewImageService {

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


    public async uploadPreviewImageToCurrentSite(file: File): Promise<string> {
        const serverUrl = `${document.location.protocol}//${document.location.host}`;
        const webInfo = await sp.web.select("Url")();
        let libServerRelativeUrl = `${webInfo.Url.replace(serverUrl, "")}/${PREVIEW_IMAGES_LIBRARY_WEBRELATIVE_URL}`;
        try {
            await sp.web.getList(libServerRelativeUrl).select("Id")();
            console.debug("The library does exist");
        } catch (er) {
            console.debug("The library does not exist, will try to create it");
            const createdList = await sp.web.lists.add(PREVIEW_IMAGES_LIBRARY_WEBRELATIVE_URL, "A library to store the Site Designs preview images", 101);
            const { ServerRelativeUrl } = await createdList.list.rootFolder.select("ServerRelativeUrl")();
            libServerRelativeUrl = ServerRelativeUrl;
        }

        const createdFile = await sp.web.getFolderByServerRelativeUrl(libServerRelativeUrl).files.add(file.name, file, true);
        return createdFile.data.ServerRelativeUrl;
    }
}

export const SiteDesignPreviewImageServiceKey = ServiceKey.create<ISiteDesignPreviewImageService>(
    'YPCODE:SDSv2:SiteDesignPreviewImageService',
    SiteDesignPreviewImageService
);
