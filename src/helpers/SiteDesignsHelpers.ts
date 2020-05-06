import { ISiteDesign } from "../models/ISiteDesign";
import { assign } from "@microsoft/sp-lodash-subset";

export const createNewSiteDesign = (siteDesignProperties?: ISiteDesign) => (assign({
    Id: null,
    Title: null,
    Description: null,
    Version: 1,
    IsDefault: false,
    PreviewImageAltText: null,
    PreviewImageUrl: null,
    SiteScriptIds: [],
    WebTemplate: ""
}, (siteDesignProperties || {})));