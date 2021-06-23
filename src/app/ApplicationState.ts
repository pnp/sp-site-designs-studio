import { ISiteDesign, ISiteDesignWithGrantedRights } from "../models/ISiteDesign";
import { IBaseAppState } from "./App";
import { ActionType } from "./IApplicationAction";
import { ISiteScript } from "../models/ISiteScript";
import { ServiceScope } from "@microsoft/sp-core-library";
import { MessageBarType } from "office-ui-fabric-react";
import { BaseComponentContext } from "@microsoft/sp-component-base";

export type Page = "Home" | "SiteDesignsList" | "SiteDesignEdition" | "SiteScriptsList" | "SiteScriptEdition";

export interface IUserMessage {
    message: string;
    messageType: MessageBarType;
}

export interface IApplicationState extends IBaseAppState<ActionType> {
    page: Page;
    currentSiteDesign: ISiteDesignWithGrantedRights;
    currentSiteScript: ISiteScript;
    allAvailableSiteDesigns: ISiteDesign[];
    allAvailableSiteScripts: ISiteScript[];
    componentContext: BaseComponentContext;
    serviceScope: ServiceScope;
    isLoading: boolean;
    userMessage: IUserMessage;
}

export const initialAppState: IApplicationState = {
    page: "Home",
    currentSiteDesign: null,
    currentSiteScript: null,
    componentContext: null,
    serviceScope: null,
    allAvailableSiteDesigns: [],
    allAvailableSiteScripts: [],
    isLoading: false,
    userMessage: null
};
