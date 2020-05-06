import { IApplicationState } from "./ApplicationState";
import { IEditSiteDesignActionArgs, ActionType, ISetAllAvailableSiteDesigns, ISetAllAvailableSiteScripts, IEditSiteScriptActionArgs, ISetLoadingArgs, ISetUserMessageArgs } from "./IApplicationAction";
import { IAction } from "./App";

export const Reducers: (applicationState: IApplicationState, action: IAction<ActionType>) => IApplicationState =
    (applicationState: IApplicationState, action: IAction<ActionType>) => {
        if (!action) {
            return applicationState;
        }

        const actionArgs = action as any;


        switch (action.type) {
            case "GO_TO":
                return {
                    ...applicationState,
                    page: actionArgs.page
                };
            case "EDIT_SITE_DESIGN":
                const editSiteDesignAction = (actionArgs as IEditSiteDesignActionArgs);
                const currentSiteDesign = editSiteDesignAction.siteDesign;
                if (editSiteDesignAction.additionalSiteScriptIds) {
                    currentSiteDesign.SiteScriptIds = [...(currentSiteDesign.SiteScriptIds || []), ...editSiteDesignAction.additionalSiteScriptIds];
                }
                return {
                    ...applicationState,
                    page: "SiteDesignEdition",
                    currentSiteDesign: currentSiteDesign
                };
            case "EDIT_SITE_SCRIPT":
                return {
                    ...applicationState,
                    page: "SiteScriptEdition",
                    currentSiteScript: (actionArgs as IEditSiteScriptActionArgs).siteScript
                };
            case "SET_ALL_AVAILABLE_SITE_DESIGNS":
                return {
                    ...applicationState,
                    allAvailableSiteDesigns: (actionArgs as ISetAllAvailableSiteDesigns).siteDesigns
                };
            case "SET_ALL_AVAILABLE_SITE_SCRIPTS":
                return {
                    ...applicationState,
                    allAvailableSiteScripts: (actionArgs as ISetAllAvailableSiteScripts).siteScripts
                };
            case "SET_LOADING":
                return {
                    ...applicationState,
                    isLoading: (actionArgs as ISetLoadingArgs).loading
                };
            case "SET_USER_MESSAGE":
                return {
                    ...applicationState,
                    userMessage: (actionArgs as ISetUserMessageArgs).userMessage
                };
            default:
                return applicationState;
        }
    };