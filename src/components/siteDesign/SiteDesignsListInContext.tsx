import * as React from "react";
import { useAppContext } from "../../app/App";
import { IApplicationState } from "../../app/ApplicationState";
import { ActionType, IEditSiteDesignActionArgs } from "../../app/IApplicationAction";
import { ISiteDesignsListAllOptionalProps, SiteDesignsList } from "./SiteDesignsList";
import { ISiteDesign } from "../../models/ISiteDesign";
import { createNewSiteDesign } from "../../helpers/SiteDesignsHelpers";


/**
 * This component users the global app context to pass all the site designs to the actual List component
 * @param props 
 */
export const SiteDesignsListInContext = (props: ISiteDesignsListAllOptionalProps) => {
    const [appContext, executeAction] = useAppContext<IApplicationState, ActionType>();

    const onSiteDesignClick = (siteDesign: ISiteDesign) => {
        executeAction("EDIT_SITE_DESIGN", { siteDesign } as IEditSiteDesignActionArgs);
    };

    const onNewSiteDesignAdded = () => {
        const siteDesign: ISiteDesign = createNewSiteDesign();
        executeAction("EDIT_SITE_DESIGN", { siteDesign } as IEditSiteDesignActionArgs);
    };

    return <SiteDesignsList siteDesigns={appContext.allAvailableSiteDesigns}
        onSiteDesignClicked={onSiteDesignClick}
        onSeeMore={() => executeAction("GO_TO", { page: "SiteDesignsList" })}
        onAdd={onNewSiteDesignAdded}  {...props} />;
};