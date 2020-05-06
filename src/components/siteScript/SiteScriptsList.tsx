import * as React from "react";
import { useState } from "react";
import {
    DocumentCard,
    DocumentCardDetails,
    DocumentCardTitle,
    DocumentCardType
} from 'office-ui-fabric-react/lib/DocumentCard';
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { Link } from "office-ui-fabric-react/lib/Link";
import { ISize } from 'office-ui-fabric-react/lib/Utilities';
import { GridLayout } from "@pnp/spfx-controls-react/lib/GridLayout";
import { ActionType, IEditSiteScriptActionArgs } from "../../app/IApplicationAction";
import styles from "./SiteScriptsList.module.scss";
import { useAppContext } from "../../app/App";
import { IApplicationState } from "../../app/ApplicationState";
import { ISiteScript } from "../../models/ISiteScript";
import { NewSiteScriptPanel } from "./NewSiteScriptPanel";

export interface ISiteScriptsListProps {
    preview?: boolean;
    addNewDisabled?: boolean;
}

const PREVIEW_ITEMS_COUNT = 3;

export const SiteScriptsList = (props: ISiteScriptsListProps) => {

    const [appContext, executeAction] = useAppContext<IApplicationState, ActionType>();

    const [isAdding, setIsAdding] = useState<boolean>(false);

    const onSiteScriptClick = (siteScript: ISiteScript) => {
        executeAction("EDIT_SITE_SCRIPT", { siteScript } as IEditSiteScriptActionArgs);
    };

    const onAddNewScript = () => {
        setIsAdding(true);
    };

    const renderSiteScriptGridItem = (siteScript: ISiteScript, finalSize: ISize, isCompact: boolean): JSX.Element => {
        if (!siteScript) {
            // If site script is not set, it is the Add new tile
            return <div
                className={styles.add}
                data-is-focusable={true}
                role="listitem"
                aria-label={"Add a new Site Script"}
            >
                <DocumentCard
                    type={isCompact ? DocumentCardType.compact : DocumentCardType.normal}
                    onClick={(ev: React.SyntheticEvent<HTMLElement>) => onAddNewScript()}>
                    <div className={styles.iconBox}>
                        <div className={styles.icon}>
                            <Icon iconName="Add" />
                        </div>
                    </div>
                </DocumentCard>
            </div>;
        }

        return <div
            data-is-focusable={true}
            role="listitem"
            aria-label={siteScript.Title}
        >
            <DocumentCard
                type={isCompact ? DocumentCardType.compact : DocumentCardType.normal}
                onClick={(ev: React.SyntheticEvent<HTMLElement>) => onSiteScriptClick(siteScript)}>
                <div className={styles.iconBox}>
                    <div className={styles.icon}>
                        <Icon iconName="Script" />
                    </div>
                </div>
                <DocumentCardDetails>
                    <DocumentCardTitle
                        title={siteScript.Title}
                        shouldTruncate={true}
                    />
                </DocumentCardDetails>
            </DocumentCard>
        </div>;
    };

    let items = [...appContext.allAvailableSiteScripts];
    if (props.preview) {
        items = items.slice(0, PREVIEW_ITEMS_COUNT);
    }
    if (!props.addNewDisabled) {
        items.push(null);
    }
    const seeMore = props.preview && appContext.allAvailableSiteScripts.length > PREVIEW_ITEMS_COUNT;
    return <div className={styles.SiteDesignsList}>
        <NewSiteScriptPanel isOpen={isAdding} onCancel={() => setIsAdding(false)} />
        <div className={styles.row}>
            <div className={styles.column}>
                <GridLayout
                    ariaLabel="List of Site Scripts."
                    items={items}
                    onRenderGridItem={renderSiteScriptGridItem}
                />
                {seeMore && <div className={styles.seeMore}>
                    {`There are more than ${PREVIEW_ITEMS_COUNT} available Site Scripts in your tenant. `}
                    <Link onClick={() => executeAction("GO_TO", { page: "SiteScriptsList" })}>See all Site Scripts</Link>
                </div>}
            </div>
        </div>
    </div>;
};