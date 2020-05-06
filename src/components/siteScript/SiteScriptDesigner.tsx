import * as React from "react";
import { useAppContext } from "../../app/App";
import { IApplicationState } from "../../app/ApplicationState";
import { SiteScriptSchemaServiceKey } from "../../services/siteScriptSchema/SiteScriptSchemaService";
import { RenderingServiceKey } from "../../services/rendering/RenderingService";
import { ActionType } from "../../app/IApplicationAction";
import styles from "./SiteScriptDesigner.module.scss";
import { Adder, IAddableItem } from "../common/Adder/Adder";
import { IconButton, Link, Label, Icon, Stack } from "office-ui-fabric-react";
import { useState, useEffect } from "react";
import { getTrimmedText } from "../../utils/textUtils";
import { useConstCallback } from "@uifabric/react-hooks";
import { ISiteScriptContentUIWrapper, ISiteScriptActionUIWrapper, SiteScriptContentUIWrapper } from "../../helpers/ScriptContentUIHelper";
import { ISiteScriptContent } from "../../models/ISiteScript";
import { usePrevious } from "../../helpers/hooks";
import { SortableElement, SortableHandle, SortableContainer } from "react-sortable-hoc";

interface ISortEndEventArgs {
    oldIndex: number;
    newIndex: number;
    collection: any[];
}

export interface ISiteScriptActionDesignerBlockProps {
    siteScriptAction: ISiteScriptActionUIWrapper;
    parentSiteScriptAction?: ISiteScriptActionUIWrapper;
    siteScriptContentUI: ISiteScriptContentUIWrapper;
    isSorting: boolean;
    index: number;
    actionsCount: number;
    onSiteScriptContentUIChanged: (siteScriptContent: ISiteScriptContentUIWrapper) => void;
}

const SEE_PROPERTIES_DEFAULT_COUNT = 2;
const SUMMARY_VALUE_MAX_LEN = 60;

export const SiteScriptActionDesignerBlock = (props: ISiteScriptActionDesignerBlockProps) => {

    const [appContext] = useAppContext<IApplicationState, ActionType>();

    // Get service references
    const siteScriptSchemaService = appContext.serviceScope.consume(SiteScriptSchemaServiceKey);
    const rendering = appContext.serviceScope.consume(RenderingServiceKey);
    const [seeMore, setSeeMore] = useState<boolean>(false);
    const [isSortingSubactions, setIsSortingSubactions] = useState<boolean>(false);


    const getActionDescription = () => siteScriptSchemaService.getActionDescription(props.siteScriptAction, props.parentSiteScriptAction);
    const getActionLabel = () => siteScriptSchemaService.getActionTitle(props.siteScriptAction, props.parentSiteScriptAction);
    const getAddableActions = () => {
        const groupLabel = `${getActionLabel()} - Subactions`;
        return {
            [groupLabel]: siteScriptSchemaService.getAvailableSubActions(props.siteScriptAction).map(a => ({
                group: groupLabel,
                iconName: "SetAction",
                key: a.verb,
                text: a.label,
                item: a
            } as IAddableItem))
        };
    };

    const toggleEdit = () => {
        props.onSiteScriptContentUIChanged(props.siteScriptContentUI.toggleEditing(props.siteScriptAction));
    };

    const onActionUpdated = (siteScriptAction: ISiteScriptActionUIWrapper) => {
        props.onSiteScriptContentUIChanged(props.siteScriptContentUI.replaceAction(siteScriptAction));
    };

    const onActionAdded = (verb: string, parentSiteScriptAction?: ISiteScriptActionUIWrapper) => {
        const newAction = parentSiteScriptAction
            ? siteScriptSchemaService.getNewSubActionFromVerb(parentSiteScriptAction.verb, verb)
            : siteScriptSchemaService.getNewActionFromVerb(verb);
        const updatedContentUI = parentSiteScriptAction
            ? props.siteScriptContentUI.addSubAction(parentSiteScriptAction, newAction)
            : props.siteScriptContentUI.addAction(newAction);
        props.onSiteScriptContentUIChanged(updatedContentUI);
    };

    const onActionRemoved = (removedAction: ISiteScriptActionUIWrapper, parentSiteScriptAction?: ISiteScriptActionUIWrapper) => {
        const updatedContentUI = parentSiteScriptAction
            ? props.siteScriptContentUI.removeSubAction(parentSiteScriptAction, removedAction)
            : props.siteScriptContentUI.removeAction(removedAction);
        props.onSiteScriptContentUIChanged(updatedContentUI);
    };

    const renderSummaryContent = (() => {
        const summaryValues = siteScriptSchemaService.getPropertiesAndValues(props.siteScriptAction, props.parentSiteScriptAction);
        if (!seeMore) {
            const previewSummary = summaryValues.slice(0, SEE_PROPERTIES_DEFAULT_COUNT);
            const displaySeeMoreLink = summaryValues.length >= SEE_PROPERTIES_DEFAULT_COUNT && !seeMore;
            return <ul>
                {previewSummary.map((pv, index) => <li key={`${props.siteScriptAction.$uiKey}_prop_${index}`}>{pv.property}: <strong title={pv.value}>{(!pv.value && pv.value !== false) ? "Not set" : getTrimmedText(pv.value, SUMMARY_VALUE_MAX_LEN)}</strong></li>)}
                {displaySeeMoreLink && <li key={`${props.siteScriptAction.$uiKey}_more_prop`}><Link onClick={() => setSeeMore(true)}>...</Link></li>}
            </ul>;
        } else {
            return <ul>
                {summaryValues.map((pv, index) => <li key={`${props.siteScriptAction.$uiKey}_prop_${index}`}>{pv.property}: <strong title={pv.value}>{!pv.value ? "Not set" : getTrimmedText(pv.value, SUMMARY_VALUE_MAX_LEN)}</strong></li>)}
            </ul>;
        }
    });

    const DragHandle = SortableHandle(() =>
        <div>
            <Icon iconName={props.index === (props.actionsCount - 1) ? "SortUp" : props.index === 0 ? "SortDown" : "Sort"} />
        </div>
    );

    const onSubActionSortChanged = (args: ISortEndEventArgs) => {
        props.onSiteScriptContentUIChanged(props.siteScriptContentUI.reorderSubActions(props.siteScriptAction.$uiKey, args.newIndex, args.oldIndex));
    };


    const toggleSortingSubactions = () => {
        if (!isSortingSubactions) {
            setTimeout(() => {
                props.onSiteScriptContentUIChanged(props.siteScriptContentUI.clearEditing([props.parentSiteScriptAction.$uiKey]));
            }, 0);
        }
        setIsSortingSubactions(!isSortingSubactions);
    };

    const renderScriptSubAction = (scriptActionUI: ISiteScriptActionUIWrapper, index: number) => {

        const scriptActionBlock = <SiteScriptActionDesignerBlock key={scriptActionUI.$uiKey}
            siteScriptAction={scriptActionUI}
            parentSiteScriptAction={props.siteScriptAction}
            siteScriptContentUI={props.siteScriptContentUI}
            onSiteScriptContentUIChanged={props.onSiteScriptContentUIChanged}
            isSorting={isSortingSubactions}
            index={index}
            actionsCount={(props.siteScriptAction
                && props.siteScriptAction.subactions
                && props.siteScriptAction.subactions.length) || 0}
        />;

        if (isSortingSubactions) {
            const SortableItem = SortableElement(() => scriptActionBlock);
            return <SortableItem key={scriptActionUI.$uiKey} index={index} />;
        } else {
            return scriptActionBlock;
        }
    };

    const renderScriptSubActionsList = () => {
        if (!props.siteScriptAction.subactions) {
            return null;
        }

        if (isSortingSubactions) {
            const SortableListContainer = SortableContainer(({ items }) => {
                return <div>{items.map(renderScriptSubAction)}</div>;
            });

            return <SortableListContainer
                items={props.siteScriptAction.subactions}
                // onSortStart={(args) => this._onSortStart(args)}
                onSortEnd={(args: any) => onSubActionSortChanged(args)}
                lockToContainerEdges={true}
                useDragHandle={false}
            />;
        } else {
            return <div>{props.siteScriptAction.subactions.map(renderScriptSubAction)}</div>;
        }
    };



    const hasSubActions = siteScriptSchemaService.hasSubActions(props.siteScriptAction);
    const isEditing = props.siteScriptContentUI.editingActionKeys.indexOf(props.siteScriptAction.$uiKey) >= 0;
    return <div className={`${styles.siteScriptAction} ${isEditing ? styles.isEditing : ""} ${props.isSorting ? styles.isSorting : ''}`}>
        <h4 title={getActionDescription()}>
            {props.isSorting && <span className={styles.sortIndex}>{props.index + 1}</span>}
            <span className={styles.actionLabelText}>{getActionLabel()}</span>
        </h4>
        <div className={styles.tools}>
            <Stack horizontal tokens={{ childrenGap: 3 }}>
                {!isEditing && props.isSorting && <DragHandle />}
                {!props.isSorting && <IconButton iconProps={{ iconName: isEditing ? "Accept" : "Edit" }} onClick={() => toggleEdit()} />}
                {!props.isSorting && !isEditing && <IconButton iconProps={{ iconName: "Delete" }} onClick={() => onActionRemoved(props.siteScriptAction, props.parentSiteScriptAction)} />}
            </Stack>
        </div>
        <div className={`${styles.summary} ${!isEditing || isSortingSubactions ? styles.isNotEditing : styles.isEditing}`}>
            {renderSummaryContent()}
        </div>
        {isEditing && <div className={`${styles.properties} ${isEditing ? styles.isEditing : ""}`}>
            {rendering.renderActionProperties(props.siteScriptAction,
                props.parentSiteScriptAction,
                (o) => onActionUpdated({ ...props.siteScriptAction, ...o } as ISiteScriptActionUIWrapper), ['verb', 'subactions', '$uiKey', '$isEditing'])}
            {hasSubActions && <div className={styles.subactions}>
                <div className={styles.row}>
                    <div className={styles.column10}>
                        <Label>Subactions</Label>
                    </div>
                    <div className={styles.column2}>
                        {props.siteScriptAction.subactions && props.siteScriptAction.subactions.length > 1 && <IconButton iconProps={{ iconName: "Sort" }}
                            // styles={{ root: { position: "absolute", top: -32, right: 9 } }}
                            checked={isSortingSubactions}
                            onClick={toggleSortingSubactions} />}
                    </div>
                </div>
                {renderScriptSubActionsList()}
                {!isSortingSubactions && <Adder items={getAddableActions()}
                    searchBoxPlaceholderText="Search a sub action..."
                    onSelectedItem={(item) => onActionAdded(item.key, props.siteScriptAction)} />}
            </div>}
        </div>}
    </div>;
};

export interface ISiteScriptDesignerProps {
    siteScriptContent: ISiteScriptContent;
    onSiteScriptContentUpdated: (updatedContent: ISiteScriptContent) => void;
}

export const SiteScriptDesigner = (props: ISiteScriptDesignerProps) => {
    const [appContext] = useAppContext<IApplicationState, ActionType>();

    // Get service references
    const siteScriptSchemaService = appContext.serviceScope.consume(SiteScriptSchemaServiceKey);
    const [isSorting, setIsSorting] = useState<boolean>(false);
    const [contentUI, setContentUI] = useState<ISiteScriptContentUIWrapper>(new SiteScriptContentUIWrapper(props.siteScriptContent));
    const previousContentUI = usePrevious<ISiteScriptContentUIWrapper>(contentUI);
    const updateUITimeoutRef = React.useRef<any>(null);
    useEffect(() => {
        const newContentUI = new SiteScriptContentUIWrapper(props.siteScriptContent);
        if (previousContentUI) {
            newContentUI.editingActionKeys = previousContentUI.editingActionKeys;
        }
        setContentUI(newContentUI);
    }, [props.siteScriptContent]);

    const onUIUpdated = (uiWrapper: ISiteScriptContentUIWrapper) => {
        setContentUI(uiWrapper);

        if (updateUITimeoutRef.current) {
            clearTimeout(updateUITimeoutRef.current);
        }
        updateUITimeoutRef.current = setTimeout(() => {
            props.onSiteScriptContentUpdated(uiWrapper.toSiteScriptContent());
            clearTimeout(updateUITimeoutRef.current);
            updateUITimeoutRef.current = null;
        }, 0);
    };

    const onActionAdded = (verb: string) => {
        const newAction = siteScriptSchemaService.getNewActionFromVerb(verb);
        const updatedContentUI = contentUI.addAction(newAction);
        setContentUI(updatedContentUI);

        if (updateUITimeoutRef.current) {
            clearTimeout(updateUITimeoutRef.current);
        }
        updateUITimeoutRef.current = setTimeout(() => {
            props.onSiteScriptContentUpdated(updatedContentUI.toSiteScriptContent());
            clearTimeout(updateUITimeoutRef.current);
            updateUITimeoutRef.current = null;
        }, 0);
    };

    const getAddableActions = useConstCallback(() => {
        return {
            "Actions": siteScriptSchemaService.getAvailableActions().map(a => ({
                group: "Actions",
                iconName: "SetAction",
                key: a.verb,
                text: a.label,
                item: a
            } as IAddableItem))
        };
    });

    const onSortChanged = (args: ISortEndEventArgs) => {
        props.onSiteScriptContentUpdated(contentUI.reorderActions(args.newIndex, args.oldIndex).toSiteScriptContent());
    };

    const toggleSorting = () => {
        if (!isSorting) {
            setTimeout(() => {
                setContentUI(contentUI.clearEditing());
            }, 0);
        }
        setIsSorting(!isSorting);
    };

    const renderScriptAction = (scriptActionUI: ISiteScriptActionUIWrapper, index: number) => {

        const scriptActionBlock = <SiteScriptActionDesignerBlock key={scriptActionUI.$uiKey}
            siteScriptAction={scriptActionUI}
            siteScriptContentUI={contentUI}
            onSiteScriptContentUIChanged={onUIUpdated}
            isSorting={isSorting}
            index={index}
            actionsCount={(contentUI.actions || 0) && contentUI.actions.length}
        />;

        if (isSorting) {
            const SortableItem = SortableElement(() => scriptActionBlock);
            return <SortableItem key={scriptActionUI.$uiKey} index={index} />;
        } else {
            return scriptActionBlock;
        }
    };

    const renderScriptActionsList = () => {
        if (!contentUI.actions) {
            return null;
        }

        if (isSorting) {
            const SortableListContainer = SortableContainer(({ items }) => {
                return <div>{items.map(renderScriptAction)}</div>;
            });

            return <SortableListContainer
                items={contentUI.actions}
                // onSortStart={(args) => this._onSortStart(args)}
                onSortEnd={(args: any) => onSortChanged(args)}
                lockToContainerEdges={true}
                useDragHandle={false}
            />;
        } else {
            return <div>{contentUI.actions.map(renderScriptAction)}</div>;
        }
    };

    return <div className={styles.SiteScriptDesigner}>
        {contentUI.actions && contentUI.actions.length > 1 && <IconButton iconProps={{ iconName: "Sort" }}
            styles={{ root: { position: "absolute", top: -32, right: 9 } }}
            checked={isSorting}
            onClick={toggleSorting} />}
        {renderScriptActionsList()}
        {!isSorting && <Adder items={getAddableActions()}
            searchBoxPlaceholderText="Search an action..."
            onSelectedItem={(item) => onActionAdded(item.key)} />}
    </div>;
};