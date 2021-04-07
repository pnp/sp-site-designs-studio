import * as React from "react";
import { useState, useEffect } from "react";
import { find } from "@microsoft/sp-lodash-subset";
import { Dropdown } from "office-ui-fabric-react/lib/Dropdown";
import { Label } from "office-ui-fabric-react/lib/Label";
import { Stack } from "office-ui-fabric-react/lib/Stack";
import { MessageBarType } from "office-ui-fabric-react/lib/MessageBar";
import { Toggle } from "office-ui-fabric-react/lib/Toggle";
import { ProgressIndicator } from "office-ui-fabric-react/lib/ProgressIndicator";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { ImageFit } from "office-ui-fabric-react/lib/Image";
import { DocumentCardPreview, IDocumentCardPreviewProps } from "office-ui-fabric-react/lib/DocumentCard";
import { ActionButton, PrimaryButton, DefaultButton } from "office-ui-fabric-react/lib/Button";
import { SortableContainer, SortableHandle, SortableElement } from 'react-sortable-hoc';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import styles from "./SiteDesignEditor.module.scss";
import { WebTemplate, ISiteDesignWithGrantedRights, ISiteDesignGrantedPrincipal } from "../../models/ISiteDesign";
import { useAppContext } from "../../app/App";
import { IApplicationState } from "../../app/ApplicationState";
import { ActionType, ISetAllAvailableSiteDesigns, IGoToActionArgs, ISetUserMessageArgs } from "../../app/IApplicationAction";
import { Adder, IAddableItem } from "../common/Adder/Adder";
import { SiteDesignsServiceKey } from "../../services/siteDesigns/SiteDesignsService";
import { ISiteScript } from "../../models/ISiteScript";
import { Confirm } from "../common/confirm/Confirm";
import { PeoplePicker, PrincipalType } from "../common/peoplePicker/PeoplePicker";
import { SiteDesignPreviewImageServiceKey } from "../../services/siteDesignPreviewImage/SiteDesignPreviewImageService";
import { getPrincipalTypeFromName, getPrincipalAlias } from "../../utils/spUtils";
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { IPrincipal } from "../../models/IPrincipal";

export interface ISiteDesignEditorProps {
    siteDesign: ISiteDesignWithGrantedRights;
}

export interface ISiteDesignAssociatedSiteScriptsProps extends ISiteDesignEditorProps {
    onAssociatedSiteScriptAdded: (siteScriptId: string) => void;
    onAssociatedSiteScriptRemoved: (siteScriptId: string) => void;
    onAssociatedSiteScriptsReordered: (reordered: string[]) => void;
}

interface ISortEndEventArgs {
    oldIndex: number;
    newIndex: number;
    collection: any[];
}

export const SiteDesignAssociatedScripts = (props: ISiteDesignAssociatedSiteScriptsProps) => {
    const [appContext, execute] = useAppContext<IApplicationState, ActionType>();

    if (!appContext.allAvailableSiteScripts) {
        return <div>Loading site scripts...</div>;
    }

    const selectedSiteScripts = props.siteDesign.SiteScriptIds
        .map(id => find(appContext.allAvailableSiteScripts, ss => ss.Id == id))
        // Ensure the associated scripts are still available
        .filter(i => !!i);
    const adderItems: IAddableItem[] = appContext.allAvailableSiteScripts.filter(ss => props.siteDesign.SiteScriptIds.indexOf(ss.Id) < 0).map(item => ({
        iconName: "Script",
        group: null,
        key: item.Id,
        text: item.Title,
        item
    }));

    const renderSiteScriptItem = (siteScript: ISiteScript, index: number) => {
        const DragHandle = SortableHandle(() => (
            <div className={styles.column10}>
                <h4>{siteScript.Title}</h4>
                <div>{siteScript.Description}</div>
            </div>
        ));

        const SortableItem = SortableElement(({ value }) => <div key={`selectedSiteScript_${index}`} className={styles.selectedSiteScript}>
            <div className={styles.row}>
                <DragHandle />
                <div className={`${styles.column2} ${styles.righted}`}>
                    <ActionButton iconProps={{ iconName: "Delete" }} onClick={() => props.onAssociatedSiteScriptRemoved(value.Id)} />
                </div>
            </div>
        </div>);

        return <SortableItem key={`SiteScript_${siteScript.Id}`} value={siteScript} index={index} />;
    };

    const SortableListContainer = SortableContainer(({ items }) => {
        return <div>{items.map(renderSiteScriptItem)}</div>;
    });

    const onSortChanged = (args: ISortEndEventArgs) => {
        const toSortSiteScriptIds = [...props.siteDesign.SiteScriptIds];
        toSortSiteScriptIds.splice(args.oldIndex, 1);
        toSortSiteScriptIds.splice(args.newIndex, 0, props.siteDesign.SiteScriptIds[args.oldIndex]);
        props.onAssociatedSiteScriptsReordered(toSortSiteScriptIds);
    };

    return <>
        <Label>Associated Site Scripts: </Label>
        <SortableListContainer
            items={selectedSiteScripts}
            // onSortStart={(args) => this._onSortStart(args)}
            onSortEnd={(args: any) => onSortChanged(args)}
            lockToContainerEdges={true}
            useDragHandle={true}
        />
        <Adder items={{ "Available Site Scripts": adderItems }}
            onSelectedItem={(item) => props.onAssociatedSiteScriptAdded(item.item.Id)} />
    </>;
};

export const SiteDesignEditor = (props: ISiteDesignEditorProps) => {

    const [appContext, execute] = useAppContext<IApplicationState, ActionType>();
    const siteDesignsService = appContext.serviceScope.consume(SiteDesignsServiceKey);
    const siteDesignPreviewImageService = appContext.serviceScope.consume(SiteDesignPreviewImageServiceKey);

    const [editingSiteDesign, setEditingSiteDesign] = useState<ISiteDesignWithGrantedRights>(props.siteDesign);
    const [isSaving, setIsSaving] = useState<boolean>(false);

    const setLoading = (loading: boolean) => {
        execute("SET_LOADING", { loading });
    };

    useEffect(() => {
        if (!props.siteDesign.Id) {
            setEditingSiteDesign(props.siteDesign);
            return;
        }

        setLoading(true);
        siteDesignsService.getSiteDesign(props.siteDesign.Id).then(siteDesign => {
            // Ensure associated site scripts from props are added to edited site design
            const presetAssociatedSiteScripts = props.siteDesign.SiteScriptIds
                ? props.siteDesign.SiteScriptIds.filter(ssId => siteDesign.SiteScriptIds.indexOf(ssId) < 0)
                : [];
            siteDesign.SiteScriptIds.push(...presetAssociatedSiteScripts);
            setEditingSiteDesign(siteDesign);
        }).catch(error => {
            console.error(`The Site Design ${props.siteDesign.Id} could not be loaded`, error);
        }).then(() => {
            setLoading(false);
        });
    }, [props.siteDesign]);

    const onTitleChanged = (ev: any, title: string) => {
        setEditingSiteDesign({ ...editingSiteDesign, Title: title });
    };

    const onDescriptionChanged = (ev: any, description: string) => {
        setEditingSiteDesign({ ...editingSiteDesign, Description: description });
    };

    const onIsDefaultChanged = (ev: any, isDefault: boolean) => {
        setEditingSiteDesign({ ...editingSiteDesign, IsDefault: isDefault });
    };

    const onPreviewImageAltTextChanged = (ev: any, previewImageAltText: string) => {
        setEditingSiteDesign({ ...editingSiteDesign, PreviewImageAltText: previewImageAltText });
    };

    const onPreviewImageRemoved = () => {
        setEditingSiteDesign({ ...editingSiteDesign, PreviewImageUrl: "", PreviewImageAltText: "" });
    };

    const onVersionChanged = (ev: any, version: string) => {
        const versionInt = parseInt(version);
        if (!isNaN(versionInt)) {
            setEditingSiteDesign({ ...editingSiteDesign, Version: versionInt });
        }
    };

    const onWebTemplateChanged = (webTemplate: string) => {
        setEditingSiteDesign({ ...editingSiteDesign, WebTemplate: webTemplate });
    };

    const onAssociatedSiteScriptsAdded = (siteScriptId: string) => {
        if (!editingSiteDesign.SiteScriptIds) {
            return;
        }

        const newSiteScriptIds = editingSiteDesign.SiteScriptIds.filter(sid => sid != siteScriptId).concat(siteScriptId);
        setEditingSiteDesign({ ...editingSiteDesign, SiteScriptIds: newSiteScriptIds });
    };


    const onAssociatedSiteScriptsRemoved = (siteScriptId: string) => {
        if (!editingSiteDesign.SiteScriptIds) {
            return;
        }

        const newSiteScriptIds = editingSiteDesign.SiteScriptIds.filter(sid => sid != siteScriptId);
        setEditingSiteDesign({ ...editingSiteDesign, SiteScriptIds: newSiteScriptIds });
    };

    const onAssociatedSiteScriptsReordered = (reorderedSiteScriptIds: string[]) => {
        if (!editingSiteDesign.SiteScriptIds) {
            return;
        }

        setEditingSiteDesign({ ...editingSiteDesign, SiteScriptIds: reorderedSiteScriptIds });
    };

    const onPreviewImageChanged = async (previewImageFile: IFilePickerResult) => {
        if (previewImageFile.fileAbsoluteUrl) {
            setEditingSiteDesign({ ...editingSiteDesign, PreviewImageUrl: previewImageFile.fileAbsoluteUrl });
        } else {
            const file: File = await previewImageFile.downloadFileContent();
            // Will add the file in a "SiteDesignsPreviewImages" library in the current site
            const serverUrl = `${document.location.protocol}//${document.location.host}`;
            const uploadedFileUrl = await siteDesignPreviewImageService.uploadPreviewImageToCurrentSite(file);
            const previewImageUrl = `${serverUrl}/${uploadedFileUrl}`;
            setEditingSiteDesign({ ...editingSiteDesign, PreviewImageUrl: previewImageUrl });
        }
    };

    const isValidForSave: () => [boolean, string?] = () => {
        if (!editingSiteDesign) {
            return [false, "Current Site Design not defined"];
        }

        if (!editingSiteDesign.Title) {
            return [false, "Please set the title of the Site Design..."];
        }

        if (!editingSiteDesign.WebTemplate) {
            return [false, "Please set the web template of the Site Design..."];
        }

        return [true];
    };

    const onGrantedChange = (items: IPrincipal[]) => {
        const grantedRightsPrincipals: ISiteDesignGrantedPrincipal[] = items.map(i => (
            {
                id: null,
                displayName: i.displayName,
                principalName: i.principalName,
                type: i.type || getPrincipalTypeFromName(i.principalName),
                alias: getPrincipalAlias(i)
            }));
        setEditingSiteDesign({ ...editingSiteDesign, grantedRightsPrincipals });
    };

    const onSave = async () => {
        setIsSaving(true);
        try {
            await siteDesignsService.saveSiteDesign(editingSiteDesign);
            const refreshedSiteDesigns = await siteDesignsService.getSiteDesigns();
            execute("SET_USER_MESSAGE", {
                userMessage: {
                    message: `${editingSiteDesign.Title} has been successfully saved.`,
                    messageType: MessageBarType.success
                }
            } as ISetUserMessageArgs);
            execute("SET_ALL_AVAILABLE_SITE_DESIGNS", { siteDesigns: refreshedSiteDesigns } as ISetAllAvailableSiteDesigns);
            // If it is a brand new design, force redirect to the script list
            if (!editingSiteDesign.Id) {
                execute("GO_TO", { page: "SiteDesignsList" } as IGoToActionArgs);
            }
        } catch (error) {
            execute("SET_USER_MESSAGE", {
                userMessage: {
                    message: `${editingSiteDesign.Title} could not be saved. Please make sure you have SharePoint administrator privileges...`,
                    messageType: MessageBarType.error
                }
            } as ISetUserMessageArgs);
            console.error(error);
        }
        setIsSaving(false);
    };

    const onDelete = async () => {

        if (!await Confirm.show({
            title: `Delete Site Design`,
            message: `Are you sure you want to delete ${editingSiteDesign.Title || "this Site Design"} ?`
        })) {
            return;
        }

        setIsSaving(true);
        try {
            await siteDesignsService.deleteSiteDesign(editingSiteDesign);
            const refreshedSiteDesigns = await siteDesignsService.getSiteDesigns();
            execute("SET_USER_MESSAGE", {
                userMessage: {
                    message: `${editingSiteDesign.Title} has been successfully deleted.`,
                    messageType: MessageBarType.success
                }
            } as ISetUserMessageArgs);
            execute("SET_ALL_AVAILABLE_SITE_DESIGNS", { siteDesigns: refreshedSiteDesigns } as ISetAllAvailableSiteDesigns);
            execute("GO_TO", { page: "SiteDesignsList" } as IGoToActionArgs);
        } catch (error) {
            execute("SET_USER_MESSAGE", {
                userMessage: {
                    message: `${editingSiteDesign.Title} could not be deleted. Please make sure you have SharePoint administrator privileges...`,
                    messageType: MessageBarType.error
                }
            } as ISetUserMessageArgs);
            console.error(error);
        }

        setIsSaving(false);
    };

    const previewProps: IDocumentCardPreviewProps = {
        previewImages: [
            {
                previewImageSrc: editingSiteDesign.PreviewImageUrl,
                imageFit: ImageFit.centerContain,
                height: 300
            }
        ]
    };

    const isLoading = appContext.isLoading;
    const [isValid, validationMessage] = isValidForSave();
    return <div>
        <div className={styles.row}>
            <div className={styles.columnLayout}>
                <div className={styles.row}>
                    <div className={styles.column11}>
                        <TextField
                            styles={{
                                field: {
                                    fontSize: "32px",
                                    lineHeight: "45px",
                                    height: "45px"
                                },
                                root: {
                                    height: "60px",
                                    marginTop: "5px",
                                    marginBottom: "5px"
                                }
                            }}
                            placeholder="Enter the name of the Site Design..."
                            borderless
                            readOnly={isLoading}
                            value={editingSiteDesign.Title}
                            onChange={onTitleChanged} />
                        {isLoading && <ProgressIndicator />}
                    </div>
                    {!isLoading && <div className={`${styles.column1} ${styles.righted}`}>
                        <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 15 }}>
                            {editingSiteDesign.Id && <DefaultButton disabled={isSaving} text="Delete" iconProps={{ iconName: "Delete" }} onClick={() => onDelete()} />}
                            <PrimaryButton disabled={isSaving || !isValid} title={!isValid && validationMessage} text="Save" iconProps={{ iconName: "Save" }} onClick={() => onSave()} />
                        </Stack>
                    </div>}
                </div>
                {!isLoading && <div className={styles.row}>
                    <div className={styles.half}>
                        {editingSiteDesign.Id && <div className={styles.row}>
                            <div className={styles.column6}>
                                <TextField
                                    label="Id"
                                    readOnly
                                    value={editingSiteDesign.Id} />
                            </div>
                            <div className={styles.column4}>
                                <Dropdown
                                    label="Site Template"
                                    options={[
                                        { key: WebTemplate.TeamSite.toString(), text: 'Team Site' },
                                        { key: WebTemplate.CommunicationSite.toString(), text: 'Communication Site' }
                                    ]}
                                    selectedKey={editingSiteDesign.WebTemplate}
                                    onChange={(_, v) => onWebTemplateChanged(v.key as string)}
                                />
                            </div>
                            <div className={styles.column2}>
                                <TextField
                                    label="Version"
                                    value={editingSiteDesign.Version && editingSiteDesign.Version.toString()}
                                    onChange={onVersionChanged} />
                            </div>
                        </div>}
                        {!editingSiteDesign.Id && <div className={styles.row}>
                            <div className={styles.column}>
                                <Dropdown
                                    label="Site Template"
                                    options={[
                                        { key: WebTemplate.TeamSite.toString(), text: 'Team Site' },
                                        { key: WebTemplate.CommunicationSite.toString(), text: 'Communication Site' }
                                    ]}
                                    selectedKey={editingSiteDesign.WebTemplate}
                                    onChange={(_, v) => onWebTemplateChanged(v.key as string)}
                                />
                            </div>
                        </div>}
                        <div className={styles.row}>
                            <div className={styles.column}>
                                <Toggle
                                    label="Is Default ?"
                                    checked={editingSiteDesign.IsDefault}
                                    onChange={onIsDefaultChanged}
                                />
                            </div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column}>
                                <TextField
                                    label="Description"
                                    value={editingSiteDesign.Description}
                                    multiline={true}
                                    borderless
                                    rows={5}
                                    onChange={onDescriptionChanged}
                                />
                            </div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column}>
                                <PeoplePicker
                                    serviceScope={appContext.serviceScope}
                                    onChange={onGrantedChange}
                                    selectedItems={editingSiteDesign.grantedRightsPrincipals || []}
                                    principalType={PrincipalType.All}
                                    label="Granted to"
                                    typePicker="normal"
                                />
                                {(!editingSiteDesign.grantedRightsPrincipals || editingSiteDesign.grantedRightsPrincipals.length == 0)
                                    && <Label styles={{ root: { fontSize: 12 } }}>
                                        <Icon iconName="InfoSolid" styles={{ root: { marginRight: 10 } }} />
                                        If nobody is explicity granted, this Site Design will be available for everyone
                                        </Label>}
                            </div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column}>
                                <SiteDesignAssociatedScripts siteDesign={editingSiteDesign}
                                    onAssociatedSiteScriptAdded={onAssociatedSiteScriptsAdded}
                                    onAssociatedSiteScriptRemoved={onAssociatedSiteScriptsRemoved}
                                    onAssociatedSiteScriptsReordered={onAssociatedSiteScriptsReordered} />
                            </div>
                        </div>
                    </div>
                    <div className={styles.siteDesignImage}>
                        <div className={styles.righted}>
                            <Stack horizontal horizontalAlign="end">
                                {editingSiteDesign.PreviewImageUrl && <ActionButton text="Remove preview image" iconProps={{ iconName: "Delete" }} onClick={onPreviewImageRemoved} />}
                                <FilePicker
                                    accepts={[".gif", ".jpg", ".jpeg", ".bmp", ".png"]}
                                    buttonIcon="FileImage"
                                    onSave={onPreviewImageChanged}
                                    buttonLabel={editingSiteDesign.PreviewImageUrl ? "Modify preview image" : "Add a preview image"}
                                    context={appContext.componentContext}
                                />
                            </Stack>
                        </div>
                        {!editingSiteDesign.PreviewImageUrl && <div><Placeholder iconName='FileImage'
                            iconText='No preview image...'
                            description='There is no defined preview image for this Site Design...' />
                        </div>}
                        {editingSiteDesign.PreviewImageUrl && <div className={styles.imgPlaceholder}>
                            <DocumentCardPreview {...previewProps} />
                        </div>}
                        <div>
                            {editingSiteDesign.PreviewImageUrl && <TextField
                                value={editingSiteDesign.PreviewImageAltText}
                                borderless
                                styles={{
                                    field: {
                                        fontSize: "16px",
                                        lineHeight: "30px",
                                        height: "30px",
                                        textAlign: "center"
                                    },
                                    root: {
                                        height: "30px",
                                        width: "80%",
                                        margin: "auto",
                                        marginTop: "5px"
                                    }
                                }}
                                placeholder="Enter the alternative text for preview image..."
                                onChange={onPreviewImageAltTextChanged}
                            />}
                        </div>
                    </div>
                </div>}
            </div>
        </div>
    </div>;
};
