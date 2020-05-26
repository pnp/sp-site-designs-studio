import * as React from "react";
import { useEffect, useRef, useReducer } from "react";
import styles from "./SiteScriptEditor.module.scss";
import { TextField, PrimaryButton, Label, Stack, DefaultButton, ProgressIndicator, MessageBarType, CommandButton, IContextualMenuProps, Panel, PanelType, Pivot, PivotItem, Icon } from "office-ui-fabric-react";
import { useAppContext } from "../../app/App";
import { IApplicationState } from "../../app/ApplicationState";
import { ActionType, ISetAllAvailableSiteScripts, IGoToActionArgs, ISetUserMessageArgs } from "../../app/IApplicationAction";
import { SiteScriptDesigner } from "./SiteScriptDesigner";
import { SiteDesignsServiceKey } from "../../services/siteDesigns/SiteDesignsService";
import { ISiteScript, ISiteScriptContent } from "../../models/ISiteScript";
import CodeEditor, { monaco } from "@monaco-editor/react";
import { SiteScriptSchemaServiceKey } from "../../services/siteScriptSchema/SiteScriptSchemaService";
import { Confirm } from "../common/confirm/Confirm";
import { toJSON } from "../../utils/jsonUtils";
import { ExportServiceKey } from "../../services/export/ExportService";
import { ExportPackage } from "../../helpers/ExportPackage";
import { ExportPackageViewer } from "../exportPackageViewer/ExportPackageViewer";
import { SiteDesignPicker, NEW_SITE_DESIGN_KEY } from "../common/siteDesignPicker/SiteDesignPicker";
import { useTraceUpdate } from "../../helpers/hooks";
import { createNewSiteDesign } from "../../helpers/SiteDesignsHelpers";

export interface ISiteScriptEditorProps {
    siteScript: ISiteScript;
}

type ExportType = "json" | "PnPPowershell" | "PnPTemplate" | "o365_PS" | "o365_Bash";
type UpdatedCodeFrom = "UI" | "CODE" | "OTHER";

interface ISiteScriptEditorState {
    siteScriptMetadata: ISiteScript;
    siteScriptContent: ISiteScriptContent;
    updatedContentFrom: UpdatedCodeFrom;
    isValidCode: boolean;
    isSaving: boolean;
    isAssociatingToSiteDesign: boolean;
    isExportUIVisible: boolean;
    currentExportPackage: ExportPackage;
    currentExportType: ExportType;
}

interface ISetSiteScriptAction {
    type: "SET_SITE_SCRIPT";
    siteScript: ISiteScript;
}

interface IUpdateSiteScriptMetadataAction {
    type: "UPDATE_SITE_SCRIPT_METADATA";
    siteScript: ISiteScript;
}

interface IUpdateSiteScriptContentAction {
    type: "UPDATE_SITE_SCRIPT_CONTENT";
    content: ISiteScriptContent;
    from?: UpdatedCodeFrom;
    isValidCode?: boolean;
}

interface ISetIsSavingAction {
    type: "SET_ISSAVING";
    isSaving: boolean;
}

interface ISetExportPackageAction {
    type: "SET_EXPORTPACKAGE";
    exportPackage: ExportPackage;
    exportType?: ExportType;
}

interface ISetIsAssociatingToSiteDesign {
    type: "SET_ISASSOCIATINGTOSITEDESIGN";
    isAssociatingToSiteDesign: boolean;
}

type SiteScriptEditorAction = ISetSiteScriptAction
    | IUpdateSiteScriptMetadataAction
    | IUpdateSiteScriptContentAction
    | ISetIsSavingAction
    | ISetExportPackageAction
    | ISetIsAssociatingToSiteDesign;

const SiteScriptEditorReducer: (state: ISiteScriptEditorState, action: SiteScriptEditorAction) => ISiteScriptEditorState =
    (state, action) => {
        let updatedState = state;
        switch (action.type) {
            case "SET_SITE_SCRIPT":
                updatedState = {
                    ...state,
                    siteScriptMetadata: action.siteScript,
                    siteScriptContent: action.siteScript.Content
                };
                break;
            case "UPDATE_SITE_SCRIPT_METADATA":
                updatedState = {
                    ...state,
                    siteScriptMetadata: action.siteScript ? { ...state.siteScriptMetadata, ...action.siteScript } : state.siteScriptMetadata
                };
                break;
            case "UPDATE_SITE_SCRIPT_CONTENT":
                updatedState = {
                    ...state,
                    siteScriptContent: action.content ? { ...state.siteScriptContent, ...action.content } : state.siteScriptContent,
                    isValidCode: action.isValidCode || true,
                    updatedContentFrom: action.from || "OTHER"
                };
                break;
            case "SET_ISSAVING":
                updatedState = { ...state, isSaving: action.isSaving };
                break;
            case "SET_EXPORTPACKAGE":
                updatedState = {
                    ...state,
                    currentExportPackage: action.exportPackage,
                    currentExportType: action.exportPackage ? action.exportType : "json",
                    isExportUIVisible: !!action.exportPackage
                };
                break;
            case "SET_ISASSOCIATINGTOSITEDESIGN":
                updatedState = {
                    ...state,
                    isAssociatingToSiteDesign: action.isAssociatingToSiteDesign
                };
                break;
            default:
                if (DEBUG) {
                    console.debug("SiteScriptEditor:: state unchanged");
                }
                return state;
        }
        if (DEBUG) {
            console.debug("SiteScriptEditor:: state changed: ", updatedState, " due to action ", action);
        }
        return updatedState;
    };

export const SiteScriptEditor = (props: ISiteScriptEditorProps) => {
    useTraceUpdate('SiteScriptEditor', props);
    const [appContext, execute] = useAppContext<IApplicationState, ActionType>();

    // Get service references
    const siteDesignsService = appContext.serviceScope.consume(SiteDesignsServiceKey);
    const siteScriptSchemaService = appContext.serviceScope.consume(SiteScriptSchemaServiceKey);
    const exportService = appContext.serviceScope.consume(ExportServiceKey);

    const [state, dispatchState] = useReducer(SiteScriptEditorReducer, {
        siteScriptMetadata: null,
        siteScriptContent: null,
        currentExportPackage: null,
        currentExportType: "json",
        isExportUIVisible: false,
        isSaving: false,
        isAssociatingToSiteDesign: false,
        isValidCode: true,
        updatedContentFrom: null
    });
    const { siteScriptMetadata,
        siteScriptContent,
        isValidCode,
        isExportUIVisible,
        currentExportType,
        currentExportPackage,
        isAssociatingToSiteDesign,
        isSaving } = state;

    // Use refs
    const codeEditorRef = useRef<any>();
    const titleFieldRef = useRef<any>();
    const selectedSiteDesignRef = useRef<string>();


    const setLoading = (loading: boolean) => {
        execute("SET_LOADING", { loading });
    };
    // Use Effects
    useEffect(() => {
        if (!props.siteScript.Id) {
            dispatchState({ type: "SET_SITE_SCRIPT", siteScript: props.siteScript });

            if (titleFieldRef.current) {
                titleFieldRef.current.focus();
            }
            return;
        }

        setLoading(true);
        console.debug("Loading site script...", props.siteScript.Id);
        siteDesignsService.getSiteScript(props.siteScript.Id).then(loadedSiteScript => {
            dispatchState({ type: "SET_SITE_SCRIPT", siteScript: loadedSiteScript });
            console.debug("Loaded: ", loadedSiteScript);
        }).catch(error => {
            console.error(`The Site Script ${props.siteScript.Id} could not be loaded`, error);
        }).then(() => {
            setLoading(false);
        });
    }, [props.siteScript]);

    const onTitleChanged = (ev: any, title: string) => {
        const siteScript = { ...siteScriptMetadata, Title: title };
        dispatchState({ type: "UPDATE_SITE_SCRIPT_METADATA", siteScript });
    };

    let currentDescription = useRef<string>(siteScriptMetadata && siteScriptMetadata.Description);
    const onDescriptionChanging = (ev: any, description: string) => {
        currentDescription.current = description;
    };

    const onDescriptionInputBlur = (ev: any) => {
        const siteScript = { ...siteScriptMetadata, Description: currentDescription.current };
        dispatchState({ type: "UPDATE_SITE_SCRIPT_METADATA", siteScript });
    };

    const onVersionChanged = (ev: any, version: string) => {
        const versionInt = parseInt(version);
        if (!isNaN(versionInt)) {
            const siteScript = { ...siteScriptMetadata, Version: versionInt };
            dispatchState({ type: "UPDATE_SITE_SCRIPT_METADATA", siteScript });
        }
    };

    const onSiteScriptContentUpdatedFromUI = (content: ISiteScriptContent) => {
        dispatchState({ type: "UPDATE_SITE_SCRIPT_CONTENT", content, from: "UI" });
    };

    const onSave = async () => {
        dispatchState({ type: "SET_ISSAVING", isSaving: true });
        try {
            const isNewSiteScript = !siteScriptMetadata.Id;
            const toSave: ISiteScript = { ...siteScriptMetadata, Content: state.siteScriptContent };
            const updated = await siteDesignsService.saveSiteScript(toSave);
            const refreshedSiteScripts = await siteDesignsService.getSiteScripts();
            execute("SET_USER_MESSAGE", {
                userMessage: {
                    message: `${siteScriptMetadata.Title} has been successfully saved.`,
                    messageType: MessageBarType.success
                }
            } as ISetUserMessageArgs);
            execute("SET_ALL_AVAILABLE_SITE_SCRIPTS", { siteScripts: refreshedSiteScripts } as ISetAllAvailableSiteScripts);
            dispatchState({ type: "SET_SITE_SCRIPT", siteScript: updated });
            if (isNewSiteScript) {
                // Ask if the new Site Script should be associated to a Site Design
                if (await Confirm.show({
                    title: `Associate to Site Design`,
                    message: `Do you want to associate the new ${(siteScriptMetadata && siteScriptMetadata.Title)} to a Site Design ?`,
                    cancelLabel: 'No',
                    okLabel: 'Yes'
                })) {
                    dispatchState({ type: "SET_ISASSOCIATINGTOSITEDESIGN", isAssociatingToSiteDesign: true });
                }
            }
        } catch (error) {
            execute("SET_USER_MESSAGE", {
                userMessage: {
                    message: `${siteScriptMetadata.Title} could not be saved. Please make sure you have SharePoint administrator privileges...`,
                    messageType: MessageBarType.error
                }
            } as ISetUserMessageArgs);
            console.error(error);
        }
        dispatchState({ type: "SET_ISSAVING", isSaving: false });

    };

    const onDelete = async () => {
        if (!await Confirm.show({
            title: `Delete Site Script`,
            message: `Are you sure you want to delete ${(siteScriptMetadata && siteScriptMetadata.Title) || "this Site Script"} ?`
        })) {
            return;
        }

        dispatchState({ type: "SET_ISSAVING", isSaving: true });
        try {
            await siteDesignsService.deleteSiteScript(siteScriptMetadata);
            const refreshedSiteScripts = await siteDesignsService.getSiteScripts();
            execute("SET_USER_MESSAGE", {
                userMessage: {
                    message: `${siteScriptMetadata.Title} has been successfully deleted.`,
                    messageType: MessageBarType.success
                }
            } as ISetUserMessageArgs);
            execute("SET_ALL_AVAILABLE_SITE_SCRIPTS", { siteScripts: refreshedSiteScripts } as ISetAllAvailableSiteScripts);
            execute("GO_TO", { page: "SiteScriptsList" } as IGoToActionArgs);
        } catch (error) {
            execute("SET_USER_MESSAGE", {
                userMessage: {
                    message: `${siteScriptMetadata.Title} could not be deleted. Please make sure you have SharePoint administrator privileges...`,
                    messageType: MessageBarType.error
                }
            } as ISetUserMessageArgs);
            console.error(error);
        }
        dispatchState({ type: "SET_ISSAVING", isSaving: false });
    };

    const onAssociateSiteScript = () => {
        if (selectedSiteDesignRef.current != NEW_SITE_DESIGN_KEY) {
            execute("EDIT_SITE_DESIGN", { siteDesign: { Id: selectedSiteDesignRef.current }, additionalSiteScriptIds: [siteScriptMetadata.Id] });
        } else if (selectedSiteDesignRef.current) {
            execute("EDIT_SITE_DESIGN", { siteDesign: createNewSiteDesign(), additionalSiteScriptIds: [siteScriptMetadata.Id] });
        }
    };

    const onExportRequested = (exportType?: ExportType) => {
        const toExport: ISiteScript = { ...siteScriptMetadata, Content: siteScriptContent };
        let exportPromise: Promise<ExportPackage> = null;
        switch (exportType) {
            case "PnPPowershell":
                exportPromise = exportService.generateSiteScriptPnPPowershellExportPackage(toExport);
                break;
            case "PnPTemplate":
                break; // Not yet supported
            case "o365_PS":
                exportPromise = exportService.generateSiteScriptO365CLIExportPackage(toExport, "Powershell");
                break;
            case "o365_Bash":
                exportPromise = exportService.generateSiteScriptO365CLIExportPackage(toExport, "Bash");
                break;
            case "json":
            default:
                exportPromise = exportService.generateSiteScriptJSONExportPackage(toExport);
                break;
        }

        if (exportPromise) {
            exportPromise.then(exportPackage => {
                dispatchState({ type: "SET_EXPORTPACKAGE", exportPackage, exportType });
            });
        }
    };

    let codeUpdateTimeoutHandle: any = null;
    const onCodeChanged = (updatedCode: string) => {
        if (!updatedCode) {
            return;
        }

        if (codeUpdateTimeoutHandle) {
            clearTimeout(codeUpdateTimeoutHandle);
        }

        // if (updatedContentFrom == "UI") {
        //     // Not trigger the change of state if the script content was updated from UI
        //     console.debug("The code has been modified after a change in designer. The event will not be propagated");
        //     dispatchState({ type: "UPDATE_SITE_SCRIPT", siteScript: null, from: "CODE" });
        //     return;
        // }

        codeUpdateTimeoutHandle = setTimeout(() => {
            try {
                const jsonWithIgnoredComments = updatedCode.replace(/\/\*(.*)\*\//g,'');
                if (siteScriptSchemaService.validateSiteScriptJson(jsonWithIgnoredComments)) {
                    const content = JSON.parse(jsonWithIgnoredComments) as ISiteScriptContent;
                    dispatchState({ type: "UPDATE_SITE_SCRIPT_CONTENT", content, isValidCode: true, from: "CODE" });
                } else {
                    dispatchState({ type: "UPDATE_SITE_SCRIPT_CONTENT", content: null, isValidCode: false, from: "CODE" });

                }
            } catch (error) {
                console.warn("Code is not valid site script JSON");
            }
        }, 500);
    };

    const editorDidMount = (_, editor) => {

        const schema = siteScriptSchemaService.getSiteScriptSchema();
        codeEditorRef.current = editor;
        monaco.init().then(monacoApi => {
            monacoApi.languages.json.jsonDefaults.setDiagnosticsOptions({
                schemas: [{
                    uri: 'schema.json',
                    schema
                }],

                validate: true,
                allowComments: false
            });
        }).catch(error => {
            console.error("An error occured while trying to configure code editor");
        });

        const editorModel = editor.getModel();
        console.log("Editor model: ", editorModel);

        editor.onDidChangeModelContent(ev => {
            if (codeEditorRef && codeEditorRef.current) {
                onCodeChanged(codeEditorRef.current.getValue());
            }
        });
    };

    const checkIsValidForSave: () => [boolean, string?] = () => {
        if (!siteScriptMetadata) {
            return [false, "Current Site Script not defined"];
        }

        if (!siteScriptMetadata.Title) {
            return [false, "Please set the title of the Site Script..."];
        }

        if (!isValidCode) {
            return [false, "Please check the validity of the code..."];
        }

        return [true];
    };

    const isLoading = appContext.isLoading;
    const [isValidForSave, validationMessage] = checkIsValidForSave();
    if (!siteScriptMetadata || !siteScriptContent) {
        return null;
    }

    return <div className={styles.SiteScriptEditor}>
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
                            placeholder="Enter the name of the Site Script..."
                            borderless
                            componentRef={titleFieldRef}
                            value={siteScriptMetadata.Title}
                            onChange={onTitleChanged} />
                        {isLoading && <ProgressIndicator />}
                    </div>
                    {!isLoading && <div className={`${styles.column1} ${styles.righted}`}>
                        <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 15 }}>
                            <CommandButton disabled={isSaving} iconProps={{ iconName: "More" }} menuProps={{
                                items: [
                                    (siteScriptMetadata.Id && {
                                        key: 'deleteScript',
                                        text: 'Delete',
                                        iconProps: { iconName: 'Delete' },
                                        onClick: onDelete
                                    }),
                                    {
                                        key: 'export',
                                        text: 'Export',
                                        iconProps: { iconName: 'Download' },
                                        onClick: () => onExportRequested(),
                                        disabled: !isValidForSave
                                    }
                                ].filter(i => !!i),
                            } as IContextualMenuProps} />
                            <PrimaryButton disabled={isSaving || !isValidForSave} text="Save" iconProps={{ iconName: "Save" }} onClick={() => onSave()} />
                        </Stack>
                    </div>}
                </div>
                <div className={styles.row}>
                    {siteScriptMetadata.Id && <div className={styles.half}>
                        <div className={styles.row}>
                            <div className={styles.column8}>
                                <TextField
                                    label="Id"
                                    readOnly
                                    value={siteScriptMetadata.Id} />
                            </div>
                            <div className={styles.column4}>
                                <TextField
                                    label="Version"
                                    value={siteScriptMetadata.Version.toString()}
                                    onChange={onVersionChanged} />
                            </div>
                        </div>
                    </div>}
                    <div className={styles.half}>
                        <TextField
                            label="Description"
                            value={siteScriptMetadata.Description}
                            multiline={true}
                            rows={2}
                            borderless
                            placeholder="Enter a description for the Site Script..."
                            onChange={onDescriptionChanging}
                            onBlur={onDescriptionInputBlur}
                        />
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles.column}>
                        <Label>Actions</Label>
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles.designerWorkspace}>
                        <SiteScriptDesigner
                            siteScriptContent={siteScriptContent}
                            onSiteScriptContentUpdated={onSiteScriptContentUpdatedFromUI} />
                    </div>
                    <div className={styles.codeEditorWorkspace}>
                        <CodeEditor
                            height="80vh"
                            language="json"
                            options={{
                                folding: true,
                                renderIndentGuides: true,
                                minimap: {
                                    enabled: false
                                }
                            }}
                            value={toJSON(siteScriptContent)}
                            editorDidMount={editorDidMount}
                        />
                    </div>
                </div>
            </div>
        </div>
        {/* Association to a Site Design */}
        <Panel isOpen={isAssociatingToSiteDesign}
            type={PanelType.smallFixedFar}
            headerText="Associate to a Site Design"
            onRenderFooterContent={(p) => <Stack horizontalAlign="end" horizontal tokens={{ childrenGap: 10 }}>
                <PrimaryButton iconProps={{ iconName: "Check" }} text="Continue" onClick={() => onAssociateSiteScript()} />
                <DefaultButton text="Cancel" onClick={() => dispatchState({ type: "SET_ISASSOCIATINGTOSITEDESIGN", isAssociatingToSiteDesign: false })} /></Stack>}
        >
            <SiteDesignPicker serviceScope={appContext.serviceScope}
                label="Site Design"
                onSiteDesignSelected={(siteDesignId) => {
                    console.log("Selected site design: ", siteDesignId);
                    selectedSiteDesignRef.current = siteDesignId;
                }}
                hasNewSiteDesignOption
                displayPreview />
        </Panel>
        {/* Export options */}
        <Panel isOpen={isExportUIVisible}
            type={PanelType.large}
            headerText="Export Site Script"
            onRenderFooterContent={(p) => <Stack horizontalAlign="end" horizontal tokens={{ childrenGap: 10 }}>
                <PrimaryButton iconProps={{ iconName: "Download" }} text="Download" onClick={() => currentExportPackage && currentExportPackage.download()} />
                <DefaultButton text="Cancel" onClick={() => dispatchState({ type: "SET_EXPORTPACKAGE", exportPackage: null })} /></Stack>}>
            <Pivot
                selectedKey={currentExportType}
                onLinkClick={(item) => onExportRequested(item.props.itemKey as ExportType)}
                headersOnly={true}
            >
                <PivotItem headerText="JSON" itemKey="json" />
                <PivotItem headerText="PnP Powershell" itemKey="PnPPowershell" />
                {/* <PivotItem headerText="PnP Template" itemKey="PnPTemplate" /> */}
                <PivotItem headerText="O365 CLI (Powershell)" itemKey="o365_PS" />
                <PivotItem headerText="O365 CLI (Bash)" itemKey="o365_Bash" />
            </Pivot>
            {currentExportPackage && <ExportPackageViewer exportPackage={currentExportPackage} />}
        </Panel>
    </div >;
};