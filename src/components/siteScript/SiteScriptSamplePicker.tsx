import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { DocumentCard, Icon, DocumentCardDetails, DocumentCardTitle, DocumentCardType, ISize, Dropdown, IDropdownOption, ProgressIndicator, MessageBarType, Pivot, PivotItem, MessageBar, SearchBox, Stack } from "office-ui-fabric-react";

import styles from "./NewSiteScriptPanel.module.scss";
import { useAppContext } from "../../app/App";
import { IApplicationState } from "../../app/ApplicationState";
import { GridLayout } from "@pnp/spfx-controls-react/lib/GridLayout";
import { ISiteScriptSample } from "../../models/ISiteScriptSample";
import { ISiteScriptSamplesRepository } from "../../models/ISiteScriptSamplesRepository";
import { SiteScriptSamplesServiceKey } from "../../services/siteScriptSamples/SiteScriptSamplesService";
import CodeEditor from "@monaco-editor/react";

export interface ISiteScriptSamplePickerProps {
    selectedSample: ISiteScriptSample;
    onSelectedSample: (sample: ISiteScriptSample) => void;
}

interface ISelectableSiteScriptSample {
    sample: ISiteScriptSample;
    selected: boolean;
}

export const SiteScriptSamplePicker = (props: ISiteScriptSamplePickerProps) => {
    const [appContext, execute] = useAppContext<IApplicationState, any>();
    const samplesService = appContext.serviceScope.consume(SiteScriptSamplesServiceKey);

    const [repo, setRepo] = useState<{
        availableRepositories: ISiteScriptSamplesRepository[];
        currentRepository: ISiteScriptSamplesRepository;
    }>({
        availableRepositories: [],
        currentRepository: null
    });
    const [selectedSampleKey, setSelectedSampleKey] = useState<string>(null);
    const [availableSamples, setAvailableSamples] = useState<ISiteScriptSample[]>([]);
    const [isLoadingAllSamples, setIsLoadingAllSamples] = useState<boolean>(true);
    const [isLoadingSampleDetails, setIsLoadingSampleDetails] = useState<boolean>(false);
    const [searchCriteria, setSearchCriteria] = useState<string>('');

    const getSamples = async (repository: ISiteScriptSamplesRepository) => {
        repository = repository || repo.currentRepository;
        return repository ? await samplesService.getSamples(repository) : [];
    };

    const initialLoad = async () => {
        try {
            const availableRepositories = await samplesService.getAvailableRepositories();
            // Select the first available repository and get the samples from it
            const currentRepository = availableRepositories && availableRepositories.length > 0 ? availableRepositories[0] : null;
            setRepo({ currentRepository, availableRepositories });

            const foundSamples = await getSamples(currentRepository);
            setAvailableSamples(foundSamples);
            setIsLoadingAllSamples(false);
        } catch (error) {
            setIsLoadingAllSamples(false);
            // TODO Determine the reason of the error
            execute("SET_USER_MESSAGE", {
                userMessage: {
                    message: "All samples cannot be loaded...",
                    userMessageType: MessageBarType.error
                }
            });
        }

    };

    const justMounted = useRef<boolean>(false);
    useEffect(() => {
        initialLoad();
        justMounted.current = true;
    }, []);

    const loadSampleData = async (sample: ISiteScriptSample) => {
        try {
            setIsLoadingSampleDetails(true);
            const loadedSample = await samplesService.getSample(sample);
            setIsLoadingSampleDetails(false);
            return loadedSample;
        } catch (error) {
            setIsLoadingSampleDetails(false);
            execute("SET_USER_MESSAGE", {
                userMessage: {
                    message: "Sample cannot be loaded...",
                    userMessageType: MessageBarType.error
                }
            });
            return null;
        }
    };

    const selectSample = async (sample: ISiteScriptSample) => {
        setSelectedSampleKey(sample.key);
        const loadedSample = await loadSampleData(sample);
        if (props.onSelectedSample) {
            props.onSelectedSample(loadedSample);
        }
    };

    const renderSampleItem = (item: ISelectableSiteScriptSample, finalSize: ISize, isCompact: boolean): JSX.Element => {

        return <div
            data-is-focusable={true}
            role="listitem"
            aria-label={item.sample.key}
            className={item.selected ? styles.selectedSample : ''}
        >
            <DocumentCard
                type={isCompact ? DocumentCardType.compact : DocumentCardType.normal}
                onClick={(ev: React.SyntheticEvent<HTMLElement>) => selectSample(item.sample)}>
                <div className={styles.iconBox}>
                    <div className={styles.icon}>
                        <Icon iconName="Script" />
                    </div>
                </div>
                <DocumentCardDetails>
                    <DocumentCardTitle
                        title={item.sample.key}
                        shouldTruncate={true}
                    />
                </DocumentCardDetails>
            </DocumentCard>
        </div>;
    };

    return <div className={styles.row}>
        {isLoadingAllSamples && <div className={styles.column}>
            <ProgressIndicator label="Loading..." />
        </div>}
        {!isLoadingAllSamples && <div className={`${selectedSampleKey ? styles.column6 : styles.column}`}>
            <Stack tokens={{ childrenGap: 3 }}>
                <SearchBox underlined value={searchCriteria} onChange={setSearchCriteria} placeholder={"Search a sample..."} />
                <div className={styles.sampleGallery}>
                    <GridLayout
                        ariaLabel="List of Site Scripts samples."
                        items={availableSamples
                            .filter(s => !searchCriteria || s.key.toLowerCase().indexOf(searchCriteria.toLowerCase()) > -1)
                            .map(s => ({ sample: s, selected: s.key == selectedSampleKey }))}
                        onRenderGridItem={renderSampleItem}
                    />
                </div>
            </Stack>
        </div>}
        {selectedSampleKey && <div className={`${styles.column6}`}>
            <div key="sample-readme">
                {isLoadingSampleDetails ? <ProgressIndicator label="Loading..." />
                    : <div>
                        {props.selectedSample &&
                            <div>
                                <Pivot>
                                    <PivotItem itemKey="readme" headerText="README">
                                        {props.selectedSample.readmeHtml
                                            ? <div className={styles.readmeViewer} dangerouslySetInnerHTML={{ __html: props.selectedSample.readmeHtml }} />
                                            : <MessageBar messageBarType={MessageBarType.warning}>
                                                {"README of this sample cannot be displayed..."}
                                            </MessageBar>}
                                    </PivotItem>
                                    <PivotItem itemKey="json" headerText="JSON">
                                        {props.selectedSample.jsonContent
                                            ? <div>
                                                {props.selectedSample.hasPreprocessedJsonContent && <MessageBar messageBarType={MessageBarType.warning}>
                                                    {"The JSON of this sample has been preprocessed..."}<br />
                                                    {`Visit`}<a href={props.selectedSample.webSite} target="_blank">{`this page`}</a> {` to view the original sample`}
                                                </MessageBar>}
                                                <CodeEditor
                                                    height="70vh"
                                                    language="json"
                                                    options={{
                                                        readOnly: true,
                                                        folding: true,
                                                        renderIndentGuides: true,
                                                        minimap: {
                                                            enabled: false
                                                        }
                                                    }}
                                                    value={props.selectedSample.jsonContent}
                                                />
                                            </div>
                                            : <MessageBar messageBarType={MessageBarType.warning}>
                                                {"A usable JSON file cannot be found in this sample..."}<br />
                                                {`Visit`}<a href={props.selectedSample.webSite} target="_blank">{`this page`}</a> {` to explore this sample`}
                                            </MessageBar>}
                                    </PivotItem>
                                </Pivot>
                            </div>
                        }
                    </div>}
            </div>
        </div>}
    </div>;
};