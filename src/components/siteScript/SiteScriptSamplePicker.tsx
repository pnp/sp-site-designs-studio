import * as React from "react";
import { useState } from "react";
import { DocumentCard, Icon, DocumentCardDetails, DocumentCardTitle, DocumentCardType, ISize, Dropdown, IDropdownOption } from "office-ui-fabric-react";

import styles from "./NewSiteScriptPanel.module.scss";
import { GridLayout } from "@pnp/spfx-controls-react/lib/GridLayout";
import { ISiteScriptSample } from "../../models/ISiteScriptSample";
import { IGithubRepository } from "../../models/ISiteScriptSamplesRepository";
import { find } from "@microsoft/sp-lodash-subset";


export interface ISiteScriptSamplePickerProps {
    onSelectedSample: (sample: ISiteScriptSample) => void;
}

interface ISiteScriptSamplePickerState {
    availableSamples: ISiteScriptSample[];
    selectedSample: ISiteScriptSample;
    previewedSample: ISiteScriptSample;
    avilableRepositories: IGithubRepository[];
    currentRepository: IGithubRepository;
    isLoading: boolean;
}

export const SiteScriptSamplePicker = (props: ISiteScriptSamplePickerProps) => {
    const [state, setState] = useState<ISiteScriptSamplePickerState>({
        isLoading: true,
        availableSamples: [],
        avilableRepositories: [],
        currentRepository: null,
        selectedSample: null,
        previewedSample: null
    });

    if (state.isLoading) {
        return null;
    }

    const onSelectedSample = (sample: ISiteScriptSample) => {
        setState({ ...state, selectedSample: sample });
        if (props.onSelectedSample) {
            props.onSelectedSample(sample);
        }
    };

    const renderSampleItem = (sample: ISiteScriptSample, finalSize: ISize, isCompact: boolean): JSX.Element => {

        return <div
            data-is-focusable={true}
            role="listitem"
            aria-label={sample.title}
        >
            <DocumentCard
                onMouseEnter={_ => setState({ ...state, previewedSample: sample })}
                type={isCompact ? DocumentCardType.compact : DocumentCardType.normal}
                onClick={(ev: React.SyntheticEvent<HTMLElement>) => onSelectedSample(sample)}>
                <div className={styles.iconBox}>
                    <div className={styles.icon}>
                        <Icon iconName="Script" />
                    </div>
                </div>
                <DocumentCardDetails>
                    <DocumentCardTitle
                        title={sample.title}
                        shouldTruncate={true}
                    />
                </DocumentCardDetails>
            </DocumentCard>
        </div>;
    };

    const onRepositoryChange = (option: IDropdownOption) => {
        if (option && option.key) {
            const selectedRepo = find(state.avilableRepositories, r => r.key == option.key);
            if (selectedRepo) {
                // TODO Reload samples from selected repository
                setState({ ...state, currentRepository: selectedRepo });
            }
        }
    };

    return <div className={styles.row}>
        {state.avilableRepositories && state.avilableRepositories.length > 1 && <div className={`${styles.column}`}>
            <Dropdown label="Repository"
                options={state.avilableRepositories.map(r => ({ key: r.key, text: r.key }))}
                onChange={(ev, option) => } />
        </div>}
        <div className={`${styles.column} ${styles.column8}`}>
            <GridLayout
                ariaLabel="List of Site Scripts samples."
                items={state.availableSamples}
                onRenderGridItem={renderSampleItem}
            />
        </div>
        <div className={`${styles.column} ${styles.column8}`}>
            <div key="sample-readme">

            </div>
        </div>
    </div>;
};