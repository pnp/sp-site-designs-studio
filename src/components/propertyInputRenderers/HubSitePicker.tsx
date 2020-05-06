import * as React from 'react';
import { ComboBox, IComboBoxOption, SelectableOptionMenuItemType } from 'office-ui-fabric-react';
import { useAppContext } from "../../app/App";
import { IApplicationState } from "../../app/ApplicationState";
import { ActionType } from "../../app/IApplicationAction";
import { IHubSite } from '../../models/IHubSite';
import { HubSitesServiceKey } from '../../services/hubSites/HubSitesService';
import { IRenderer } from '../../services/rendering/RenderingService';
import { useState, useEffect } from 'react';

export interface IHubSitePickerProps {
    value: string;
    label: string;
    onValueChanged: (value: string) => void;
}

const loadingOptionKey = 'loadingHubSites';
const noAvailableOptionKey = 'NoAvailableHubSites';


export const HubSitePicker = (props: IHubSitePickerProps) => {

    const [appContext] = useAppContext<IApplicationState, ActionType>();
    const [availableHubSites, setAvailableHubSites] = useState<IHubSite[]>([]);
    const [isLoading, setIsLoading] = useState<boolean>(true);
    const hubSitesService = appContext.serviceScope.consume(HubSitesServiceKey);

    const loadHubSites = async () => {
        setIsLoading(true);
        try {
            const hubSites = await hubSitesService.getHubSites();
            setAvailableHubSites(hubSites);
        } catch (error) {
            console.error(error);
        }
        setIsLoading(false);
    };

    const getHubSites = () => {
        if (isLoading) {
            return [{ key: loadingOptionKey, text: "Loading available hub sites..." }];
        }

        if (availableHubSites.length == 0) {
            return [{ key: noAvailableOptionKey, text: "No available hub sites on this tenant..." }];
        }

        return availableHubSites.map(hs => ({ key: hs.id, text: hs.title }));
    };

    const onValueChange = (ev: any, option: IComboBoxOption, index: number, value: string) => {
        if (option && option.key) {
            // Only if selected options is not the displayed "loading" or "no available hub sites"
            if ([loadingOptionKey, noAvailableOptionKey].indexOf(option.key.toString()) < 0) {
                props.onValueChanged(option.key.toString());
            }
        } else if (value) {
            props.onValueChanged(value);
        }
    };

    useEffect(() => {
        loadHubSites();
    }, []);

    return <ComboBox
        label={props.label}
        allowFreeform={true}
        autoComplete="off"
        selectedKey={props.value}
        options={getHubSites()}
        onChange={onValueChange}
    />;
};

export const hubSitePickerRenderer: IRenderer = (label: string, value: any, onValueChange: (changedValue: any) => void) => {
    return <HubSitePicker label={label} value={value} onValueChanged={onValueChange} />;
};
