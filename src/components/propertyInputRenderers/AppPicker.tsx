import * as React from 'react';
import { ComboBox, IComboBoxOption, SelectableOptionMenuItemType } from 'office-ui-fabric-react';
import { useAppContext } from "../../app/App";
import { IApplicationState } from "../../app/ApplicationState";
import { ActionType } from "../../app/IApplicationAction";
import { IApp } from '../../models/IApp';
import { AppsServiceKey } from '../../services/apps/AppsService';
import { IRenderer } from '../../services/rendering/RenderingService';
import { useState, useEffect } from 'react';

export interface IAppPickerProps {
	value: string;
	label: string;
	onValueChanged: (value: string) => void;
}

const loadingOptionKey = 'loadingApps';
const noAvailableOptionKey = 'NoAvailableApps';


export const AppPicker = (props:IAppPickerProps) => {
    const [appContext] = useAppContext<IApplicationState, ActionType>();
    const [availableApps, setAvailableApps] = useState<IApp[]>([]);
    const [isLoading, setIsLoading] = useState<boolean>(true);
    const appsService = appContext.serviceScope.consume(AppsServiceKey);

    const loadAvailableApps = async () => {
        setIsLoading(true);
        try {
            const apps = await appsService.getAvailableApps();
            setAvailableApps(apps);
        } catch (error) {
            console.error(error);
        }
        setIsLoading(false);
    };

    const getAvailableApps = () => {
        if (isLoading) {
            return [{ key: loadingOptionKey, text: "Loading available apps..." }];
        }

        if (availableApps.length == 0) {
            return [{ key: noAvailableOptionKey, text: "No available apps in this tenant app catalog..." }];
        }

        return availableApps.map(hs => ({ key: hs.id, text: hs.title }));
    };

    const onValueChange = (ev: any, option: IComboBoxOption, index: number, value: string) => {
        if (option && option.key) {
            // Only if selected options is not the displayed "loading" or "no available apps"
            if ([loadingOptionKey, noAvailableOptionKey].indexOf(option.key.toString()) < 0) {
                props.onValueChanged(option.key.toString());
            }
        } else if (value) {
            props.onValueChanged(value);
        }
    };

    useEffect(() => {
        loadAvailableApps();
    }, []);

    return <ComboBox
        label={props.label}
        allowFreeform={true}
        autoComplete="off"
        selectedKey={props.value}
        options={getAvailableApps()}
        onChange={onValueChange}
    />;
};

export const appPickerRenderer: IRenderer = (label: string, value: any, onValueChange: (changedValue: any) => void) => {
    return <AppPicker label={label} value={value} onValueChanged={onValueChange} />;
};
