import * as React from 'react';
import { ComboBox, IComboBoxOption, SelectableOptionMenuItemType } from 'office-ui-fabric-react';
import { useAppContext } from "../../app/App";
import { IApplicationState } from "../../app/ApplicationState";
import { ActionType } from "../../app/IApplicationAction";
import { ITheme } from '../../models/ITheme';
import { ThemeServiceKey } from '../../services/themes/ThemesService';
import { IRenderer } from '../../services/rendering/RenderingService';
import { useState, useEffect } from 'react';

export interface IThemePickerProps {
    value: string;
    label: string;
    onValueChanged: (value: string) => void;
}

const loadingOptionKey = 'loadingThemes';
const noAvailableOptionKey = 'noAvailableThemes';


export const ThemePicker = (props: IThemePickerProps) => {
    const [appContext] = useAppContext<IApplicationState, ActionType>();
    const [availableCustomThemes, setAvailableCustomThemes] = useState<ITheme[]>([]);
    const [isLoading, setIsLoading] = useState<boolean>(true);
    const themesService = appContext.serviceScope.consume(ThemeServiceKey);

    const loadAvailableThemes = async () => {
        setIsLoading(true);
        try {
            const themes = await themesService.getCustomThemes();
            setAvailableCustomThemes(themes);
        } catch (error) {
            console.error(error);
        }
        setIsLoading(false);
    };

    const getAvailableThemes = () => {
        if (isLoading) {
            return [{ key: loadingOptionKey, text: "Loading available custom themes..." }];
        }

        if (availableCustomThemes.length == 0) {
            return [{ key: noAvailableOptionKey, text: "No available custom themes in this tenant..." }];
        }

        return availableCustomThemes.map(hs => ({ key: hs.name, text: hs.name }));
    };

    const onValueChange = (ev: any, option: IComboBoxOption, index: number, value: string) => {
        if (option && option.key) {
            // Only if selected options is not the displayed "loading" or "no available themes"
            if ([loadingOptionKey, noAvailableOptionKey].indexOf(option.key.toString()) < 0) {
                props.onValueChanged(option.key.toString());
            }
        } else if (value) {
            props.onValueChanged(value);
        }
    };

    useEffect(() => {
        loadAvailableThemes();
    }, []);

    return <ComboBox
        label={props.label}
        allowFreeform={true}
        autoComplete="off"
        selectedKey={props.value}
        options={getAvailableThemes()}
        onChange={onValueChange}
    />;
};

export const themePickerRenderer: IRenderer = (label: string, value: any, onValueChange: (changedValue: any) => void) => {
    return <ThemePicker label={label} value={value} onValueChanged={onValueChange} />;
};
