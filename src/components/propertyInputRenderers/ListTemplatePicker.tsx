import * as React from 'react';
import { IDropdownOption, Dropdown } from 'office-ui-fabric-react';
import { IRenderer } from '../../services/rendering/RenderingService';

export interface IListTemplatePickerProps {
    value: string;
    label: string;
    onValueChanged: (value: any) => void;
}

export const ListTemplatePicker = (props: IListTemplatePickerProps) => {


    const getSupportedListTemplates: () => IDropdownOption[] = () => [
        { key: 100, text: 'Generic List' },
        { key: 101, text: 'Document Library' },
        { key: 102, text: 'Survey' },
        { key: 103, text: 'Links' },
        { key: 104, text: 'Announcements' },
        { key: 105, text: 'Contacts' },
        { key: 106, text: 'Events' },
        { key: 107, text: 'Tasks' },
        { key: 108, text: 'Discussion Board' },
        { key: 109, text: 'Picture Library' },
        { key: 119, text: 'Site Pages' },
        { key: 1100, text: 'Issues Tracking' }
    ];

    const onValueChange = (ev: any, option: IDropdownOption) => {
        props.onValueChanged(option.key);
    };

    return <Dropdown
        label={props.label}
        selectedKey={props.value}
        options={getSupportedListTemplates()}
        onChange={onValueChange}
    />;
};

export const listTemplatePickerRenderer: IRenderer = (label: string, value: any, onValueChange: (changedValue: any) => void) => {
    return <ListTemplatePicker label={label} value={value} onValueChanged={onValueChange} />;
};
