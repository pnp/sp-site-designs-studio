import * as React from "react";
import { useEffect, useState, useRef } from "react";
import { ServiceScope } from '@microsoft/sp-core-library';
import { Dropdown, IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";
import { IList } from "../../../models/IList";
import { SitesServiceKey } from "../../../services/sites/SitesService";
import { find } from "office-ui-fabric-react";

export interface IListPickerProps {
    label?: string;
    webUrl: string;
    serviceScope: ServiceScope;
    multiselect?: boolean;
    onListSelected?: (list: IList) => void;
    onListsSelected?: (lists: IList[]) => void;
}

export const ListPicker = (props: IListPickerProps) => {

    const [availableLists, setAvailableLists] = useState<IList[]>([]);
    const selectedItemsRef = useRef<IDropdownOption[]>([]);

    useEffect(() => {
        // Load the lists from the specified web
        const sitesService = props.serviceScope.consume(SitesServiceKey);
        sitesService.getSiteLists(props.webUrl).then(lists => setAvailableLists(lists));
    }, [props.webUrl]);

    const getSelectedLists = () => {
        if (!selectedItemsRef.current) {
            return [];
        }

        return selectedItemsRef.current.map(option => find(availableLists, l => l.url == option.key as string));
    };

    const onChange = (ev: any, option: IDropdownOption, index: number) => {
        selectedItemsRef.current = selectedItemsRef.current.filter(s => s.key != option.key);

        if (props.multiselect && props.onListsSelected) {
            if (option.selected) {
                selectedItemsRef.current.push(option);
            } else {
                selectedItemsRef.current.splice(index, 1);
            }
            const selectedLists = getSelectedLists();
            props.onListsSelected(selectedLists);
        } else if (props.onListSelected) {

            if (option) {
                selectedItemsRef.current = [option];
            }
            const selectedLists = getSelectedLists();
            props.onListSelected(selectedLists.length > 0 ? selectedLists[0] : null);
        }
    };

    return <Dropdown
        disabled={!availableLists || (availableLists && availableLists.length == 0)}
        label={props.label}
        multiSelect={props.multiselect}
        options={availableLists.map(al => ({ key: al.url, text: al.title }))}
        onChange={onChange} />;
};