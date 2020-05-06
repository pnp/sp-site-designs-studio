import * as React from 'react';
import { ServiceScope } from '@microsoft/sp-core-library';
import {
    CompactPeoplePicker,
    IBasePickerSuggestionsProps,
    NormalPeoplePicker
} from 'office-ui-fabric-react/lib/Pickers';
import { find } from 'office-ui-fabric-react/lib/Utilities';
import { IPersonaProps, IPersona } from 'office-ui-fabric-react/lib/Persona';
import { TenantServiceKey } from "../../../services/tenant/tenantService";
import { IPrincipal } from '../../../models/IPrincipal';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { useState, useEffect } from 'react';



export enum PrincipalType {
    User = 1,
    Group = 2,
    All = User | Group
}

export interface IPeoplePickerProps {
    serviceScope: ServiceScope;
    label: string;
    selectedItems?: IPrincipal[];
    typePicker: "normal" | "compact";
    principalType: PrincipalType;
    onChange?: (items: IPrincipal[]) => void;
}

const suggestionProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: 'Suggested People',
    noResultsFoundText: 'No results found',
    loadingText: 'Loading'
};

export const PeoplePicker = (props: IPeoplePickerProps) => {


    const [ensuredSelectedItems, setEnsuredSelectedItems] = useState<IPrincipal[]>(props.selectedItems || []);
    const tenant = props.serviceScope.consume(TenantServiceKey);

    const mapPrincipalToPersonaProps = (p: IPrincipal) => ({
        text: p.displayName,
        secondaryText: p.principalName,
        ["$$principalType"]: p.type
    } as IPersonaProps);
    const mapPersonaPropsToPrincipal = (p: IPersonaProps) => ({
        displayName: p.text,
        principalName: p.secondaryText,
        type: p["$$principalType"]
    } as IPrincipal);

    useEffect(() => {
        if (props.selectedItems) {
            tenant.ensurePrincipalInfo(props.selectedItems).then(ensured => {
                setEnsuredSelectedItems(ensured);
            });
        }

    }, [props.selectedItems]);

    const searchPeople: (terms: string, toSkip: IPrincipal[]) => IPersonaProps[] | Promise<IPersonaProps[]> = async (terms, toSkip) => {

        let found = await tenant.searchPrincipals(terms);
        if (toSkip && toSkip.length > 0) {
            found = found.filter(p => !find(toSkip, ts => ts.principalName == p.principalName));
        }
        return found.map(mapPrincipalToPersonaProps);
    };

    const onFilterChanged = (filterText: string) => {
        if (filterText) {
            if (filterText.length > 2) {
                return searchPeople(filterText, props.selectedItems);
            }
        } else {
            return [];
        }
    };

    const onChange = (changedItems: IPersonaProps[]) => {
        if (props.onChange) {
            props.onChange(changedItems.map(mapPersonaPropsToPrincipal));
        }
    };

    const selectedItems = ensuredSelectedItems.map(mapPrincipalToPersonaProps);

    if (props.typePicker == "normal") {
        return <>
            {props.label && <Label>{props.label}</Label>}
            <NormalPeoplePicker
                onChange={onChange}
                onResolveSuggestions={onFilterChanged}
                getTextFromItem={(persona: IPersonaProps) => persona.text}
                pickerSuggestionsProps={suggestionProps}
                selectedItems={selectedItems}
                className={'ms-PeoplePicker'}
            />
        </>;
    }
    else {
        return <>
            {props.label && <Label>{props.label}</Label>}
            <CompactPeoplePicker
                onChange={onChange}
                onResolveSuggestions={onFilterChanged}
                getTextFromItem={(persona: IPersonaProps) => persona.text}
                pickerSuggestionsProps={suggestionProps}
                selectedItems={selectedItems}
                className={'ms-PeoplePicker'}
            />
        </>;
    }
};