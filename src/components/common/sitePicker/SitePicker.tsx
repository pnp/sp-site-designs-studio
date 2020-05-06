import * as React from "react";
import {
    TagPicker,
    ITag,
} from 'office-ui-fabric-react/lib/Pickers';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { ISPSite } from "../../../models/ISPSite";
import { SitesServiceKey } from "../../../services/sites/SitesService";
import styles from "./SitePicker.module.scss";
import { ServiceScope } from '@microsoft/sp-core-library';
import { Label } from "office-ui-fabric-react/lib/Label";
import { useState } from "react";
import { TextField, IconButton } from "office-ui-fabric-react";

export interface ISitePickerProps {
    serviceScope: ServiceScope;
    label: string;
    currentSiteSelectedByDefault?: boolean;
    onSiteSelected: (siteUrl: string) => void;
}

const SuggestedSiteItem = (props: ITag) => {

    // TODO Improve initial resolving
    const siteInitials = props.name && props.name.length && props.name.toUpperCase()[0];

    return <div>
        <Stack horizontal tokens={{ childrenGap: 5 }}>
            <div className={styles.siteSquare}>{siteInitials}</div>
            {props.name}
        </Stack>
    </div>;
};

export const SitePicker = (props: ISitePickerProps) => {

    const sitesService = props.serviceScope.consume(SitesServiceKey);
    const [selectedSite, setSelectedSite] = useState<ISPSite>(null);
    const [isSearching, setIsSearching] = useState<boolean>(true);
    const onFilterChanged = async (filterText: string) => {
        return filterText
            ? (await sitesService.getSiteByNameOrUrl(filterText)).map(s => ({ key: s.url, name: s.title } as ITag))
            : [];
    };

    const onSelectedSiteChanged = (site: ISPSite) => {
        setSelectedSite(site);
        setIsSearching(false);
        props.onSiteSelected((site && site.url) || "");
    };

    const onSelectedTagChanged = (tag: ITag) => {
        onSelectedSiteChanged({ title: tag.name, url: tag.key, id: null });
        return tag;
    };

    // return <div className={styles.SitePicker}>
    return <div className={""/* styles.SitePicker */}>

        {!isSearching
            ? <Stack horizontal tokens={{ childrenGap: 5 }}>
                <TextField label={props.label} value={selectedSite && selectedSite.url} onChange={(_, v) => onSelectedSiteChanged({ url: v, title: '', id: '' })} />
                <IconButton styles={{ root: { position: "relative", top: 30 } }} iconProps={{ iconName: "SearchAndApps" }} onClick={() => setIsSearching(true)} />
            </Stack>
            : <div>{props.label && <Label>{props.label}</Label>}
                <Stack horizontal tokens={{ childrenGap: 5 }}>
                    <TagPicker
                        removeButtonAriaLabel="Remove"
                        onRenderSuggestionsItem={SuggestedSiteItem as any}
                        onResolveSuggestions={onFilterChanged}
                        pickerSuggestionsProps={{
                            suggestionsHeaderText: 'Suggested sites',
                            noResultsFoundText: 'No sites Found',
                        }}
                        resolveDelay={800}
                        itemLimit={1}
                        onItemSelected={onSelectedTagChanged}
                    />
                <IconButton styles={{ root: { position: "relative" } }} iconProps={{ iconName: "Edit" }} onClick={() => setIsSearching(false)} />
                </Stack>
            </div>}
    </div>;
};