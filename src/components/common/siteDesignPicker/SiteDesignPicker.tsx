import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { ServiceScope } from '@microsoft/sp-core-library';
import { find } from "@microsoft/sp-lodash-subset";
import {
    Dropdown, IDropdownOption
} from 'office-ui-fabric-react/lib/Dropdown';
import { IDocumentCardPreviewProps, DocumentCardPreview } from "office-ui-fabric-react/lib/DocumentCard";
import { ImageFit } from "office-ui-fabric-react/lib/Image";
import { SiteDesignsServiceKey } from "../../../services/siteDesigns/SiteDesignsService";
import { ISiteDesign } from "../../../models/ISiteDesign";

export interface ISiteDesignPickerProps {
    serviceScope: ServiceScope;
    label: string;
    hasNewSiteDesignOption?: boolean;
    displayPreview?: boolean;
    onSiteDesignSelected: (siteDesignId: string) => void;
}

export const NEW_SITE_DESIGN_KEY = "###NEW_SITE_DESIGN###";
const LOADING_KEY = "###LOADING###";

export const SiteDesignPicker = (props: ISiteDesignPickerProps) => {

    const siteDesignsService = props.serviceScope.consume(SiteDesignsServiceKey);
    const [state, setState] = useState<{ siteDesignsOptions: IDropdownOption[]; selectedSiteDesign: ISiteDesign, isLoading: boolean }>({
        siteDesignsOptions: [
            {
                key: LOADING_KEY,
                text: "Loading..."
            }
        ],
        selectedSiteDesign: null,
        isLoading: true
    });
    const loadedSiteDesigns = useRef<ISiteDesign[]>([]);
    useEffect(() => {
        siteDesignsService.getSiteDesigns().then(siteDesigns => {
            const newItemOptionArray = props.hasNewSiteDesignOption ? [{ key: NEW_SITE_DESIGN_KEY, text: "New Site Design" }] : [];
            const siteDesignsOptions = newItemOptionArray.concat(...siteDesigns.map(sd => ({ key: sd.Id, text: sd.Title })));
            loadedSiteDesigns.current = siteDesigns;
            setState({
                siteDesignsOptions,
                selectedSiteDesign: null,
                isLoading: false
            });
        }).catch(error => {
            console.error("Site Designs cannot be loaded.", error);
            setState({
                siteDesignsOptions: [],
                selectedSiteDesign: null,
                isLoading: false
            });
        });
    }, []);

    const onSelectedSiteDesign = (ev: any, option: IDropdownOption) => {
        const selectedSiteDesign = find(loadedSiteDesigns.current, sd => sd.Id == option.key);
        setState({
            siteDesignsOptions: state.siteDesignsOptions,
            isLoading: state.isLoading,
            selectedSiteDesign
        });
        props.onSiteDesignSelected(option.key as string);
    };

    const previewProps: IDocumentCardPreviewProps = {
        previewImages: [
            {
                previewImageSrc: state.selectedSiteDesign && state.selectedSiteDesign.PreviewImageUrl,
                imageFit: ImageFit.centerContain,
                height: 300
            }
        ]
    };

    return <div className={""/* styles.SitePicker */}>
        <Dropdown label={props.label}
            options={state.siteDesignsOptions}
            onChange={onSelectedSiteDesign} />
        {props.displayPreview && state.selectedSiteDesign && <div>
            <DocumentCardPreview {...previewProps} />
            <pre>{state.selectedSiteDesign.Description}</pre>
        </div>}
    </div>;
};