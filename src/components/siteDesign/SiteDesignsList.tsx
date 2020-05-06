import * as React from "react";
import {
    DocumentCard,
    DocumentCardPreview,
    DocumentCardDetails,
    DocumentCardTitle,
    IDocumentCardPreviewProps,
    DocumentCardType
} from 'office-ui-fabric-react/lib/DocumentCard';
import { ImageFit } from 'office-ui-fabric-react/lib/Image';
import { ISize } from 'office-ui-fabric-react/lib/Utilities';
import { GridLayout } from "@pnp/spfx-controls-react/lib/GridLayout";
import { ISiteDesign } from "../../models/ISiteDesign";
import styles from "./SiteDesignsList.module.scss";
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { Link } from "office-ui-fabric-react/lib/Link";

export interface ISiteDesignsListAllOptionalProps {
    siteDesigns?: ISiteDesign[];
    preview?: boolean;
    addNewDisabled?: boolean;
    onSiteDesignClicked?: (siteDesign: ISiteDesign) => void;
    onAdd?: () => void;
    onSeeMore?: () => void;
}

export interface ISiteDesignsListProps extends ISiteDesignsListAllOptionalProps {
    siteDesigns: ISiteDesign[];
}

const PREVIEW_ITEMS_COUNT = 3;

export const SiteDesignsList = (props: ISiteDesignsListProps) => {

    const renderGridItem = (siteDesign: ISiteDesign, finalSize: ISize, isCompact: boolean): JSX.Element => {

        if (!siteDesign) {
            // If site script is not set, it is the Add new tile
            return <div
                className={styles.add}
                data-is-focusable={true}
                role="listitem"
                aria-label={"Add a new Site Design"}
            >
                <DocumentCard
                    type={isCompact ? DocumentCardType.compact : DocumentCardType.normal}
                    onClick={(ev: React.SyntheticEvent<HTMLElement>) => props.onAdd && props.onAdd()}>
                    <div className={styles.iconBox}>
                        <div className={styles.icon}>
                            <Icon iconName="Add" />
                        </div>
                    </div>
                </DocumentCard>
            </div>;
        }


        const previewProps: IDocumentCardPreviewProps = {
            previewImages: [
                {
                    previewImageSrc: siteDesign.PreviewImageUrl,
                    imageFit: ImageFit.cover,
                    height: 130
                }
            ]
        };

        return <div
            data-is-focusable={true}
            role="listitem"
            aria-label={siteDesign.Title}
        >
            <DocumentCard
                type={isCompact ? DocumentCardType.compact : DocumentCardType.normal}
                onClick={(ev: React.SyntheticEvent<HTMLElement>) => props.onSiteDesignClicked(siteDesign)}>
                <DocumentCardPreview {...previewProps} />
                <DocumentCardDetails>
                    <DocumentCardTitle
                        title={siteDesign.Title}
                        shouldTruncate={true}
                    />
                </DocumentCardDetails>
            </DocumentCard>
        </div>;
    };

    let items = [...props.siteDesigns || []];
    if (props.preview) {
        items = items.slice(0, PREVIEW_ITEMS_COUNT);
    }
    if (!props.addNewDisabled) {
        items.push(null);
    }
    const seeMore = props.preview && props.siteDesigns && props.siteDesigns.length > PREVIEW_ITEMS_COUNT;
    return <div className={styles.SiteDesignsList}>
        <div className={styles.row}>
            <div className={styles.column}>
                <GridLayout
                    ariaLabel="List of Site Designss."
                    items={items}
                    onRenderGridItem={renderGridItem}
                />
                {seeMore && <div className={styles.seeMore}>
                    {`There are more than ${PREVIEW_ITEMS_COUNT} available Site Designs in your tenant. `}
                    <Link onClick={() => props.onSeeMore && props.onSeeMore()}>See all Site Designs</Link>
                </div>}
            </div>
        </div>
    </div>;
};