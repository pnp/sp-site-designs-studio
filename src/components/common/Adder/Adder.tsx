import * as React from "react";
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { IconButton } from "office-ui-fabric-react/lib/Button";
import { Callout, DirectionalHint } from "office-ui-fabric-react/lib/Callout";
import { Label } from "office-ui-fabric-react/lib/Label";
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import styles from "./Adder.module.scss";
import { useRef, useState, useEffect } from "react";
import { TextField, ITextField, Stack } from "office-ui-fabric-react";
import { getTrimmedText } from "../../../utils/textUtils";

export interface IAddableItem {
    iconName: string;
    text: string;
    key: string;
    group: string;
    item: any;
}

export interface IAddableItemsGroups {
    [group: string]: IAddableItem[];
}

export interface IAdderProps {
    items: IAddableItemsGroups;
    onSelectedItem: (item: IAddableItem) => void;
    noAvailableItemsText?: string;
    searchBoxPlaceholderText?: string;
}

const DEFAULT_NO_AVAILABLE_ITEMS_TEXT = "No available items...";


const iconClass = mergeStyles({
    width: "82px",
    height: "28px",
    minHeight: "28px",
    fontSize: "28px",
    lineHeight: "28px"
});

const ITEM_MAX_TEXT_LEN = 22;

const renderItem = (item: IAddableItem, index: number, onSelectedItem: (item: IAddableItem) => void) => (
    <div className={styles.item} key={`${item.key}_${index}`} onClick={() => onSelectedItem(item)} title={item.text} >
        <Icon iconName={item.iconName} className={iconClass} />
        <div className={styles.label}>
            <Label>{getTrimmedText(item.text, ITEM_MAX_TEXT_LEN)}</Label>
        </div>
    </div>);

const renderItemsList = (items: IAddableItemsGroups, filter: string, onSelectedItem: (item: IAddableItem) => void, noAvailableItemsText?: string) => <div className={styles.itemsList}>
    {Object.keys(items).map(group => {
        const groupItems = items[group];
        return <section key={`group_${group}_header`}>
            <header className={styles.groupHeader}>{group}</header>
            <div key={`group_${group}_items`} className={styles.groupContent}>
                {(!groupItems || groupItems.length == 0)
                    ? (noAvailableItemsText || DEFAULT_NO_AVAILABLE_ITEMS_TEXT)
                    : groupItems.filter(gi => gi.text.toLowerCase().indexOf(filter) >= 0).map((item, ndx) => renderItem(item, ndx, onSelectedItem))}
            </div>
        </section>;
    })}

</div>;

export const Adder = (props: IAdderProps) => {

    const [isSelecting, setIsSelecting] = useState<boolean>(false);
    const [searchCriteria, setSearchCriteria] = useState<string>('');
    const addButtonRef = useRef<HTMLDivElement>();
    const searchBoxRef = useRef<ITextField>();

    useEffect(() => {
        if (isSelecting && searchBoxRef && searchBoxRef.current) {
            searchBoxRef.current.focus();
        }
    });

    const onSelected = (item: IAddableItem) => {
        props.onSelectedItem(item);
        setIsSelecting(false);
        setSearchCriteria('');
    };

    const onSearchCriteriaChanged = (ev, criteria) => {
        setSearchCriteria(criteria);
    };


    return <div className={styles.Adder}>
        <button className={`${styles.button} ${isSelecting ? styles.isSelecting : ""}`} onClick={() => setIsSelecting(true)}>
            <div ref={addButtonRef} className={styles.plusIcon}>
                <Icon iconName="Add" />
            </div>
        </button>
        {isSelecting && <Callout
            role="alertdialog"
            gapSpace={0}
            target={addButtonRef.current}
            onDismiss={() => setIsSelecting(false)}
            setInitialFocus={true}
            directionalHint={DirectionalHint.bottomCenter}
        >
            <div className={styles.row}>
                <div className={styles.fullWidth}>
                    <Stack horizontal>
                        <div className={styles.iconSearch}>
                            <Icon iconName="Search" />
                        </div>
                        <TextField
                            key="ItemSearchBox"
                            borderless
                            componentRef={searchBoxRef}
                            placeholder={props.searchBoxPlaceholderText || "Search an item..."} onChange={onSearchCriteriaChanged} />
                    </Stack>
                </div>
            </div>
            {renderItemsList(props.items, searchCriteria && searchCriteria.toLowerCase(), onSelected, props.noAvailableItemsText)}
        </Callout>}
    </div>;
};