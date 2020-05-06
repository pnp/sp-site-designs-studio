import { ISiteScriptContent, ISiteScriptAction, ISiteScript } from "../models/ISiteScript";
import { find, clone } from "@microsoft/sp-lodash-subset";

export interface ISiteScriptActionUIWrapper {
    $uiKey: string;
    verb: string;
    subactions?: ISiteScriptActionUIWrapper[];
}

export interface IActionsReplacer {
    [key: string]: () => ISiteScriptActionUIWrapper;
}

export interface ISiteScriptContentUIWrapper {
    $schema?: string;
    actions: ISiteScriptActionUIWrapper[];
    bindata: {};
    version: number;
    editingActionKeys: string[];
    toSiteScriptContent(actionReplacers?: IActionsReplacer): ISiteScriptContent;
    toJSON(): string;
    isEqualToRawJSON(json: string): boolean;
    addAction(action: ISiteScriptAction): ISiteScriptContentUIWrapper;
    removeAction(action: ISiteScriptActionUIWrapper): ISiteScriptContentUIWrapper;
    addSubAction(parentAction: ISiteScriptActionUIWrapper, action: ISiteScriptAction): ISiteScriptContentUIWrapper;
    removeSubAction(parentAction: ISiteScriptActionUIWrapper, action: ISiteScriptActionUIWrapper): ISiteScriptContentUIWrapper;
    toggleEditing(action: ISiteScriptActionUIWrapper, parentAction?: ISiteScriptActionUIWrapper): ISiteScriptContentUIWrapper;
    clearEditing(exceptedKeys?: string[]): ISiteScriptContentUIWrapper;
    replaceAction(action: ISiteScriptActionUIWrapper): ISiteScriptContentUIWrapper;
    reorderActions(newIndex: number, oldIndex: number): ISiteScriptContentUIWrapper;
    reorderSubActions(parentActionKey: string, newIndex: number, oldIndex: number): ISiteScriptContentUIWrapper;
}

interface IKeyCounters {
    [key: string]: number;
}

export class SiteScriptContentUIWrapper implements ISiteScriptContentUIWrapper {
    public $schema: string;
    public actions: ISiteScriptActionUIWrapper[];
    public bindata: any;
    public version: number;
    public editingActionKeys: string[] = [];
    private keyCounters: IKeyCounters = {};


    constructor(siteScriptContent: ISiteScriptContent) {
        if (!siteScriptContent) {
            return;
        }

        this.$schema = siteScriptContent.$schema;
        this.bindata = siteScriptContent.bindata;
        this.version = siteScriptContent.version;
        this._initializeActionUIWrappers(siteScriptContent);
    }

    public clone(clonedProperties?: ISiteScriptContentUIWrapper) {
        const cloned = new SiteScriptContentUIWrapper(null);
        cloned.$schema = (clonedProperties || this).$schema;
        cloned.bindata = (clonedProperties || this).bindata;
        cloned.version = (clonedProperties || this).version;
        cloned.editingActionKeys = (clonedProperties || this).editingActionKeys.map(k => k);
        // For code readability, ensure subactions is the last object property
        cloned.actions = (clonedProperties || this).actions.map(a => ({ ...a, subactions: a.subactions || undefined }));
        return cloned;
    }

    private _getActionUIWrapper(action: ISiteScriptAction, parentActionKey?: string, extraProperties?: any): ISiteScriptActionUIWrapper {
        const $uiKey = this._getActionKey(action, parentActionKey);
        let a = ({
            ...action,
            $uiKey,
            subactions: this._getSubActions(action, $uiKey)
        });
        if (extraProperties) {
            a = { ...a, ...extraProperties };
        }
        return a;
    }

    private _getActionKey(action: ISiteScriptAction, parentActionKey?: string): string {
        const defaultKey = `${parentActionKey ? `${parentActionKey}_` : ''}${action.verb}`;

        if (this.keyCounters[defaultKey] || this.keyCounters[defaultKey] == 0) {
            this.keyCounters[defaultKey]++;
        } else {
            this.keyCounters[defaultKey] = 0;
        }
        return `${defaultKey}_${this.keyCounters[defaultKey]}`;
    }

    private _getSubActions(action: ISiteScriptAction, parentActionKey: string): ISiteScriptActionUIWrapper[] {
        if (typeof action.subactions === "undefined") {
            return undefined;
        }

        return (action.subactions || []).map(subaction => this._getActionUIWrapper(subaction, parentActionKey));
    }

    private _initializeActionUIWrappers(siteScriptContent: ISiteScriptContent) {
        // Deep clone the site script content
        this.actions = (siteScriptContent.actions || []).map(a => (this._getActionUIWrapper(a)));
    }

    public addAction(action: ISiteScriptAction): ISiteScriptContentUIWrapper {
        const cloned = this.clone();
        const newKey = this._getActionKey(action);
        cloned.editingActionKeys.push(newKey);
        cloned.actions.push({ $uiKey: newKey, ...action } as ISiteScriptActionUIWrapper);
        return cloned;
    }

    public removeAction(action: ISiteScriptActionUIWrapper): ISiteScriptContentUIWrapper {
        const cloned = this.clone();
        cloned.editingActionKeys = this.editingActionKeys.filter(k => k != action.$uiKey);
        cloned.actions = cloned.actions.filter(a => a.$uiKey != action.$uiKey) as any[];
        return cloned;
    }

    public addSubAction(parentAction: ISiteScriptActionUIWrapper, action: ISiteScriptAction): ISiteScriptContentUIWrapper {
        const cloned = this.clone();
        const foundParentAction = find(cloned.actions, a => a.$uiKey == parentAction.$uiKey);
        if (!foundParentAction.subactions) {
            foundParentAction.subactions = [];
        }
        const newKey = this._getActionKey(action, parentAction.$uiKey);
        cloned.editingActionKeys.push(newKey);
        foundParentAction.subactions.push({ $uiKey: newKey, ...action } as ISiteScriptActionUIWrapper);
        return cloned;
    }

    public removeSubAction(parentAction: ISiteScriptActionUIWrapper, action: ISiteScriptActionUIWrapper): ISiteScriptContentUIWrapper {
        const cloned = this.clone();
        const foundParentAction = find(cloned.actions, a => a.$uiKey == parentAction.$uiKey);
        cloned.editingActionKeys = this.editingActionKeys.filter(k => k != action.$uiKey);
        foundParentAction.subactions = foundParentAction.subactions.filter(a => a.$uiKey != action.$uiKey);
        return cloned;
    }

    public toggleEditing(action: ISiteScriptActionUIWrapper): ISiteScriptContentUIWrapper {
        const cloned = this.clone();

        if (cloned.editingActionKeys.indexOf(action.$uiKey) >= 0) {
            // Remove the current action from edition
            cloned.editingActionKeys = cloned.editingActionKeys.filter(k => k != action.$uiKey);
        } else {
            // Add the current action to edition
            cloned.editingActionKeys.push(action.$uiKey);
        }
        return cloned;
    }

    public clearEditing(exceptedKeys: string[]): ISiteScriptContentUIWrapper {
        const cloned = this.clone();
        cloned.editingActionKeys = exceptedKeys ? cloned.editingActionKeys.filter(k => exceptedKeys.indexOf(k) >= 0) : [];
        return cloned;
    }

    public replaceAction(action: ISiteScriptActionUIWrapper): ISiteScriptContentUIWrapper {
        const cloned = this.clone();
        cloned.actions = (this.actions || []).map(a => ({
            ...(a.$uiKey == action.$uiKey ? action : a),
            subactions: a.subactions && a.subactions.map(sa => ({ ...(sa.$uiKey == action.$uiKey ? action : sa) }))
        }));
        return cloned;
    }

    public reorderActions(newIndex: number, oldIndex: number): ISiteScriptContentUIWrapper {
        const cloned = this.clone();
        const actionToMove = cloned.actions[oldIndex];
        cloned.actions.splice(oldIndex, 1);
        cloned.actions.splice(newIndex, 0, actionToMove);
        return cloned;
    }

    public reorderSubActions(parentActionKey: string, newIndex: number, oldIndex: number): ISiteScriptContentUIWrapper {
        const cloned = this.clone();
        const parentAction = find(cloned.actions, a => a.$uiKey == parentActionKey);
        if (!parentAction || !parentAction.subactions) {
            console.warn(`Parent action could not be found with key ${parentActionKey}`);
            return;
        }

        const actionToMove = parentAction.subactions[oldIndex];
        parentAction.subactions.splice(oldIndex, 1);
        parentAction.subactions.splice(newIndex, 0, actionToMove);
        return cloned;
    }

    public toSiteScriptContent(actionReplacers?: IActionsReplacer): ISiteScriptContent {
        const rawContent: ISiteScriptContent = {
            $schema: this.$schema,
            bindata: this.bindata,
            version: this.version,
            actions: (this.actions || []).map(a => {
                let effectiveAction = actionReplacers && actionReplacers[a.$uiKey] ? actionReplacers[a.$uiKey]() : a;
                const clonedAction = { ...effectiveAction };
                delete clonedAction.$uiKey;
                if (clonedAction.subactions && clonedAction.subactions.length > 0) {
                    clonedAction.subactions = clonedAction.subactions.map(sa => {
                        let effectiveSubAction = actionReplacers && actionReplacers[sa.$uiKey] ? actionReplacers[sa.$uiKey]() : sa;
                        const clonedSubAction = { ...effectiveSubAction };
                        delete clonedSubAction.$uiKey;
                        return clonedSubAction as ISiteScriptActionUIWrapper;
                    });
                }
                return clonedAction as ISiteScriptAction;
            })
        };

        return rawContent;
    }

    public toJSON(): string {
        const scriptContent = this.toSiteScriptContent();
        return JSON.stringify(scriptContent, null, 4);
    }

    public isEqualToRawJSON(json: string): boolean {
        const currentJson = this.toJSON();
        return currentJson == json;
    }

}