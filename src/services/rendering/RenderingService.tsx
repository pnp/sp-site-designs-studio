import * as React from "react";
import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import { IPropertySchema } from "../../models/IPropertySchema";
import { ISiteScriptSchemaService, SiteScriptSchemaServiceKey } from "../siteScriptSchema/SiteScriptSchemaService";
import { ISiteScriptAction } from '../../models/ISiteScript';
import { PropertyEditor } from "../../components/common/genericObjectEditor/GenericObjectEditor";
import { Label } from "office-ui-fabric-react";

export interface IRenderer {
    (label: string, value: any, onValueChange: (changedValue: any) => void): JSX.Element;
}

export interface ILabelResolver {
    (propertyName: string): string;
}

export interface IRenderingService {
    customizeActionPropertyRendering(mainActionVerb: string, subActionVerb: string, propertyName: string, renderer: IRenderer, labelResolver?: ILabelResolver): void;
    renderActionProperty(siteScriptAction: ISiteScriptAction, parentSiteScriptAction: ISiteScriptAction, propertyName: string, value: any, onChanged: (value: any) => void, ignoredProperties: string[]): JSX.Element;
    renderActionProperties(siteScriptAction: ISiteScriptAction, parentSiteScriptAction: ISiteScriptAction, onActionChange: (scriptAction: ISiteScriptAction) => void, ignoredProperties: string[]): JSX.Element;
}

interface IActionPropertyRenderingOptions {
    [mainActionKey: string]: {
        renderer: IRenderer;
        propertyName: string;
        labelResolver: ILabelResolver;
        subactions: {
            [subActionKey: string]: {
                labelResolver: ILabelResolver;
                propertyName: string;
                renderer: IRenderer;
            }
        }
    };
}

interface IActionPropertySchemas {
    [mainActionKey: string]: {
        schema: IPropertySchema;
        subactions: {
            [subActionKey: string]: {
                schema: IPropertySchema;
            }
        }
    };
}


class RenderingService implements IRenderingService {

    private _registeredActionPropertyRenderingOptions: IActionPropertyRenderingOptions = {};
    private _cacheActionSchemas: IActionPropertySchemas = {};
    private siteScriptSchemaService: ISiteScriptSchemaService;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            this.siteScriptSchemaService = serviceScope.consume(SiteScriptSchemaServiceKey);
        });
    }

    private _ensureActionPropertySchema(verb: string): IPropertySchema {
        if (!this._cacheActionSchemas[verb]) {
            // Cache the action schema
            this._cacheActionSchemas[verb] = {
                schema: this.siteScriptSchemaService.getActionSchemaByVerb(verb),
                subactions: {}
            };
        }
        return this._cacheActionSchemas[verb].schema;
    }

    private _ensureSubActionPropertySchema(parentVerb: string, verb: string): IPropertySchema {
        this._ensureActionPropertySchema(parentVerb);
        if (!this._cacheActionSchemas[parentVerb][verb]) {
            this._cacheActionSchemas[parentVerb][verb] = {
                schema: this.siteScriptSchemaService.getSubActionSchemaByVerbs(parentVerb, verb)
            };
        }
        return this._cacheActionSchemas[parentVerb][verb].schema;
    }

    public customizeActionPropertyRendering(mainActionVerb: string, subActionVerb: string, propertyName: string, renderer: IRenderer, labelResolver?: ILabelResolver): void {
        if (!mainActionVerb) {
            throw new Error("Main action verb is not defined");
        }

        if (!this._registeredActionPropertyRenderingOptions[mainActionVerb]) {
            this._registeredActionPropertyRenderingOptions[mainActionVerb] = { renderer: null, propertyName: null, subactions: null, labelResolver: null };
        }
        this._ensureActionPropertySchema(mainActionVerb);
        if (subActionVerb) {
            this._ensureSubActionPropertySchema(mainActionVerb, subActionVerb);
            this._registeredActionPropertyRenderingOptions[mainActionVerb].subactions[subActionVerb] = { propertyName, renderer, labelResolver };
        } else {
            this._registeredActionPropertyRenderingOptions[mainActionVerb].labelResolver = labelResolver;
            this._registeredActionPropertyRenderingOptions[mainActionVerb].renderer = renderer;
            this._registeredActionPropertyRenderingOptions[mainActionVerb].propertyName = propertyName;
        }
    }

    private _tryGetActionPropertyLabelResolver(mainActionVerb: string, subActionVerb: string): ILabelResolver {
        if (this._registeredActionPropertyRenderingOptions[mainActionVerb]) {
            if (subActionVerb) {
                if (this._registeredActionPropertyRenderingOptions[mainActionVerb][subActionVerb]) {
                    return this._registeredActionPropertyRenderingOptions[mainActionVerb][subActionVerb].labelResolver;
                }
            } else {
                return this._registeredActionPropertyRenderingOptions[mainActionVerb].labelResolver;
            }
        }
        return null;
    }
    private _getActionPropertyLabel(mainActionVerb: string, subActionVerb: string, propertyName: string, propertyDefinition: IPropertySchema): string {
        const labelResolver = this._tryGetActionPropertyLabelResolver(mainActionVerb, subActionVerb);
        if (labelResolver) {
            const foundLabel = labelResolver(propertyName);
            if (foundLabel) {
                return foundLabel;
            }
        }

        // TODO Handle this from specified field label getter
        // Try translate from built-in resources
        // let key = 'PROP_' + field;
        // if (strings[key]) {
        //     return strings[key];
        // } else 
        if (propertyDefinition && propertyDefinition.title) {
            return propertyDefinition.title;
        } else {
            return '';
        }
    }

    public renderActionProperty(siteScriptAction: ISiteScriptAction, parentSiteScriptAction: ISiteScriptAction, propertyName: string, value: any, onChanged: (value: any) => void, ignoredProperties: string[]): JSX.Element {
        if (!siteScriptAction) {
            throw new Error("action verb is not defined");
        }
        const mainActionVerb = parentSiteScriptAction ? parentSiteScriptAction.verb : siteScriptAction.verb;
        const subActionVerb = parentSiteScriptAction ? siteScriptAction.verb : null;

        if (propertyName == "subactions") {
            if (subActionVerb) {
                throw new Error("Subactions cannot have subactions");
            }

            const subactions: ISiteScriptAction[] = siteScriptAction.subactions;
            return <>
                <Label>Subactions</Label>
                {subactions.map((sa, index) => this.renderActionProperties(sa, siteScriptAction, (changedSubAction) => {
                    const updatedSubActions = (subactions).map((usa, updatedIndex) => updatedIndex == index ? changedSubAction : usa);
                    onChanged(updatedSubActions);
                }, ignoredProperties))}
            </>;
        }

        let actionSchema: IPropertySchema = null;

        if (subActionVerb) {
            actionSchema = this._ensureSubActionPropertySchema(mainActionVerb, subActionVerb);
        } else {
            actionSchema = this._ensureActionPropertySchema(mainActionVerb);
        }
        if (!actionSchema) {
            console.warn("Action schema could not be resolved");
            return null;
        }

        let isPropertyRequired = (actionSchema.required
            && actionSchema.required.length
            && actionSchema.required.indexOf(propertyName) > -1) || false;

        const propertySchema = actionSchema.properties[propertyName];
        const propertyLabel = this._getActionPropertyLabel(mainActionVerb, subActionVerb, propertyName, propertySchema);
        let content: JSX.Element = null;
        // Get the rendering config for main action (if any)
        const mainActionRenderingConfig = this._registeredActionPropertyRenderingOptions[mainActionVerb];
        if (subActionVerb) {
            // If the current action is a subaction
            // Get the rendering config for sub action (if any)
            const subActionRenderingConfig = mainActionRenderingConfig
                && mainActionRenderingConfig.subactions
                && mainActionRenderingConfig.subactions[subActionVerb];
            // If there is no rendering config for current action and property
            if (!subActionRenderingConfig || subActionRenderingConfig.propertyName != propertyName) {
                // Use default property editor
                content = <PropertyEditor
                    value={siteScriptAction[propertyName]}
                    required={isPropertyRequired}
                    onChange={onChanged}
                    label={propertyLabel}
                    schema={propertySchema}
                />;
            } else {
                // Otherwise, use the specified renderer
                content = subActionRenderingConfig.renderer(propertyLabel, value, onChanged);
            }
        } else {
            // If the current action is a main action
            // If there is no rendering config for current action and property
            if (!mainActionRenderingConfig || mainActionRenderingConfig.propertyName != propertyName) {
                // Use default property editor
                content = <PropertyEditor
                    value={siteScriptAction[propertyName]}
                    required={isPropertyRequired}
                    onChange={onChanged}
                    label={propertyLabel}
                    schema={propertySchema}
                />;
            } else {
                // Otherwise, use the specified renderer
                content = mainActionRenderingConfig.renderer(propertyLabel, value, onChanged);
            }
        }
        return content;
    }

    public renderActionProperties(siteScriptAction: ISiteScriptAction, parentSiteScriptAction: ISiteScriptAction, onActionChange: (scriptAction: ISiteScriptAction) => void, ignoredProperties: string[]): JSX.Element {
        ignoredProperties = ignoredProperties || ['verb', 'subactions'];

        const actionSchema = parentSiteScriptAction
            ? this._ensureSubActionPropertySchema(parentSiteScriptAction.verb, siteScriptAction.verb)
            : this._ensureActionPropertySchema(siteScriptAction.verb);

        return <>
            {Object.keys(actionSchema.properties).filter(p => ignoredProperties.indexOf(p) < 0).map(p => this.renderActionProperty(siteScriptAction, parentSiteScriptAction, p, siteScriptAction[p], v => {
                const updated = { ...siteScriptAction, [p]: v };
                onActionChange(updated);
            }, ignoredProperties))}
        </>;
    }
}

export const RenderingServiceKey = ServiceKey.create<IRenderingService>('YPCODE:SDSv2:RenderingService', RenderingService);
