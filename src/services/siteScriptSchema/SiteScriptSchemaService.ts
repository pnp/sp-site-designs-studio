import { get, find } from '@microsoft/sp-lodash-subset';
import { ISiteScriptAction, ISiteScriptContent } from '../../models/ISiteScript';
import DefaultSchema from '../../schema/schema';
import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import { HttpClient, SPHttpClient } from '@microsoft/sp-http';
import { IPropertySchema } from '../../models/IPropertySchema';
const Ajv = require('ajv');

export interface IActionDescriptor {
    verb: string;
    label: string;
    description: string;
}

export interface IPropertyValuePair {
    property: string;
    value: any;
}

export interface ISiteScriptSchemaService {
    configure(schemaJSONorURL?: string, forceReconfigure?: boolean): Promise<any>;
    validateSiteScriptJson(json: string): boolean;
    getNewSiteScript(): ISiteScriptContent;
    getNewActionFromVerb(verb: string): ISiteScriptAction;
    getNewSubActionFromVerb(parentVerb: string, verb: string): ISiteScriptAction;
    getSiteScriptSchema(): any;
    getActionSchema(action: ISiteScriptAction): IPropertySchema;
    getActionSchemaByVerb(actionVerb: string): IPropertySchema;
    getActionTitle(action: ISiteScriptAction, parentAction?: ISiteScriptAction): string;
    getActionTitleByVerb(actionVerb: string, parentActionVerb?: string): string;
    getLabelFromActionSchema(actionSchema: IPropertySchema): string;
    getActionDescription(action: ISiteScriptAction, parentAction?: ISiteScriptAction): string;
    getActionDescriptionByVerb(actionVerb: string, parentActionVerb?: string): string;
    getDescriptionFromActionSchema(actionSchema: IPropertySchema): string;
    getSubActionSchemaByVerbs(parentActionVerb: string, subActionVerb: string): any;
    getSubActionSchema(parentAction: ISiteScriptAction, subAction: ISiteScriptAction): IPropertySchema;
    getAvailableActions(): IActionDescriptor[];
    getAvailableSubActions(parentAction: ISiteScriptAction): IActionDescriptor[];
    getPropertiesAndValues(action: ISiteScriptAction, parentAction?: ISiteScriptAction): IPropertyValuePair[];
    hasSubActions(action: ISiteScriptAction): boolean;
}

export class SiteScriptSchemaService implements ISiteScriptSchemaService {
    private schema: any = null;
    private isConfigured: boolean = false;
    private availableActions: IActionDescriptor[] = null;
    private availableSubActionByVerb: { [key: string]: IActionDescriptor[] } = null;
    private availableActionSchemas: { [key: string]: IPropertySchema } = null;
    private availableSubActionSchemasByVerb: { [key: string]: { [subActionKey: string]: IPropertySchema } } = null;
    private ajv = new (Ajv as any)({ schemaId: 'auto', extendRefs: true });

    constructor(private serviceScope: ServiceScope) { }

    private _getElementSchema(object: any, property: string = null): any {
        let value = !property ? object : object[property];
        if (value['$ref']) {
            let path = value['$ref'];
            return this._getPropertyFromPath(this.schema, path);
        }

        return value;
    }

    private _getPropertyFromPath(object: any, path: string, separator: string = '/'): any {
        path = path.replace('#/', '').replace('#', '').replace(new RegExp(separator, 'g'), '.');
        return get(object, path);
    }


    public getVerbFromActionSchema(actionSchema: IPropertySchema): string {
        if (
            !actionSchema.properties ||
            !(actionSchema.properties as any).verb ||
            !(actionSchema.properties as any).verb.enum ||
            !(actionSchema.properties as any).verb.enum.length
        ) {
            throw new Error('Invalid Action schema');
        }

        return (actionSchema.properties as any).verb.enum[0];
    }

    public getLabelFromActionSchema(actionSchema: IPropertySchema): string {
        let titleFromSchema = actionSchema.title;
        if (!titleFromSchema) {
            return this.getVerbFromActionSchema(actionSchema);
        }

        return titleFromSchema;
    }

    public getDescriptionFromActionSchema(actionSchema: IPropertySchema): string {
        let descriptionFromSchema = actionSchema.description;
        if (!descriptionFromSchema) {
            return '';
        }

        return descriptionFromSchema;
    }

    private _getSubActionsSchemaFromParentActionSchema(parentActionDefinition: any): IPropertySchema[] {
        if (!parentActionDefinition.properties) {
            throw new Error('Invalid Action schema');
        }

        if (!parentActionDefinition.properties.subactions) {
            return null;
        }

        if (
            parentActionDefinition.properties.subactions.type != 'array' ||
            !parentActionDefinition.properties.subactions.items ||
            !parentActionDefinition.properties.subactions.items.anyOf
        ) {
            throw new Error('Invalid Action schema');
        }

        return parentActionDefinition.properties.subactions.items.anyOf.map((subActionSchema) =>
            this._getElementSchema(subActionSchema)
        );
    }

    public configure(schemaJSONorURL?: string, forceReconfigure: boolean = false): Promise<void> {
        return new Promise((resolve, reject) => {
            if (this.isConfigured && !forceReconfigure) {
                resolve();
                return;
            }

            this._loadSchema(schemaJSONorURL)
                .then((schema) => {
                    if (!schema) {
                        reject('Schema cannot be found');
                        return;
                    }

                    this.schema = schema;
                    try {
                        // Get available action schemas
                        let actionsArraySchema = this.schema.properties.actions;

                        if (!actionsArraySchema.type || actionsArraySchema.type != 'array') {
                            throw new Error('Invalid Actions schema');
                        }

                        if (!actionsArraySchema.items || !actionsArraySchema.items.anyOf) {
                            throw new Error('Invalid Actions schema');
                        }

                        let actionsArraySchemaItems = actionsArraySchema.items;

                        // Get Main Actions schema
                        let availableActionSchemasAsArray: any[] = actionsArraySchemaItems.anyOf.map((action) =>
                            this._getElementSchema(action)
                        );
                        this.availableActionSchemas = {};
                        availableActionSchemasAsArray.forEach((actionSchema) => {
                            // Keep the current action schema
                            let actionVerb = this.getVerbFromActionSchema(actionSchema);
                            this.availableActionSchemas[actionVerb] = actionSchema;

                            // Check if the current action has subactions
                            let subActionSchemas = this._getSubActionsSchemaFromParentActionSchema(actionSchema);
                            if (subActionSchemas) {
                                // If yes, keep the sub actions schema and verbs

                                // Keep the list of subactions verbs
                                if (!this.availableSubActionByVerb) {
                                    this.availableSubActionByVerb = {};
                                }
                                this.availableSubActionByVerb[actionVerb] = subActionSchemas.map((sa) => ({
                                    verb: this.getVerbFromActionSchema(sa),
                                    label: this.getLabelFromActionSchema(sa),
                                    description: this.getDescriptionFromActionSchema(sa)
                                }));

                                // Keep the list of subactions schemas
                                if (!this.availableSubActionSchemasByVerb) {
                                    this.availableSubActionSchemasByVerb = {};
                                }
                                this.availableSubActionSchemasByVerb[actionVerb] = {};
                                subActionSchemas.forEach((sas) => {
                                    let subActionVerb = this.getVerbFromActionSchema(sas);
                                    this.availableSubActionSchemasByVerb[actionVerb][subActionVerb] = sas;
                                });
                            }
                        });
                        this.availableActions = availableActionSchemasAsArray.map((a) => ({
                            verb: this.getVerbFromActionSchema(a),
                            label: this.getLabelFromActionSchema(a),
                            description: this.getDescriptionFromActionSchema(a)
                        }));

                        this.isConfigured = true;
                        resolve();
                    } catch (error) {
                        console.error("Error:", error);
                        reject(error);
                    }
                })
                .catch((error) => reject(error));
        });
    }

    public validateSiteScriptJson(json: string): boolean {
        const parsedJson = JSON.parse(json);
        let valid = this.ajv.validate(this.schema, parsedJson);
        if (!valid) {
            console.error('Schema errors: ', this.ajv.errors);
            return false;
        }

        return true;
    }

    public getNewSiteScript(): ISiteScriptContent {
        return {
            $schema: 'schema.json',
            actions: [],
            bindata: {},
            version: 1
        };
    }

    private _getPropertyDefaultValueFromSchema(schema: any, propertyName: string): any {
        let propSchema = schema.properties[propertyName];
        if (propSchema) {
            switch (propSchema.type) {
                case 'string':
                    return '';
                case 'boolean':
                    return false;
                case 'number':
                    return 0;
                case 'object':
                    return {};
                case 'array':
                    return [];
                default:
                    if (propSchema.enum && propSchema.enum.length) {
                        return propSchema.enum[0];
                    } else {
                        return null;
                    }
            }
        } else {
            return null;
        }
    }

    public getNewActionFromVerb(verb: string): ISiteScriptAction {
        let newAction: ISiteScriptAction = {
            verb: verb
        };
        let actionSchema = this.getActionSchema(newAction);

        // Add default values for required properties of the action
        if (actionSchema && actionSchema.properties) {
            Object.keys(actionSchema.properties)
                .filter((p) => p != 'verb' && (p == 'subaction'
                    || !actionSchema.required
                    || actionSchema.required.indexOf(p) >= 0)
                    // Exclusive required properties are not preset
                    // || (actionSchema.anyOf && find(actionSchema.anyOf, s => s.required && s.required.indexOf(p) >= 0))
                    )
                .forEach((p) => (newAction[p] = this._getPropertyDefaultValueFromSchema(actionSchema, p)));
        }

        return newAction;
    }

    public getNewSubActionFromVerb(parentVerb: string, verb: string): ISiteScriptAction {
        let parentAction: ISiteScriptAction = {
            verb: parentVerb
        };
        let newAction: ISiteScriptAction = {
            verb: verb
        };
        let actionSchema = this.getSubActionSchema(parentAction, newAction);

        // Add default values for required properties of the action
        if (actionSchema && actionSchema.properties) {
            Object.keys(actionSchema.properties)
                .filter((p) => p != 'verb' && (!actionSchema.required || actionSchema.required.indexOf(p) >= 0))
                .forEach((p) => (newAction[p] = this._getPropertyDefaultValueFromSchema(actionSchema, p)));
        }

        return newAction;
    }

    public getSiteScriptSchema(): any {
        return this.schema;
    }

    public getActionSchema(action: ISiteScriptAction): IPropertySchema {
        return this.getActionSchemaByVerb(action.verb);
    }

    public getActionSchemaByVerb(actionVerb: string): IPropertySchema {
        if (!this.isConfigured) {
            throw new Error(
                'The Schema Service is not properly configured. Make sure the configure() method has been called.'
            );
        }

        let directResolvedSchema = this.availableActionSchemas[actionVerb];
        if (directResolvedSchema) {
            return directResolvedSchema;
        }

        // Try to find the schema by case insensitive key
        let availableActionKeys = Object.keys(this.availableActionSchemas);
        let foundKeys = availableActionKeys.filter((k) => k.toUpperCase() == actionVerb.toUpperCase());
        let actionSchemaKey = foundKeys.length == 1 ? foundKeys[0] : null;
        return this.availableActionSchemas[actionSchemaKey];
    }

    public getActionTitle(action: ISiteScriptAction, parentAction: ISiteScriptAction): string {
        return this.getActionTitleByVerb(action.verb, parentAction && parentAction.verb);
    }

    public getActionTitleByVerb(actionVerb: string, parentActionVerb: string): string {
        let actionSchema = parentActionVerb
            ? this.getSubActionSchemaByVerbs(parentActionVerb, actionVerb)
            : this.getActionSchemaByVerb(actionVerb);
        return actionSchema.title;
    }

    public getActionDescription(action: ISiteScriptAction, parentAction: ISiteScriptAction): string {
        return this.getActionDescriptionByVerb(action.verb, parentAction && parentAction.verb);
    }

    public getActionDescriptionByVerb(actionVerb: string, parentActionVerb: string): string {
        let actionSchema = parentActionVerb
            ? this.getSubActionSchemaByVerbs(parentActionVerb, actionVerb)
            : this.getActionSchemaByVerb(actionVerb);
        return actionSchema.description;
    }

    public getSubActionSchemaByVerbs(parentActionVerb: string, subActionVerb: string): IPropertySchema {
        if (!this.isConfigured) {
            throw new Error(
                'The Schema Service is not properly configured. Make sure the configure() method has been called.'
            );
        }

        let availableSubActionSchemas = this.availableSubActionSchemasByVerb[parentActionVerb];
        let directResolvedSchema = availableSubActionSchemas[subActionVerb];
        if (directResolvedSchema) {
            return directResolvedSchema;
        }

        // Try to find the schema by case insensitive key
        let availableSubActionKeys = Object.keys(availableSubActionSchemas);
        let foundKeys = availableSubActionKeys.filter((k) => k.toUpperCase() == subActionVerb.toUpperCase());
        let subActionSchemaKey = foundKeys.length == 1 ? foundKeys[0] : null;
        return availableSubActionSchemas[subActionSchemaKey];
    }

    public getSubActionSchema(parentAction: ISiteScriptAction, subAction: ISiteScriptAction): IPropertySchema {
        return this.getSubActionSchemaByVerbs(parentAction.verb, subAction.verb);
    }

    public getAvailableActions(): IActionDescriptor[] {
        if (!this.isConfigured) {
            throw new Error(
                'The Schema Service is not properly configured. Make sure the configure() method has been called.'
            );
        }

        return this.availableActions;
    }

    public getAvailableSubActions(parentAction: ISiteScriptAction): IActionDescriptor[] {
        if (!this.isConfigured) {
            throw new Error(
                'The Schema Service is not properly configured. Make sure the configure() method has been called.'
            );
        }

        return this.availableSubActionByVerb[parentAction.verb];
    }

    private _loadSchema(schemaJSONorURL: string): Promise<any> {
        return new Promise((resolve, reject) => {
            // If argument is not set, use the embedded default schema
            if (!schemaJSONorURL) {
                resolve(DefaultSchema);
                return;
            }

            if (
                schemaJSONorURL.indexOf('/') == 0 ||
                schemaJSONorURL.indexOf('http://') == 0 ||
                schemaJSONorURL.indexOf('https://') == 0
            ) {
                // The argument is a URL
                // Fetch the schema at the specified URL
                this._getSchemaFromUrl(schemaJSONorURL).then((schema) => resolve(schema)).catch((error) => {
                    console.error('An error occured while trying to fetch schema from URL', error);
                    reject(error);
                });
            } else {
                // The argument is supposed to be JSON stringified
                try {
                    let schema = JSON.parse(schemaJSONorURL);
                    resolve(schema);
                } catch (error) {
                    console.error('An error occured while parsing JSON string', error);
                    reject(error);
                }
            }
        });
    }

    private _getSchemaFromUrl(url: string): Promise<any> {
        // Use spHttpClient if it is a SPO URL, use regular httpClient otherwise
        if (url.indexOf('.sharepoint.com') > -1) {
            let spHttpClient: SPHttpClient = this.serviceScope.consume(SPHttpClient.serviceKey);
            return spHttpClient.get(url, SPHttpClient.configurations.v1).then((v) => v.json());
        } else {
            let httpClient: HttpClient = this.serviceScope.consume(HttpClient.serviceKey);
            return httpClient.get(url, HttpClient.configurations.v1).then((v) => v.json());
        }
    }

    public getPropertiesAndValues(action: ISiteScriptAction, parentAction?: ISiteScriptAction): IPropertyValuePair[] {
        if (!action) {
            return [];
        }

        const actionSchema = parentAction
            ? this.getSubActionSchema(parentAction, action)
            : this.getActionSchema(action);

        return Object.keys(actionSchema.properties).filter(p => p != "verb").map(p => (p == "subactions"
            ? {
                property: actionSchema.properties[p].title,
                value: `${(action.subactions && action.subactions.length) || 0} subactions`
            }
            : {
                property: actionSchema.properties[p].title,
                value: actionSchema.properties[p].type === "object" ? "Complex object" : action[p]
            }));
    }

    public hasSubActions(action: ISiteScriptAction): boolean {
        const actionSchema = this.getActionSchema(action);
        if (!actionSchema) {
            console.warn(`Action schema could not be resolved for verb ${action.verb}`);
            return false;
        }

        return Object.keys(actionSchema.properties).indexOf("subactions") >= 0;
    }
}

export const SiteScriptSchemaServiceKey = ServiceKey.create<ISiteScriptSchemaService>(
    'YPCODE:SDSv2:SiteScriptSchemaService',
    SiteScriptSchemaService
);
