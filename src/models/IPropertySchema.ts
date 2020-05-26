export interface IPropertySchema {
    type?: string;
    enum?: string[];
    title?: string;
    description?: string;
    properties?: { [property: string]: IPropertySchema };
    required?: string[];
    anyOf?: IPropertySchema[];
    items?: IPropertySchema;
}

export interface IObjectSchema extends IPropertySchema { }