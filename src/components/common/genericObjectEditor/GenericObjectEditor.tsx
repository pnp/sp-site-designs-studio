import * as React from 'react';
import { useEffect, useState } from 'react';
import { Dropdown, TextField, Toggle, IconButton, Stack, Label, IDropdownOption } from 'office-ui-fabric-react';
import { IPropertySchema, IObjectSchema } from '../../../models/IPropertySchema';
import { useConstCallback } from '@uifabric/react-hooks';

export interface IPropertyEditorProps {
    schema: IPropertySchema;
    value: any;
    label?: string;
    required?: boolean;
    readonly?: boolean;
    onChange: (value: any) => void;
}

const getDefaulValueFromSchema = (schema: IPropertySchema) => {
    if (schema) {
        switch (schema.type) {
            case 'string':
                return '';
            case 'boolean':
                return false;
            case 'number':
                return 0;
            case 'object':
                return {};
            default:
                return null;
        }
    } else {
        return null;
    }
};

export function PropertyEditor(props: IPropertyEditorProps) {
    let { schema,
        label,
        readonly,
        required,
        value,
        onChange } = props;

    const onDropdownChange = ((ev: any, v: IDropdownOption) => {
        onChange(v.key);
    });

    const onNumberInputChange = ((ev: any, v: any) => {
        if (typeof (v) === "number") {
            onChange(v);
        } else {
            const number = parseFloat(v as string);
            onChange(number);
        }
    });

    const onInputChange = ((ev: any, v: any) => {
        onChange(v);
    });

    if (schema.enum) {
        if (schema.enum.length > 1 && !readonly) {
            return (
                <Dropdown
                    required={required}
                    label={label}
                    selectedKey={value}
                    options={schema.enum.map((p) => ({ key: p, text: p }))}
                    onChange={onDropdownChange}
                />
            );
        } else {
            return (
                <TextField
                    label={label}
                    value={value}
                    readOnly={true}
                    required={required}
                    onChange={onInputChange}
                />
            );
        }
    } else {
        switch (schema.type) {
            case 'boolean':
                return (
                    <Toggle
                        label={label}
                        checked={value as boolean}
                        disabled={readonly}
                        onChange={onInputChange}
                    />
                );
            case 'array':
                return <>
                    <Label>{label}</Label>
                    <GenericArrayEditor
                        object={value}
                        schema={schema.items}
                        onObjectChanged={onChange} />
                </>;
            case 'object': // TODO If object is a simple dictionary (key/non-complex object values) => Display a custom control
                return <GenericObjectEditor
                    object={value}
                    schema={schema}
                    onObjectChanged={onChange}
                />;
            case 'number':
                return (
                    <TextField
                        required={required}
                        label={label}
                        value={value}
                        readOnly={readonly}
                        onChange={onNumberInputChange}
                    />
                );
            case 'string':
            default:
                return (
                    <TextField
                        required={required}
                        label={label}
                        value={value}
                        readOnly={readonly}
                        onChange={onInputChange}
                    />
                );
        }
    }
}

export interface IGenericObjectEditorProps {
    schema: IObjectSchema;
    object: any;
    defaultValues?: any;
    customRenderers?: any;
    ignoredProperties?: string[];
    readOnlyProperties?: string[];
    updateOnBlur?: boolean;
    fieldLabelGetter?: (field: string) => string;
    onObjectChanged?: (object: any) => void;
    children?: any;
}

export interface IPropertyPlaceholderProps { propertyName: string; }
export const PropertyPlaceholder = (props: IPropertyPlaceholderProps) => <></>;

export function GenericArrayEditor(arrayEditorProps: IGenericObjectEditorProps) {

    const onRemoved = ((index) => {
        arrayEditorProps.onObjectChanged(arrayEditorProps.object.filter((_, i) => i != index));
    });

    const onAdded = (() => {
        arrayEditorProps.onObjectChanged([...(arrayEditorProps.object||[]), getDefaulValueFromSchema(arrayEditorProps.schema)]);
    });

    const onUpdated = ((index, newValue) => {
        arrayEditorProps.onObjectChanged(arrayEditorProps.object.map((o, i) => i == index ? newValue : o));
    });

    const renderItems = () => {
        if (!arrayEditorProps.object) {
            return null;
        }

        const items = arrayEditorProps.object as any[];
        return items.map((item, index) => <>
            <Stack horizontal>
                <PropertyEditor key={index} value={item} schema={arrayEditorProps.schema} onChange={o => onUpdated(index, o)} />
                <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => onRemoved(index)} />
            </Stack>
        </>);
    };

    return <>
        {renderItems()}
        <IconButton iconProps={{ iconName: 'Add' }} onClick={onAdded} />
    </>;
}

export function GenericObjectEditor(props: IGenericObjectEditorProps) {

    const [objectProperties, setObjectProperties] = useState<string[]>([]);

    const getPropertyDefaultValueFromSchema = (propertyName: string, componentProps: IGenericObjectEditorProps) => {
        let propSchema = componentProps.schema.properties[propertyName];
        return getDefaulValueFromSchema(propSchema);
    };

    const getPropertyTypeFromSchema = (propertyName: string, componentProps: IGenericObjectEditorProps) => {
        let propSchema = componentProps.schema.properties[propertyName];
        if (propSchema) {
            return propSchema.type;
        } else {
            return null;
        }
    };

    const isPropertyReadOnly = useConstCallback((propertyName: string, componentProps: IGenericObjectEditorProps) => {
        if (!componentProps.readOnlyProperties || !componentProps.readOnlyProperties.length) {
            return false;
        }

        return componentProps.readOnlyProperties.indexOf(propertyName) > -1;
    });

    const onObjectPropertyChange = useConstCallback((propertyName: string, newValue: any, componentProps: IGenericObjectEditorProps) => {
        if (!componentProps.onObjectChanged) {
            return;
        }

        let propertyType = getPropertyTypeFromSchema(propertyName, componentProps);
        if (propertyType == 'number') {
            newValue = Number(newValue);
        }
        const updatedObject = { ...componentProps.object, [propertyName]: newValue };

        // Set default values for properties of the argument object if not set
        objectProperties.forEach((p) => {
            // Get the property type

            let defaultValue =
                componentProps.defaultValues && componentProps.defaultValues[p]
                    ? componentProps.defaultValues[p]
                    : getPropertyDefaultValueFromSchema(p, componentProps);

            if (!updatedObject[p] && updatedObject[p] != false && updatedObject[p] != 0) {
                updatedObject[p] = defaultValue;
            }
        });

        componentProps.onObjectChanged(updatedObject);
    });

    const getFieldLabel = useConstCallback((field: string, propertyDefinition: IPropertySchema, componentProps: IGenericObjectEditorProps) => {
        if (componentProps.fieldLabelGetter) {
            const foundLabel = componentProps.fieldLabelGetter(field);
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
    });

    const renderPropertyEditor = useConstCallback((propertyName: string, propertySchema: IPropertySchema, componentProps: IGenericObjectEditorProps) => {
        let { schema, customRenderers, defaultValues, object } = componentProps;

        // Has custom renderer for the property
        if (customRenderers && customRenderers[propertyName]) {
            // If a default value is specified for current property and it is null, apply it
            if (!object[propertyName] && defaultValues && defaultValues[propertyName]) {
                object[propertyName] = defaultValues[propertyName];
            }

            return customRenderers[propertyName](object[propertyName]);
        }

        let isPropertyRequired =
            (schema.required && schema.required.length && schema.required.indexOf(propertyName) > -1) || false;


        return <PropertyEditor
            value={object[propertyName]}
            required={isPropertyRequired}
            onChange={v => onObjectPropertyChange(propertyName, v, componentProps)}
            label={getFieldLabel(propertyName, propertySchema, componentProps)}
            schema={propertySchema}
            readonly={isPropertyReadOnly(propertyName, componentProps)}
        />;

        // if (propertySchema.enum) {
        //     if (propertySchema.enum.length > 1 || !isPropertyReadOnly(propertyName)) {
        //         return (
        //             <Dropdown
        //                 required={isPropertyRequired}
        //                 label={getFieldLabel(propertyName, propertySchema)}
        //                 selectedKey={object[propertyName]}
        //                 options={propertySchema.enum.map((p) => ({ key: p, text: p }))}
        //                 onChange={(_, value) => onObjectPropertyChange(propertyName, value.key)}
        //             />
        //         );
        //     } else {
        //         return (
        //             <TextField
        //                 label={getFieldLabel(propertyName, propertySchema)}
        //                 value={object[propertyName]}
        //                 readOnly={true}
        //                 required={isPropertyRequired}
        //                 onChange={(_, value) => onObjectPropertyChange(propertyName, value)}
        //             />
        //         );
        //     }
        // } else {
        //     switch (propertySchema.type) {
        //         case 'boolean':
        //             return (
        //                 <Toggle
        //                     label={getFieldLabel(propertyName, propertySchema)}
        //                     checked={object[propertyName] as boolean}
        //                     disabled={isPropertyReadOnly(propertyName)}
        //                     onChange={(_, value) => onObjectPropertyChange(propertyName, value)}
        //                 />
        //             );
        //         case 'array':
        //             return <GenericArrayEditor
        //                 object={object[propertyName]}
        //                 schema={propertySchema.items}
        //                 onObjectChanged={o => onObjectPropertyChange(propertyName, o)} />;
        //         // case 'object': // TODO If object is a simple dictionary (key/non-complex object values) => Display a custom control
        //         // 	return (
        //         // 		<div>
        //         // 			<div className="ms-Grid-row">
        //         // 				<div className="ms-Grid-col ms-sm12">
        //         // 					<Label>{this._getFieldLabel(propertyName)}</Label>
        //         // 				</div>
        //         // 			</div>
        //         // 			<div className="ms-Grid-row">
        //         // 				<div className="ms-Grid-col ms-sm2">
        //         // 					<Icon iconName="InfoSolid" />
        //         // 				</div>
        //         // 				<div className="ms-Grid-col ms-sm10">
        //         // 					{strings.PropertyIsComplexTypeMessage}
        //         // 					<br />
        //         // 					{strings.UseJsonEditorMessage}
        //         // 				</div>
        //         // 			</div>
        //         // 		</div>
        //         // 	);
        //         case 'number':
        //         case 'string':
        //         default:
        //             return (
        //                 <TextField
        //                     required={isPropertyRequired}
        //                     label={getFieldLabel(propertyName, propertySchema)}
        //                     value={object[propertyName]}
        //                     readOnly={isPropertyReadOnly(propertyName)}
        //                     onChange={(ev, value) => onObjectPropertyChange(propertyName, value)}
        //                 />
        //             );
        //     }
        // }
    });

    const refreshObjectProperties = useConstCallback((componentProps: IGenericObjectEditorProps) => {
        let { schema, ignoredProperties } = componentProps;

        if (schema.type != 'object') {
            throw new Error('Cannot generate Object Editor from a non-object type');
        }

        if (!schema.properties || Object.keys(schema.properties).length == 0) {
            return;
        }

        let refreshedObjectProperties = Object.keys(schema.properties);
        if (ignoredProperties && ignoredProperties.length > 0) {
            refreshedObjectProperties = refreshedObjectProperties.filter((p) => ignoredProperties.indexOf(p) < 0);
        }
        setObjectProperties(refreshedObjectProperties);
    });

    // Use effects
    useEffect(() => {
        refreshObjectProperties(props);
    }, [props.schema]);

    // TODO See if really needed
    // private editTextValues: any;
    // private _onTextFieldValueChanged(fieldName: string, value: any) {
    //     if (this.props.updateOnBlur) {
    //         if (!this.editTextValues) {
    //             this.editTextValues = {};
    //         }
    //         this.editTextValues[fieldName] = value;
    //     } else {
    //         this._onObjectPropertyChange(fieldName, value);
    //     }
    // }

    // private _onTextFieldEdited(fieldName: string) {
    //     let value = this.editTextValues && this.editTextValues[fieldName];
    //     this._onObjectPropertyChange(fieldName, value);
    //     if (value) {
    //         delete this.editTextValues[fieldName];
    //     }
    // }

    const renderChildrenRecursive = useConstCallback((node: any, editorProps: IGenericObjectEditorProps) => {
        return React.Children.map(node.props.children, (c, i) => {
            const asPlaceholder = c as React.ReactElement<IPropertyPlaceholderProps>;
            if (asPlaceholder.type == PropertyPlaceholder) {
                const propName = asPlaceholder.props.propertyName;
                return renderPropertyEditor(propName, editorProps.schema.properties[propName], props);
            } else {
                if (React.Children.count(c.props.children) == 0) {
                    return c;
                } else {
                    return React.cloneElement(c, { children: renderChildrenRecursive(c) });
                }
            }
        });
    });

    const render = () => {
        let { schema, ignoredProperties, object } = props;
        if (!object) {
            return null;
        }

        let propertyEditors = {};
        objectProperties.forEach(p => {
            if (ignoredProperties && ignoredProperties.indexOf(p) >= 0) {
                return;
            }

            propertyEditors[p] = renderPropertyEditor(p, schema.properties[p], props);
        });

        if (React.Children.count(props.children) > 0) {
            return <>
                {renderChildrenRecursive(render(), props)}
            </>;
        } else {
            return <>
                {Object.keys(propertyEditors).map(k => propertyEditors[k])}
            </>;
        }
    };

    return render();
}