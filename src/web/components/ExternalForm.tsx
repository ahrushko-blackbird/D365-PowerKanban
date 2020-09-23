import * as React from "react";
import { useAppContext } from "../domain/AppState";

import { fetchData, refresh } from "../domain/fetchData";
import { UserInputModal } from "./UserInputModalProps";
import { useActionContext } from "../domain/ActionState";
import { FlyOutField, FlyOutLookupField } from "../domain/FlyOutForm";
import { TextField } from "@fluentui/react/lib/TextField";
import { TagPicker, ITag, IBasePicker, IInputProps, IBasePickerSuggestionsProps } from "@fluentui/react/lib/Pickers";
import { Label } from "@fluentui/react/lib/Label";
import { Text } from "@fluentui/react/lib/Text";
import * as WebApiClient from "xrm-webapi-client";

interface ExternalFormProps {
}

export const ExternalForm = (props: ExternalFormProps) => {
    const [ actionState, actionDispatch ] = useActionContext();
    const [ formData, setFormData ] = React.useState({} as any);
    const [ pickData, setPickData ] = React.useState({} as { [key: string]: Array<{key: string, name: string}> });

    const fields: Array<[string, FlyOutField]> = Object.keys(actionState.flyOutForm.fields)
        .map(fieldId => [ fieldId, actionState.flyOutForm.fields[fieldId]]);

    const noCallBack = () => {
        actionState.flyOutForm.resolve({
            cancelled: true
        });
    };

    const yesCallBack = () => {
        actionState.flyOutForm.resolve({
            cancelled: false,
            values: formData
        });
    };

    const hideDialog = () => {
        actionDispatch({ type: "setFlyOutForm", payload: undefined });
    };

    const onFieldChange = (e: any) => {
        const value = e.target.value;
        const id = e.target.id;

        setFormData({...formData, [id]: value });
    };

    React.useEffect(() => {
        const lookups = fields.filter(([fieldId, field]) => field.type.toLowerCase() === "lookup");

        lookups.forEach(async ([fieldId, field]) => {
            const lookup = field as FlyOutLookupField;
            const entityNameGroups = /<\s*entity\s*name\s*=\s*["']([a-zA-Z_0-9]+)["']\s*>/gmi.exec(lookup.fetchXml);

            if (!entityNameGroups || !entityNameGroups.length) {
                return;
            }            

            const entityName = entityNameGroups[1];

            const data = await WebApiClient.Retrieve({ fetchXml: lookup.fetchXml, entityName: entityName, returnAllPages: true });
            setPickData({...pickData, [fieldId]: data.value.map((d: any) => ({ key: d[`${entityName}id`], name: d[lookup.displayField] ?? ("(No Name)") })) });
        });
    }, [ actionState.flyOutForm.fields ]);

    const textField = (fieldId: string, field: FlyOutField) => (
        <TextField key={fieldId} id={fieldId} description={field.subtext} required={field.required} multiline={field.rows && field.rows > 1} rows={field.rows ?? 1} type={field.type} label={field.label} placeholder={field.placeholder} onChange={onFieldChange} />
    );

    const onItemSelected = (fieldId: string, item: ITag) => {
        setFormData({ ...formData, [fieldId]: item?.key });
    };

    const getTextFromItem = (item: ITag) => item.name;

    const filterSelectedTags = (fieldId: string, filterText: string, tagList: ITag[]): ITag[] => {
        const data = pickData[fieldId];

        if (!data || !data.length) {
            return [];
        }

        if (!filterText) {
            return data;
        }        

        return data.filter(d => d.name.toLowerCase().indexOf(filterText.toLowerCase()) !== -1);
    };

    const pickerSuggestionsProps: IBasePickerSuggestionsProps = {
        suggestionsHeaderText: 'Suggested records',
        noResultsFoundText: 'No records found',
    };

    const inputProps: IInputProps = {
        'aria-label': 'Tag picker',
    };

    const lookupField = (fieldId: string, field: FlyOutField) => (
        <>
            <Label required={!!field.required}>{field.label}</Label>
            <TagPicker
                key={fieldId}
                removeButtonAriaLabel="Remove"
                onResolveSuggestions={(filter: string, selectedItems?: ITag[]) => filterSelectedTags(fieldId, filter, selectedItems)}
                onChange={(items) => onItemSelected(fieldId, items && items.length ? items[0] : null)}
                onEmptyResolveSuggestions={(selectedItems?: ITag[]) => filterSelectedTags(fieldId, "", selectedItems)}
                onRemoveSuggestion={() => onItemSelected(fieldId, null)}
                getTextFromItem={getTextFromItem}
                pickerSuggestionsProps={pickerSuggestionsProps}
                itemLimit={1}
                inputProps={inputProps}
            />
            { field.subtext && <Text styles={{root: { color: "#666666" } }} variant="small">{field.subtext}</Text> }
        </>
    );

    return (
        <UserInputModal okButtonDisabled={!Object.keys(actionState.flyOutForm.fields).every(fieldId => !actionState.flyOutForm.fields[fieldId].required || !!formData[fieldId])} noCallBack={noCallBack} yesCallBack={yesCallBack} finally={hideDialog} title={actionState.flyOutForm?.title} show={!!actionState.flyOutForm}>
            {Object.keys(actionState.flyOutForm.fields).map(fieldId => [ fieldId, actionState.flyOutForm.fields[fieldId]] as [string, FlyOutField]).map(([fieldId, field]) => field.type.toLowerCase() === "lookup" ? lookupField(fieldId, field) : textField(fieldId, field))}
        </UserInputModal>
    );
};