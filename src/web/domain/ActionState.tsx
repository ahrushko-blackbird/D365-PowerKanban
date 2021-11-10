import * as React from "react";
import { SavedQuery } from "./SavedQuery";
import { CardForm } from "./CardForm";
import { FlyOutForm } from "./FlyOutForm";

export enum DisplayType {
    recordForm,
    notifications
}

type Action = { type: "setSelectedRecord", payload: Xrm.LookupValue }
    | { type: "setSelectedForm", payload: CardForm }
    | { type: "setSelectedSecondaryView", payload: SavedQuery }
    | { type: "setSelectedSecondaryForm", payload: CardForm }
    | { type: "setProgressText", payload: string | undefined }
    | { type: "setWorkIndicator", payload: boolean}
    | { type: "setSelectedRecordDisplayType", payload: DisplayType }
    | { type: "setFlyOutForm", payload: FlyOutForm }
    | { type: "setConfigSelectorDisplayState", payload: boolean }
    | { type: "setSelectedRecords", payload: {[key: string]: boolean} };

export type ActionDispatch = (action: Action) => void;

export type ActionStateProps = {
    progressText?: string;
    selectedForm?: CardForm;
    selectedViewData?: { columns: Array<string>; linkEntities: Array<{ entityName: string, alias: string }> }
    selectedSecondaryView?: SavedQuery;
    selectedSecondaryForm?: CardForm;
    selectedSecondaryViewData?: { columns: Array<string>; linkEntities: Array<{ entityName: string, alias: string }> }
    selectedRecord?: Xrm.LookupValue;
    workIndicator?: boolean;
    selectedRecordDisplayType?: DisplayType;
    flyOutForm?: FlyOutForm;
    configSelectorDisplayState?: boolean;
    selectedRecords?: {[key: string]: boolean};
};

type ActionContextProps = {
    children: React.ReactNode;
};

const parseLayoutColumns = (layoutXml: string): Array<string> => {
    const parser = new DOMParser();
    const xml = parser.parseFromString(layoutXml, "application/xml");
    return Array.from(xml.documentElement.getElementsByTagName("cell")).map(c => c.getAttribute("name")!);
};

const parseLinksFromFetch = (fetchXml: string): Array<{ entityName: string, alias: string }> => {
    const parser = new DOMParser();
    const xml = parser.parseFromString(fetchXml, "application/xml");
    return Array.from(xml.documentElement.getElementsByTagName("link-entity")).map(c => ({ entityName: c.getAttribute("name")!, alias: c.getAttribute("alias")!}));
};

function stateReducer(state: ActionStateProps, action: Action): ActionStateProps {
    switch (action.type) {
        case "setSelectedRecord": {
            return { ...state, selectedRecord: action.payload };
        }
        case "setSelectedRecords": {
            return { ...state, selectedRecords: { ...state.selectedRecords, ...action.payload } };
        }
        case "setSelectedForm": {
            return { ...state, selectedForm: action.payload };
        }
        case "setSelectedSecondaryView": {
            return { ...state, selectedSecondaryView: action.payload, selectedSecondaryViewData: { columns: parseLayoutColumns(action.payload.layoutxml), linkEntities: parseLinksFromFetch(action.payload.fetchxml) } };
        }
        case "setSelectedSecondaryForm": {
            return { ...state, selectedSecondaryForm: action.payload };
        }
        case "setProgressText": {
            return { ...state, progressText: action.payload };
        }
        case "setWorkIndicator": {
            return { ...state, workIndicator: action.payload };
        }
        case "setSelectedRecordDisplayType": {
            return { ...state, selectedRecordDisplayType: action.payload };
        }
        case "setFlyOutForm": {
            return { ...state, flyOutForm: action.payload };
        }
        case "setConfigSelectorDisplayState": {
            return { ...state, configSelectorDisplayState: action.payload };
        }
    }
}

const ActionState = React.createContext<ActionStateProps | undefined>(undefined);
const ActionDispatch = React.createContext<ActionDispatch | undefined>(undefined);

export function ActionStateProvider({ children }: ActionContextProps) {
    const [state, dispatch] = React.useReducer(stateReducer, { });

    return (
        <ActionState.Provider value={state}>
            <ActionDispatch.Provider value={dispatch}>
                {children}
            </ActionDispatch.Provider>
        </ActionState.Provider>
    );
}

export function useActionState() {
    const context = React.useContext(ActionState);

    if (!context) {
        throw new Error("useActionState must be used within a state provider!");
    }

    return context;
}

export function useActionDispatch() {
    const context = React.useContext(ActionDispatch);

    if (!context) {
        throw new Error("useActionDispatch must be used within a state provider!");
    }

    return context;
}

export function useActionContext(): [ ActionStateProps, ActionDispatch ] {
    return [ useActionState(), useActionDispatch() ];
}