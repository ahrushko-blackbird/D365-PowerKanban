import { FlyOutForm } from "./FlyOutForm";

export interface CustomButton {
    id: string;
    icon: { type: string; value: string; };
    label: string;
    callBack: string;
}

export interface BoardEntity {
    logicalName: string;
    swimLaneSource: string;
    hiddenLanes: Array<number>;
    visibleLanes: Array<number>;
    transitionCallback: string;
    notificationLookup: string;
    subscriptionLookup: string;
    preventTransitions: boolean;
    defaultView: string;
    customButtons: Array<CustomButton>;
    fitLanesToScreenWidth: boolean;
    hideCountOnLane: boolean;
    defaultOpenHandler: "inline" | "sidebyside" | "modal" | "newwindow";
}

export interface SecondaryEntity extends BoardEntity {
    parentLookup: string;
}

export interface Context {
    showForm: (form: FlyOutForm) => Promise<any>;
}

export interface PrimaryEntity extends BoardEntity {

}

export interface BoardViewConfig {
    primaryEntity: PrimaryEntity;
    secondaryEntity: SecondaryEntity;
    customScriptUrl: string;
}