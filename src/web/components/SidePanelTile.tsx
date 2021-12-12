import * as React from "react";
import { useAppContext, useAppDispatch, AppStateProps, AppStateDispatch } from "../domain/AppState";
import { PrimaryButton, IconButton } from "@fluentui/react/lib/Button";
import { FieldRow } from "./FieldRow";
import { Metadata, Option } from "../domain/Metadata";
import { CardForm } from "../domain/CardForm";

import { refresh, fetchSubscriptions, fetchNotifications } from "../domain/fetchData";
import * as WebApiClient from "xrm-webapi-client";
import { Notification } from "../domain/Notification";
import { useConfigState } from "../domain/ConfigState";
import { useActionContext } from "../domain/ActionState";

import { Card, ICardTokens, ICardSectionStyles, ICardSectionTokens } from '@uifabric/react-cards';
import { ActivityItem } from "@fluentui/react/lib/ActivityItem";
import { Icon } from "@fluentui/react/lib/Icon";

interface SidePanelTileProps {
    data: Notification;
    style?: React.CSSProperties;
}

const SidePanelTileRender = (props: SidePanelTileProps) => {
    const configState = useConfigState();
    const appDispatch = useAppDispatch();
    const [ actionState, actionDispatch ] = useActionContext();

    const openInNewTab = () => {
    };

    return (
        <Card.Item>
            <ActivityItem
                key={props.data.oss_notificationid}
                activityIcon={<Icon iconName={'Chat' }/>}
                timeStamp={props.data["createdon@OData.Community.Display.V1.FormattedValue"]}
                activityDescription={[
                    <span key={1}>Event: {props.data["oss_event@OData.Community.Display.V1.FormattedValue"]}</span>,
                    <IconButton key={3} iconProps={{iconName: "OpenInNewWindow"}} title="Open in new window" onClick={openInNewTab}></IconButton>
                ]}
                comments={[
                    <span key={1}>{props.data.oss_text}</span>
                ]}
            />
        </Card.Item>
    );
};

export const SidePanelTile = React.memo(SidePanelTileRender);