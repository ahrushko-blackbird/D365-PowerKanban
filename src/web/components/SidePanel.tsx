import * as React from "react";
import { useAppContext } from "../domain/AppState";

import { fetchData, refresh, fetchNotifications } from "../domain/fetchData";
import { SidePanelTile } from "./SidePanelTile";
import * as WebApiClient from "xrm-webapi-client";
import { FieldRow } from "./FieldRow";
import { useActionContext } from "../domain/ActionState";
import { useConfigState } from "../domain/ConfigState";

import { Card, ICardTokens, ICardSectionStyles, ICardSectionTokens } from '@uifabric/react-cards';
import { PrimaryButton, IconButton } from "@fluentui/react/lib/Button";
import { getSplitBorderButtonStyle, getSplitBorderContainerStyle } from "../domain/Internationalization";

interface SidePanelProps {
}

export const SidePanel = (props: SidePanelProps) => {
  const [ actionState, actionDispatch ] = useActionContext();
  const [ eventRecord, setEventRecord ] = React.useState(undefined);
  const configState = useConfigState();
  const [ appState, appDispatch ] = useAppContext();

  const notificationRecord = actionState.selectedRecord;
  const notifications = appState.notifications[actionState.selectedRecord.id] ?? [];
  const columns = Array.from(new Set(notifications.reduce((all, cur) => [...all, ...cur.parsed.updatedFields], [] as Array<string>)));
  const eventMeta = actionState.selectedRecord.entityType === configState.config.primaryEntity.logicalName ? configState.metadata : configState.secondaryMetadata[actionState.selectedRecord.entityType];

  React.useEffect(() => {
    const fetchEventRecord = async() => {
      const data = await WebApiClient.Retrieve({ entityName: actionState.selectedRecord.entityType, entityId: actionState.selectedRecord.id, queryParams: `?$select=${columns.join(",")}`, headers: [ { key: "Prefer", value: "odata.include-annotations=\"*\"" } ] });
      setEventRecord(data);
    };
    fetchEventRecord();
  }, []);

  const closeSideBySide = () => {
    actionDispatch({ type: "setSelectedRecord", payload: undefined });
  };

  const borderStyle = getSplitBorderContainerStyle(appState);
  const borderButtonStyle = getSplitBorderButtonStyle(appState);

  return (
    <div style={{overflow: "auto"}}>
        <Card tokens={{childrenGap: "10px"}} styles={{ root: { maxWidth: "auto", minWidth: "auto", margin: "5px", padding: "10px", backgroundColor: "#f8f9fa" }}}>
        { notifications.map(n => <SidePanelTile key={n.oss_notificationid} parent={notificationRecord} data={n}></SidePanelTile>)}
        </Card>
    </div>
  );
};