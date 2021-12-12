import * as React from "react";
import { useAppContext } from "../domain/AppState";

import { fetchData, refresh, fetchNotifications } from "../domain/fetchData";
import { SidePanelTile } from "./SidePanelTile";
import { SidePanelConfiguration } from "../domain/BoardViewConfig";
import * as WebApiClient from "xrm-webapi-client";
import { FieldRow } from "./FieldRow";
import { useActionContext } from "../domain/ActionState";
import { useConfigState } from "../domain/ConfigState";

import { Card, ICardTokens, ICardSectionStyles, ICardSectionTokens } from '@uifabric/react-cards';
import { PrimaryButton, IconButton } from "@fluentui/react/lib/Button";
import { getSplitBorderButtonStyle, getSplitBorderContainerStyle } from "../domain/Internationalization";

interface SidePanelProps {
  sidePanel: SidePanelConfiguration;
}

export const SidePanel = (props: SidePanelProps) => {
  const configState = useConfigState();
  const [ records, setRecords ] = React.useState([]);
  const [ appState, appDispatch ] = useAppContext();

  React.useEffect(() => {
    const fetchRecords = async() => {
      const data = await WebApiClient.Retrieve({ entityName: props.sidePanel.entity, fetchXml: props.sidePanel.fetchXml,  headers: [ { key: "Prefer", value: "odata.include-annotations=\"*\"" } ] });
      setRecords(data);
    };
    fetchRecords();
  }, []);

  return (
    <div style={{overflow: "auto"}}>
        <Card tokens={{childrenGap: "10px"}} styles={{ root: { maxWidth: "auto", minWidth: "auto", margin: "5px", padding: "10px", backgroundColor: "#f8f9fa" }}}>
        { records.map(n => <SidePanelTile key={n.userid} data={n}></SidePanelTile>)}
        </Card>
    </div>
  );
};