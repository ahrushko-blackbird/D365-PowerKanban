import * as React from "react";
import { useAppContext } from "../domain/AppState";
import { PrimaryButton, IconButton } from "@fluentui/react/lib/Button";

import { refresh } from "../domain/fetchData";
import { useActionDispatch, useActionContext } from "../domain/ActionState";
import { useConfigDispatch, useConfigState } from "../domain/ConfigState";
import { getSplitBorderButtonStyle, getSplitBorderContainerStyle } from "../domain/Internationalization";
import { Card, ICardTokens, ICardSectionStyles, ICardSectionTokens } from '@uifabric/react-cards';
import { SidePanelTile } from "./SidePanelTile";
import { SidePanel } from "./SidePanel";
import { Pivot, PivotItem } from "@fluentui/react/lib/components/Pivot";

interface FormProps {
}

export const SidePanelHost = (props: FormProps) => {
  const [appState, appDispatch] = useAppContext();
  const [actionState, actionDispatch] = useActionContext();
  const configState = useConfigState();

  const closeSideBySide = () => {
    actionDispatch({ type: "setSelectedRecordDisplayType", payload: undefined });
    actionDispatch({ type: "setSelectedRecord", payload: undefined });
  };

  const closeAndRefresh = async () => {
    actionDispatch({ type: "setSelectedRecordDisplayType", payload: undefined });
    actionDispatch({ type: "setSelectedRecord", payload: undefined });

    await refresh(appDispatch, appState, configState, actionDispatch, actionState);
  };  

  const borderStyle = getSplitBorderContainerStyle(appState);
  const borderButtonStyle = getSplitBorderButtonStyle(appState);

  return (
      <div style={{ ...borderStyle, position: "relative", width: "100%", height: "100%" }}>
        <IconButton iconProps={{iconName: "ChromeClose"}} title="Close" onClick={closeSideBySide} style={{ ...borderButtonStyle, color: "white", backgroundColor: "#045999", position: "absolute", top: "calc(50% - 20px)", left: "-18px" }}></IconButton>
        <IconButton iconProps={{iconName: "Refresh"}} title="Close and refresh" onClick={closeAndRefresh} style={{ ...borderButtonStyle, color: "white", backgroundColor: "#045999", position: "absolute", top: "calc(50% +  20px)", left: "-18px" }}></IconButton>
        <Pivot aria-label="Side Panels">
          {
            configState.config.sidePanels.map(p => <PivotItem headerText={p.headerText} key={p.uniqueName ?? p.headerText}><SidePanel sidePanel={p} /></PivotItem>)
          }
        </Pivot>
      </div>
  );
};