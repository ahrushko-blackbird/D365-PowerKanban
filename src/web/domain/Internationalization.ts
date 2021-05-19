import * as React from "react";
import { AppStateProps } from "./AppState";

export const getSplitBorderContainerStyle = (appState: AppStateProps): React.CSSProperties => {
    return appState.pcfContext.userSettings.isRTL
    ? { borderRight: "1px solid #777" }
    : { borderLeft: "1px solid #777" }
};

export const getSplitBorderButtonStyle = (appState: AppStateProps): React.CSSProperties => {
    return appState.pcfContext.userSettings.isRTL
    ? { right: "-18px" }
    : { left: "-18px" }
};