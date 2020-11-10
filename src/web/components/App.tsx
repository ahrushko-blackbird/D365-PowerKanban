import * as React from "react";
import { AppStateProvider } from "../domain/AppState";
import { SplitView } from "./SplitView";
import { ActionStateProvider } from "../domain/ActionState";
import { ConfigStateProvider } from "../domain/ConfigState";
import { ErrorBoundary } from "./ErrorBoundary";

export interface AppProps
{
  configId?: string;
  primaryEntityLogicalName?: string;
  primaryEntityId?: string;
  appId?: string;
  primaryDataIds?: Array<string>;
}

export const App: React.FC<AppProps> = (props) => {
  return (
    <ErrorBoundary>
      <AppStateProvider primaryDataIds={props.primaryDataIds} primaryEntityId={props.primaryEntityId}>
        <ActionStateProvider>
          <ConfigStateProvider appId={props.appId} configId={props.configId} primaryEntityLogicalName={props.primaryEntityLogicalName}>
            <ErrorBoundary>
              <SplitView primaryDataIds={props.primaryDataIds} />
            </ErrorBoundary>
          </ConfigStateProvider>
        </ActionStateProvider>
      </AppStateProvider>
    </ErrorBoundary>
  );
};