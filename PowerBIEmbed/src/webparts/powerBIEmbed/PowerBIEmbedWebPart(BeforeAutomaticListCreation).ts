import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import PowerBiEmbed from "./components/PowerBIEmbed";
import { IPowerBiEmbedProps } from "./components/IPowerBIEmbedProps";
//import { sp } from "@pnp/sp/presets/all";

export interface IPowerBiEmbedWebPartProps {
  reportId: string;
  datasetId: string;
  hasRLS: boolean;
  reportWorkspaceId: string;
  datasetWorkspaceId: string;
}

export default class PowerBiEmbedWebPart extends BaseClientSideWebPart<IPowerBiEmbedWebPartProps> {
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }
  protected onAfterPropertyPaneChangesApplied(): void {
    this.render();
  }

  public render(): void {
    this._renderAsync().catch((err) => console.error("Render error:", err));
  }

  private async _renderAsync(): Promise<void> {
    const userEmail = this.context.pageContext.user.email || "";
    const reportId = this.properties.reportId;
    const datasetId = this.properties.datasetId;
    const reportWorkspaceId = this.properties.reportWorkspaceId;
    const datasetWorkspaceId = this.properties.datasetWorkspaceId;
    const hasRLS = this.properties.hasRLS;
    const element: React.ReactElement<IPowerBiEmbedProps> = React.createElement(
      PowerBiEmbed,
      {
        userEmail,
        reportId,
        datasetId,
        hasRLS,
        reportWorkspaceId,
        datasetWorkspaceId,
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: "Report",
              groupFields: [
                PropertyPaneTextField("reportId", {
                  label: "Report Id",
                }),
                PropertyPaneTextField("reportWorkspaceId", {
                  label: "Report Workspace Id",
                }),
              ],
            },
            {
              groupName: "Dataset",
              groupFields: [
                PropertyPaneTextField("datasetId", {
                  label: "Dataset Id",
                }),
                PropertyPaneTextField("datasetWorkspaceId", {
                  label: "Dataset Workspace Id",
                }),
              ],
            },
            {
              groupName: "Role",
              groupFields: [
                PropertyPaneToggle("hasRLS", {
                  label: "Has Roles",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
