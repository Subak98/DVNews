import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  IPropertyPaneCustomFieldProps,
  IPropertyPaneField,
  PropertyPaneDropdown,
  PropertyPaneFieldType,
  PropertyPaneLink,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "DvNewsWebPartStrings";
import DvNews from "./components/DvNews";
import { IDvNewsProps } from "./components/IDvNewsProps";
import { IServiceProvider } from "../../service/IServiceProvider";
import { getSP } from "../../service/pnpjsConfig";
import SharePointServiceProvider from "../../service/SharepointServiceProvider";
import {
  IPropertyFieldSite,
  PropertyFieldSitePicker,
} from "@pnp/spfx-property-controls/lib/PropertyFieldSitePicker";
import {
  PropertyFieldPeoplePicker,
  IPropertyFieldGroupOrPerson,
  PrincipalType,
} from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import { graphfi, SPFx as graphSPFx, GraphFI } from "@pnp/graph";
import PropertyFieldDropdownHost from "@pnp/spfx-property-controls/lib/propertyFields/dropdownWithCallout/PropertyFieldDropdownWithCalloutHost";

export interface IDvNewsWebPartProps {
  title: string;
  searchSites: IPropertyFieldSite[];
  adminUsers: IPropertyFieldGroupOrPerson[];
  layout: string;
  sitetitle: string;
}

export default class DvNewsWebPart extends BaseClientSideWebPart<IDvNewsWebPartProps> {
  private _serviceProvider: IServiceProvider | any;
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";
  private graph: GraphFI | any;

  public render(): void {
    const element: React.ReactElement<IDvNewsProps> = React.createElement(
      DvNews,
      {
        title: this.properties.title,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this.context.pageContext.user.email,
        provider: this._serviceProvider,
        context: this.context,
        searchSites: this.properties.searchSites,
        adminUsers: this.properties.adminUsers,
        layout: this.properties.layout,
        sitetitle: this.properties.sitetitle,
      },
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._serviceProvider = new SharePointServiceProvider(this.context);
    getSP(this.context);
    this.graph = graphfi().using(graphSPFx(this.context));
    return this._getEnvironmentMessage().then((message) => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
            case "TeamsModern":
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment,
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null,
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null,
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
  public renderLink(searchSites: any[]): any[] {
    if (
      this.properties.searchSites !== undefined &&
      this.properties.searchSites.length > 0
    ) {
      return [
        PropertyPaneLink("", {
          target: "_blank",
          href: `${this.properties.searchSites[0].url}/SitePages/Forms/ByAuthor.aspx`,
          text: "Edit News Articles",
        }),
      ];
    } else {
      return [];
    }
  }
  private getSitePickerLabel(): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
    return {
      type: PropertyPaneFieldType.Custom,
      targetProperty: "searchSites",

      properties: {
        key: "sitePickerInfo",

        onRender: (elem: HTMLElement): void => {
          elem.innerHTML = `
          <div style="
            display: flex;
            align-items: center;
            gap: 6px;
            margin-bottom: 8px;
          ">
            <span style="
              font-size: 14px;
              font-weight: 600;
              color: #323130;
            ">
              Search sites
            </span>

            <span
              title="Select the SharePoint sites where you want to upload news articles."
              style="
                display: inline-flex;
                align-items: center;
                justify-content: center;
                width: 16px;
                height: 16px;
                border-radius: 50%;
                background-color: #605e5c;
                color: #ffffff;
                font-size: 11px;
                font-weight: 600;
                
              "
            >
              i
            </span>
          </div>
        `;
        },

        onDispose: (elem: HTMLElement): void => {
          elem.innerHTML = "";
        },
      },
    };
  }
  private getLayoutLabel(): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
    return {
      type: PropertyPaneFieldType.Custom,
      targetProperty: "searchSites",

      properties: {
        key: "sitePickerInfo",

        onRender: (elem: HTMLElement): void => {
          elem.innerHTML = `
          <div style="
            display: flex;
            align-items: center;
            gap: 6px;
            margin-bottom: 8px;
          ">
            <span style="
              font-size: 14px;
              font-weight: 600;
              color: #323130;
            ">
              Select Layout
            </span>

            <span
              title="Select the layout for the News web part. *Home* displays news tagged as “All”, while *Department* displays news based on the current department site."
              style="
                display: inline-flex;
                align-items: center;
                justify-content: center;
                width: 16px;
                height: 16px;
                border-radius: 50%;
                background-color: #605e5c;
                color: #ffffff;
                font-size: 11px;
                font-weight: 600;
                
              "
            >
              i
            </span>
          </div>
        `;
        },

        onDispose: (elem: HTMLElement): void => {
          elem.innerHTML = "";
        },
      },
    };
  }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("title", {
                  label: "Web Part Title",
                  value: "News",
                }),
                this.getSitePickerLabel(),
                PropertyFieldSitePicker("searchSites", {
                  label: "",
                  initialSites: this.properties.searchSites,
                  context: this.context,
                  deferredValidationTime: 500,
                  multiSelect: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties.searchSites,
                  key: "sitesFieldId",
                }),
                ...this.renderLink(this.properties.searchSites),
                PropertyFieldPeoplePicker("adminUsers", {
                  label: "Target audience",
                  initialData: this.properties.adminUsers,
                  allowDuplicate: false,
                  principalType: [
                    PrincipalType.Users,
                    // PrincipalType.SharePoint,
                    // PrincipalType.Security,
                  ],
                  searchTextLimit: 2,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  // onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: "peopleFieldId",
                }),
                this.getLayoutLabel(),
                PropertyPaneDropdown("layout", {
                  label: "",
                  options: [
                    { key: "Home", text: "Home" },
                    { key: "Department", text: "Department" },
                  ],
                }),
                // PropertyPaneTextField("sitetitle", {
                //   label: "Site Title",
                // }),
              ],
            },
          ],
        },
      ],
    };
  }
}
