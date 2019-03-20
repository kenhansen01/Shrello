// library imports
if (Map.toString().indexOf("function Map()") === -1) {
  Map = undefined;
}
if (Promise.toString().indexOf("function Promise()") === -1) {
  Promise = undefined;
}
if (Set.toString().indexOf("function Set()") === -1) {
  Set = undefined;
}
if (WeakMap.toString().indexOf("function WeakMap()") === -1) {
  WeakMap = undefined;
}
if (WeakSet.toString().indexOf("function WeakSet()") === -1) {
  WeakSet = undefined;
}

import "core-js/es6";
import "whatwg-fetch";

import * as React from "react";
import * as ReactDom from "react-dom";
// import * as core from "core-js";
import {
  Version,
  Environment,
  EnvironmentType
} from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneButton,
  PropertyPaneButtonType,
  IWebPartContext
} from "@microsoft/sp-webpart-base";
import { sp, SPRest, Web, SiteUserProps } from "@pnp/sp";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
// webpart imports
import * as strings from "ShrelloAssetsViewWebPartStrings";
import IShrelloAssetsViewWebPartProps from "./IShrelloAssetsViewWebPartProps";
// component imports
import { IShrelloContainerProps,ShrelloContainer} from "./components/shrelloContainer";

import { ProvisionLists } from "../../assets/lists/ProvisionLists";

export default class ShrelloAssetsViewWebPart extends BaseClientSideWebPart<IShrelloAssetsViewWebPartProps> {

  private _sp: SPRest;
  private _currentUser: SiteUserProps;
  private _web: Web;
  private _pLists: ProvisionLists;
  private _shrelloContainerComponent: ShrelloContainer;

  private async provisionAssets(): Promise<void> {
    this._pLists = new ProvisionLists(this._sp);
    this.properties.assetsProvisioned = true;
    this.context.propertyPane.refresh();
    return;
  }

  protected async onInit(): Promise<void> {
    // tslint:disable-next-line:no-string-literal
    window["g_wsaenabled"] = false;

    this.context.statusRenderer.displayLoadingIndicator(this.domElement, "TASC");

    this._openPropertyPane = this._openPropertyPane.bind(this);

    await super.onInit();
    sp.setup({
      spfxContext: this.context,
      sp: {
        headers: {
          "Accept": "application/json;odata=verbose",
        },
        baseUrl: this.context.pageContext.web.absoluteUrl
      },
      enableCacheExpiration: true,
      cacheExpirationIntervalMilliseconds: 1000
    });
    initializeIcons();
    this._sp = sp;
    this._web = sp.web;
    this._currentUser = await sp.web.currentUser.get<SiteUserProps>();
    /**
     * Create the appropriate data provider, depending on where the webpart is running.
     * The DEBUG flag ensures mock data is not bundled.
     */
    // if (DEBUG && Environment.type === EnvironmentType.Local) {
    //   this._dataProvider = new MockDataProvider();
    // } else {
      // this._dataProvider = new SharePointDataProvider();
      // this._dataProvider.sp = sp;
      // this._dataProvider.webPartContext = this.context;
    // }
    // this._supportDepartmentDataProvider = await this._setDataProvider<ISupportDepartmentItem>(this.context, "Support Departments");
    // this._supportTopicDataProvider = await this._setDataProvider<ISupportTopicItem>(this.context, "Support Topics");
    // this._shrelloDataProvider = await this._setDataProvider<IShrelloItem>(this.context, "TASC Tickets");
    this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    return;
  }

  protected get dataVersion(): Version {
    return Version.parse("2.0.0");
  }

  private _openPropertyPane(): void {
    this.context.propertyPane.open();
  }

  public render(): void {
    require("./ShrelloViewStyles.css");
    const contentArea: Element = document.getElementById("s4-bodyContainer")
    || document.getElementsByClassName("ControlZone-control")[0];
    const element: React.ReactElement<IShrelloContainerProps> = React.createElement(
      ShrelloContainer,
      {
        sp: this._sp,
        context: this.context,
        contentArea: contentArea,
        currentUser: this._currentUser,
        webPartDisplayMode: this.displayMode,
        configureStartCallback: this._openPropertyPane,
      }
    );

    this._shrelloContainerComponent = <ShrelloContainer>ReactDom.render(element, this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel
                })
              ]
            },
            {
              groupName: strings.ProvisioningGroupName,
              groupFields: [
                PropertyPaneCheckbox("assetsProvisioned", {
                  text: "Assets Provisioned?"
                }),
                PropertyPaneButton("provisionAssets", {
                  text: "Provision Assets",
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this.provisionAssets.bind(this)
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
