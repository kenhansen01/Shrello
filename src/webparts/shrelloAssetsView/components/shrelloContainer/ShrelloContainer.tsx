import * as React from "react";
import { css, Fabric, CommandBar, ICommandBarItemProps, IContextualMenuItem } from "office-ui-fabric-react";
import { DisplayMode } from "@microsoft/sp-core-library";
import { IconNames } from "@uifabric/icons/lib/IconNames";
import {
  SiteUserProps,
  sp,
  ItemAddResult,
  AttachmentFileInfo,
  PagedItemCollection,
  ItemUpdateResult
} from "@pnp/sp";
import { IPrincipal } from "@pnp/spfx-controls-react/lib/common/SPEntities";
import update from "immutability-helper";

import { IShrelloItem, ISupportDepartmentItem, ISupportTopicItem } from "../../models";
import { ListTitles } from "../../config/shrelloViewConfig";

import { IShrelloContainerProps, IShrelloContainerState } from "./IShrelloContainer";
import styles from "./ShrelloContainer.module.scss";

import { ConfigurationView } from "../configurationView";
import {
  viewData,
  viewName,
  IViewObject,
  ShrelloList
} from "../shrelloList";
import ShrelloForm from "../shrelloForm/ShrelloForm";

const shrelloCommands: ICommandBarItemProps[] = [
  {
    key: "newRequest",
    name: "Add a card.",
    onClick: (ev) => { ev.preventDefault(); return false;},
    iconProps: {
      iconName: IconNames.Add
    },
    subMenuProps: {
      items: []
    },
    ariaLabel: "New. Use left and right arrow keys to navigate",
    ["data-automation-id"]: "newItemMenu"
  }
];

const shrelloFarItems: ICommandBarItemProps[] = [
  {
    key: "viewSelect",
    name: "Change View",
    onClick: (ev) => { ev.preventDefault(); return false;},
    iconProps: {
      iconName: IconNames.View
    },
    ["data-automation-id"]: "viewSelectMenu",
    subMenuProps: {
      items: viewData.filter(view => view.menuProps.name !== viewName.teamShrellos).map(view => view.menuProps)
    }
  },
  {
    key: "fullScreen",
    name: "Full Screen",
    iconProps: {
      iconName: "FullScreen"
    },
    ["data-automation-id"]: "fullScreen",
    onClick: (ev) => { ev.preventDefault(); return false;}
  }
];

function ShrelloFabricContainer(props: any): JSX.Element {
  return (
    <Fabric>
      { props.children }
    </Fabric>
  );
}

function ShrelloFabricDisplay(props: any): JSX.Element {
  return (
    <ShrelloFabricContainer>
      {
        props.showPlaceholder &&
        props.webPartDisplayMode === DisplayMode.Edit &&
        <ConfigurationView
          icon={ "Edit" }
          iconText="TASCs"
          description="TASC items, so you can get things done."
          buttonLabel="Provision Assets"
          onConfigure={ props.onConfigure }
        />
      }
      { props.showPlaceholder &&
        props.webPartDisplayMode === DisplayMode.Read &&
          <ConfigurationView
            icon={ "Edit" }
            iconText="TASCs"
            description="Get things done. Organize and share your teams TASCs. Edit this web part to start managing TASCs." />
        }
        { !props.showPlaceholder &&
          <div
            className={
              css(styles.shrello, {
                [styles.shrelloIsFullScreen]: props.fullScreen
              })
            }
            style={ props.shrelloTop }
          >
            { props.children }
          </div>
        }
    </ShrelloFabricContainer>
  );
}

function ShrelloToolsContainer(props: any): JSX.Element {
  return (
    <ShrelloFabricDisplay {...props}>
      <div
        className={ styles.shrelloContainer }
        style={{
          height: props.contentArea.clientHeight - 35
        }}
      >
        <div className={ styles.topRow }>
          <h2 className={ styles.shrelloHeading }>{ props.shrelloHeading }</h2>
        </div>
        { props.children }
      </div>
    </ShrelloFabricDisplay>
  );
}

function ShrelloTools(props: any): JSX.Element {
  return (
    <ShrelloToolsContainer {...props}>
      <CommandBar
        items={ props.shrelloCommands }
        farItems={ props.shrelloFarCommands }
      />
      <ShrelloList {...props}
        shrelloItems={ props.shrelloItems }
        supportDepartments={ props.supportDepartmentItems }
        supportTopics={ props.supportTopicItems }
        currentUser={ props.currentUser }
        selectedView={ props.selectedView }
        handleItemClick={ props.handleItemClick }
      />
      {
        props.showForm &&
        <ShrelloForm {...props}
          supportDepartments={ props.supportDepartmentItems }
          supportTopics={ props.supportTopicItems }
        />
      }
    </ShrelloToolsContainer>
  );
}

export default class ShrelloContainer1 extends React.Component<IShrelloContainerProps, IShrelloContainerState> {
  private _currentUser: SiteUserProps;
  private _shrelloItems: IShrelloItem[];
  private _pagedShrelloItems: PagedItemCollection<any>;
  private _selectedView: IViewObject;
  private _supportDepartments: ISupportDepartmentItem[];
  private _selectedSupportDepartment: ISupportDepartmentItem;
  private _supportTopics: ISupportTopicItem[];
  private _shrelloListEntityType: string;
  private _supportTopicEntityType: string;
  private _contentArea: Element;
  private _shrelloFarCommands: ICommandBarItemProps[];

  public constructor(props: IShrelloContainerProps) {
    super(props);

    this._configureWebPart = this._configureWebPart.bind(this);
    this._openItem = this._openItem.bind(this);
    this._createShrelloItem = this._createShrelloItem.bind(this);
    this._updateShrelloItem = this._updateShrelloItem.bind(this);
    this._onCloseForm = this._onCloseForm.bind(this);
    this._selectView = this._selectView.bind(this);
    this._onSelectTeamView = this._onSelectTeamView.bind(this);
    this._setView = this._setView.bind(this);
    this._refreshPagedShrelloItems = this._refreshPagedShrelloItems.bind(this);
    this._fullScreen = this._fullScreen.bind(this);

    this._contentArea = props.contentArea || document.getElementById("DeltaPlaceHolderMain")
    || document.getElementsByClassName("ControlZone-control")[0];

    this._selectedView = viewData.find(view =>
      view.menuProps.name === viewName.myRequests);

    this._shrelloFarCommands = shrelloFarItems;

    this.state = {
      shrelloHeading: "",
      fullScreen: false,
      showForm: false,
      supportDepartmentItems: [],
      supportTopicItems: [],
      shrelloItems: [],
      pagedShrelloItems: null,
      contentArea: this._contentArea,
      currentUser: props.currentUser || null,
      shrelloCommands: shrelloCommands,
      shrelloFarCommands: this._shrelloFarCommands,
      selectedView: this._selectedView
    };
  }

  public async componentDidMount(): Promise<void> {
    this._currentUser = await sp.web.currentUser.get<SiteUserProps>();
    this._supportDepartments = await sp.web.lists.getByTitle(ListTitles.SupportDepartments).items.get();
    this._supportTopics = await sp.web.lists.getByTitle(ListTitles.SupportTopics).items.get();
    this._shrelloListEntityType = await sp.web.lists.getByTitle(ListTitles.ShrelloItems).getListItemEntityTypeFullName();
    this._supportTopicEntityType = await sp.web.lists.getByTitle(ListTitles.SupportTopics).getListItemEntityTypeFullName();
    this._pagedShrelloItems = await this._getPagedItems();

    let newRequestItems: IContextualMenuItem[] = this._supportDepartments.map(department => {
      const sMenuItem: ICommandBarItemProps = {
        key: String(department.Id),
        name: department.Title,
        iconProps: {
          iconName: department.IconName
        },
        ["data-automation-id"]: `new${department.Title.replace(" ", "")}`,
        onClick: this._showForm
      };
      return sMenuItem;
    });

    let fScreenCommand: ICommandBarItemProps = this._shrelloFarCommands.find(item => item.key === "fullScreen");

    shrelloCommands.find(item => item.key === "newRequest").subMenuProps.items = newRequestItems;
    this._shrelloFarCommands.find(item => item.key === "viewSelect").subMenuProps.items.map(item => {
      item.onClick = this._selectView;
      return item;
    });
    fScreenCommand.onClick = this._fullScreen;
    fScreenCommand.name = this.state.fullScreen ? "Minimize" : "Full Screen";
    fScreenCommand.iconProps.iconName = this.state.fullScreen ? IconNames.BackToWindow : IconNames.FullScreen;

    const uDepartmentItems: ReadonlyArray<ISupportDepartmentItem> = update(this.state.supportDepartmentItems, {
      $set: this._supportDepartments
    });
    const uTopicItems: ReadonlyArray<ISupportTopicItem> = update(this.state.supportTopicItems, {
      $set: this._supportTopics
    });
    const uShrelloItems: ReadonlyArray<IShrelloItem> = update(this.state.shrelloItems, {
      $set: this._shrelloItems
    });
    const uPagedItems: PagedItemCollection<any> = update(this.state.pagedShrelloItems, {
      $set: this._pagedShrelloItems
    });

    this.setState({
      currentUser: this._currentUser,
      shrelloItems: uShrelloItems as IShrelloItem[],
      pagedShrelloItems: uPagedItems,
      supportDepartmentItems: uDepartmentItems as ISupportDepartmentItem[],
      supportTopicItems: uTopicItems as ISupportTopicItem[],
      shrelloCommands: shrelloCommands,
      shrelloFarCommands: this._shrelloFarCommands,
      // contentArea: contentArea,
      shrelloListEntityType: this._shrelloListEntityType
    });
  }

  public render(): JSX.Element {

    return (
      <ShrelloTools
      {...this.state}
        context={ this.props.context }
        onConfigure={ this._configureWebPart}
        handleItemClick={ this._openItem }
        onAddShrelloItem={ this._createShrelloItem }
        onUpdateShrelloItem={ this._updateShrelloItem }
        onSelectTeamView={ this._onSelectTeamView }
        onCloseForm={ this._onCloseForm }
        getShrelloItems={ this._refreshPagedShrelloItems }
      />
    );
  }

  private _configureWebPart = (): void => this.props.configureStartCallback();

  private _fullScreen = (ev: React.MouseEvent<HTMLElement>) => {
    ev.preventDefault();
    let fScreenCommand: ICommandBarItemProps = this._shrelloFarCommands.find(item => item.key === "fullScreen");
    fScreenCommand.name = !this.state.fullScreen ? "Minimize" : "Full Screen";
    fScreenCommand.iconProps.iconName = !this.state.fullScreen ? IconNames.BackToWindow : IconNames.FullScreen;
    this.setState({
      fullScreen: !this.state.fullScreen,
      shrelloFarCommands: this._shrelloFarCommands
    });
  }

  private _selectView = async(ev: React.MouseEvent<HTMLElement>, item: IContextualMenuItem): Promise<void> => {
    this._selectedView = viewData.find(view =>
      view.menuProps.name === item.name);
    return await this._setView();
  }

  private _onSelectTeamView = async(
    ev: React.MouseEvent<HTMLElement>,
    laneDepartment: ISupportDepartmentItem
  ): Promise<void> => {
    this._selectedView = viewData.find(view => view.menuProps.name === viewName.teamShrellos);
    this._selectedSupportDepartment = laneDepartment;
    return await this._setView();
  }

  private _setView = async(selectedView?: IViewObject): Promise<void> => {
    // this._selectedView = selectedView || this._selectedView;
    let filter: string = "";
    if (this._selectedView.menuProps.name === viewName.teamShrellos) {
      filter = this._selectedView.menuProps.filter(this._selectedSupportDepartment.Id);
    } else {
      filter = this._selectedView.menuProps.filter(this.state.currentUser.Id);
    }
    this._pagedShrelloItems = await this._getPagedItems(filter);
    const uPagedItems: PagedItemCollection<any> = update(this.state.pagedShrelloItems, {
      $set: this._pagedShrelloItems
    });
    this.setState({
      selectedView: this._selectedView,
      selectedSupportDepartment: this._selectedSupportDepartment,
      pagedShrelloItems: uPagedItems,
      shrelloHeading: this._selectedView.menuProps.name
    });
  }

  private _onCloseForm = async (): Promise<void> => {
    this._pagedShrelloItems = await this._getPagedItems();
    const uPagedItems: PagedItemCollection<any> = update(this.state.pagedShrelloItems, {
      $set: this._pagedShrelloItems
    });
    this.setState({ pagedShrelloItems: uPagedItems, showForm: false });
  }

  private _showForm = (ev: React.MouseEvent<HTMLElement>, item: IContextualMenuItem): void => {
    const tItem: IShrelloItem = {
      Title: undefined,
      TASCTypeCategory: "Unknown",
      Priority: "3 - Normal",
      SupportDepartmentId: parseInt(item.key, 10)
    };
    this.setState({ showForm: true, item: tItem });
  }

  private _refreshPagedShrelloItems = async (
    filter?: string,
    orderby?: {prop: string, ascending?: boolean}
  ): Promise<void> => {
    const pagedRefreshItems: PagedItemCollection<any> = await this._getPagedItems(filter, orderby);
    const updatedPagedItems: PagedItemCollection<any> = update(this.state.pagedShrelloItems, {
      $set: pagedRefreshItems
    });
    this.setState({ pagedShrelloItems: updatedPagedItems });
  }

  private _getPagedItems = async (
    filter?: string,
    orderby?: {prop: string, ascending?: boolean}
  ): Promise<PagedItemCollection<any>> => {

    let fltr: string = filter || "";
    if (this._selectedView.menuProps.name === viewName.teamShrellos) {
      fltr = this._selectedView.menuProps.filter(this._selectedSupportDepartment.Id);
    } else {
      fltr = this._selectedView.menuProps.filter(this.state.currentUser.Id);
    }

    const odrby: {prop: string, ascending?: boolean} = orderby ? orderby : {prop: "Created", ascending: false}; //change prop to modified
    return sp.web.lists.getByTitle(ListTitles.ShrelloItems)
      .items
      .filter(fltr)
      .orderBy(odrby.prop, odrby.ascending)
      .expand("AttachmentFiles")
      .top(2000)
      .getPaged();
  }

  private _createShrelloItem = async (item: IShrelloItem, attachments?: File[]): Promise<void> => {
    item.Requester = undefined;
    item.AssignedTo = undefined;
    item.Watching = undefined;
    item.AttachmentFiles = undefined;

    const uItem: any = item;
    uItem.AssignedToId = item.AssignedToId && !!item.AssignedToId.length
      ? { results: item.AssignedToId }
      : undefined;
    uItem.WatchingId = item.WatchingId && !!item.WatchingId.length
      ? { results: item.WatchingId }
      : undefined;

    const newItemAddResult: ItemAddResult = await sp.web.lists.getByTitle(ListTitles.ShrelloItems)
      .items
      .add(uItem, this.state.shrelloListEntityType);
    const newItem: IShrelloItem = await newItemAddResult.item
      .get<IShrelloItem>();
    if (attachments) {
      await this._addItemAttachments(newItem.Id, attachments);
    }

    this.setState({
      newShrelloItem: newItem
    });
    return;
  }

  private _openItem = async (ev: React.MouseEvent<HTMLElement>, item: IShrelloItem): Promise<void> => {
    let assignedIds: number[], watchingIds: number[];
    if (item.RequesterId) {
      item.Requester = await this._resolvePrincipal(item.RequesterId);
    }
    if(item.AssignedToId && item.AssignedToId instanceof Array) {
      assignedIds = item.AssignedToId as number[];
    } else if (item.AssignedToId && item.AssignedToId.hasOwnProperty("results")) {
      // tslint:disable-next-line:no-string-literal
      assignedIds = item.AssignedToId["results"] as number[];
    }
    if (!!assignedIds) {
      item.AssignedTo = await Promise.all<IPrincipal>(assignedIds.map(async id => {
          const aPrincipal: IPrincipal = await this._resolvePrincipal(id);
          return aPrincipal;
        }));
    }
    if(item.WatchingId && item.WatchingId instanceof Array) {
      watchingIds = item.WatchingId as number[];
    } else if (item.WatchingId && item.WatchingId.hasOwnProperty("results")) {
      // tslint:disable-next-line:no-string-literal
      watchingIds = item.WatchingId["results"] as number[];
    }
    if (!!watchingIds) {
      item.Watching = await Promise.all<IPrincipal>(watchingIds.map(async id => {
          const aPrincipal: IPrincipal = await this._resolvePrincipal(id);
          return aPrincipal;
        }));
    }
    this.setState({ showForm: true, item: item });
  }

  private async _resolvePrincipal(userId: number): Promise<IPrincipal> {
    const userInfo: any = await sp.web.siteUserInfoList
      .items
      .getById(userId)
      .select("Id","EMail","Department","JobTitle","SipAddress","Title","Picture")
      .get();
    const userPrincipal: IPrincipal = {
      id: userInfo.Id,
      email: userInfo.EMail,
      department: userInfo.Department,
      jobTitle: userInfo.JobTitle,
      sip: userInfo.SipAddress,
      title: userInfo.Title,
      value: null,
      picture: userInfo.Picture ? userInfo.Picture.Url : null
    };
    return userPrincipal;
  }

  private _addItemAttachments = async (
    itemId: number,
    files: File[]
  ): Promise<void> => {
    let fileInfos: AttachmentFileInfo[] = [];

    for (let i: number = 0; i < files.length; i++) {
      const f: File = files[i];
      const fileInfo: AttachmentFileInfo = {
        name: f.name,
        content: await this._getFileArrayBuffer(f) as ArrayBuffer
      };
      fileInfos.push(fileInfo);
    }

    await sp.web.lists.getByTitle(ListTitles.ShrelloItems)
      .items
      .getById(itemId)
      .attachmentFiles
      .addMultiple(fileInfos);

    // tslint:disable-next-line:no-string-literal
    files.forEach(file => window.URL.revokeObjectURL(file["preview"]));
  }

  private _getFileArrayBuffer = (file: File): Promise<any> => {
    const reader: FileReader = new FileReader();

    return new Promise((resolve, reject) => {

      reader.onload = () => resolve(reader.result);

      reader.readAsArrayBuffer(file);
    });
  }

  private _updateShrelloItem = async (item: IShrelloItem, attachments?: File[]): Promise<void> => {
    const comment: string = !!attachments ? "Attached new file(s)." : null;
    const concatComments: string = this._concatComments(item, comment);
    const existingItem: IShrelloItem = await sp.web.lists.getByTitle(ListTitles.ShrelloItems).items
      .getById(item.Id)
      .expand("AttachmentFiles")
      .get<IShrelloItem>();

    item.TASCComments = concatComments;

    const uItem: IShrelloItem = this._compareItems(existingItem, item);
    const updateResult: ItemUpdateResult = await sp.web.lists.getByTitle(ListTitles.ShrelloItems)
      .items
      .getById(existingItem.Id)
      .update(uItem, "*", this._shrelloListEntityType);

    if (attachments) {
      await this._addItemAttachments(existingItem.Id, attachments);
    }

    const updatedItem: IShrelloItem = await updateResult.item.get<IShrelloItem>();

    this.setState({ updatedShrelloItem: updatedItem });
    return;
  }

  private _concatComments = (item: IShrelloItem, comment: string): string => {
    const comments: string = item.TASCComments
      ? `${item.TASCComments}${comment ? `
      ${comment}`: ``}`
      : `${comment}`;
    return comments;
  }

  private _compareItems(
    existingItem: any,
    compareItem: IShrelloItem
  ): IShrelloItem {
    let itemDiff: any = {
      Id: null,
      Title: null,
      PercentComplete: null,
      Status: null,
      Priority: null,
      Body: null,
      TASCId: null,
      TASCTypeCategory: null,
      TASCComments: null,
      RequesterId: null,
      Requester: null,
      AssignedToId: null,
      AssignedTo: null,
      WatchingId: null,
      Watching: null,
      PredecessorsId: null,
      SupportDepartmentId: null,
      SupportTopicId: null,
      StartDate: null,
      DueDate: null,
      ParentID: null,
      Attachments: null,
      AttachmentFiles: null
    };

    itemDiff.TASCComments= compareItem.TASCComments
      ? compareItem.TASCComments
      : "";

    for(const prop in itemDiff) {
      if (itemDiff.hasOwnProperty(prop)) {
        itemDiff[prop] = undefined;
        if(
          compareItem.hasOwnProperty(prop)
          && prop !== "TASCComments"
          && compareItem[prop] !== existingItem[prop]
        ) {
          if (prop === "AssignedToId" || prop === "WatchingId") {
            // tslint:disable-next-line:no-string-literal
            const eResults: number[] = existingItem[prop]
              // tslint:disable-next-line:no-string-literal
              ? existingItem[prop]["results"] || []
              : [];
            // tslint:disable-next-line:no-string-literal
            const cResults: number[] = compareItem[prop] && compareItem[prop]["results"]
              // tslint:disable-next-line:no-string-literal
              ? compareItem[prop]["results"]
              : compareItem[prop] || [];

            const combinedResults: number[] = Array.from(new Set<number>(eResults.concat(cResults)));

            if (combinedResults.length) {
              const results: {results: number[]} = {
                results: combinedResults
              };

              itemDiff[prop] = results;
              itemDiff.TASCComments = itemDiff.TASCComments
                ? `${itemDiff.TASCComments} | ${ prop } was changed.`
                : `${ prop } was changed.`;
            }
          } else if (prop === "Requester" || prop === "AssignedTo" || prop === "Watching") {
            // itemDiff[prop] = undefined;
          } else if (prop === "PredecessorsId") {
            // const results: any = { results: compareItem[prop].filter(r => existingItem[prop]["results"].indexOf(r) < 0) };
            // // tslint:disable-next-line:no-string-literal
            // if (!!results.results.length) {
            //   itemDiff[prop] = results;
            //   itemDiff.TASCComments += ` | ${ prop } was changed.`;
            // }
          } else if (prop === "AttachmentFiles") {
            // tslint:disable-next-line:no-string-literal
            // const existingFiles: any[] = existingItem[prop]["results"] || [];
            // const compareFiles: any[] = compareItem[prop] || [];

            // if (!existingFiles.length && !compareFiles.length) {
            //   null;
            // }

            // const addFiles: boolean = compareFiles.length
            //   ? compareFiles.filter
            // const results: any = { results: .filter(r =>
            //   existingItem[prop]["results"].findIndex(e => e.FileName === r.FileName) < 0) };
            // if (!!results.results.length) {
            //   itemDiff[prop] = results;
            //   itemDiff.TASCComments += `
            //   ${ prop } was changed.`;
            // }
          } else {
            itemDiff[prop] = compareItem[prop] || null;
            itemDiff.TASCComments = itemDiff.TASCComments
              ? `${itemDiff.TASCComments} | ${ prop } was changed.`
              : `${ prop } was changed.`;
          }
        }
      }
    }

    if (!itemDiff.TASCComments) {
      itemDiff.TASCComments = "No changes made.";
    }
    return itemDiff as IShrelloItem;
  }

}