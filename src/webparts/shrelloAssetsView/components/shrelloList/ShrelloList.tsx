import * as React from "react";
import {
  SiteUserProps,
  PagedItemCollection
} from "@pnp/sp";
import { IconButton } from "office-ui-fabric-react";
import { IconNames } from "@uifabric/icons";

import { IShrelloItem, ISupportDepartmentItem } from "../../models";
import {
  IShrelloListProps,
  IShrelloListState,
  IShrelloItemWithChildren,
  viewName,
  viewData,
  IViewObject
} from "./IShrelloList";
import ShrelloLane, { IShrelloLaneProps } from "./ShrelloLane";
import styles from "./ShrelloList.module.scss";

function ShrelloLaneContainer(props: any): JSX.Element {
  return(
    <div className={ styles.shrelloLaneZone }>
      <div className={ styles.shrelloLaneBoard }>
        { props.children }
      </div>
      { props.pagedShrelloItems && props.pagedShrelloItems.hasNext &&
        <IconButton
          title="See more cards."
          iconProps={{ iconName: IconNames.CircleAdditionSolid, className: styles.moreResultsButton }}
          className={ styles.moreResults }
        />
      }
    </div>
  );
}

function ShrelloLanes(props: any): JSX.Element {
  const laneName: string = props.selectedView.menuProps.name;
  let laneProps: IShrelloLaneProps[] = props.selectedView.lanes;
  let shrelloLanes: JSX.Element[] = [];
  let workingBoard: boolean = true;

  const defaultTestStrings: {} = {
    "Assigned" : ["Not Started", "Initiated", "Assigned"],
    "Planned" : ["Planned"],
    "In Progress" : ["In Progress", "On Hold"],
    "Ready for Review" : ["Ready for Review"],
    "Complete" : ["Completed", "Canceled"]
  };

  let testStrings: {} = defaultTestStrings;

  switch (laneName) {
    case viewName.allByTeam:
      props.supportDepartments.map( (dept: ISupportDepartmentItem) => {
        const teamLaneProps: IShrelloLaneProps = {
          title: dept.Title,
          laneValue: dept.Id.toString(),
          itemColumn: "SupportDepartmentId",
          laneDepartment: dept
        };
        if (laneProps.findIndex(lp => lp.title === teamLaneProps.title) === -1) {
          laneProps.push(teamLaneProps);
        }
      });
      workingBoard = false;
      break;
    case viewName.myRequests:
      // tslint:disable-next-line:no-string-literal
      testStrings["In Progress"] = testStrings["In Progress"].concat(testStrings["Planned"], testStrings["Ready for Review"]);
      workingBoard = true;
      break;
    case viewName.myTASCs:
      testStrings = defaultTestStrings;
      workingBoard = true;
      break;
    case viewName.teamShrellos:
      testStrings = defaultTestStrings;
      workingBoard = true;
      break;
    default:
      workingBoard = false;
      break;
  }

  shrelloLanes = laneProps.map(lane => {
    if (lane.title === "I need help with...") {
      return <ShrelloLane
        title={ lane.title }
      />;
    }
    if (props.items && props.items.length) {
      lane.items = workingBoard
        ? props.items.filter(item => {
          const lastModified = new Date(item.Modified);
          const now = new Date();
          if (!item[lane.itemColumn]) {
            return lane.laneValue === "Assigned" && item;
          } else if (lane.laneValue === "Complete") {
            return (lastModified.getDate() > (now.getDate() - 1)
              && testStrings[lane.laneValue].findIndex(lv => lv === item[lane.itemColumn]) > -1
              && item);
          } else {
            return testStrings[lane.laneValue].findIndex(lv => lv === item[lane.itemColumn]) > -1 && item;
          }
        })
        : props.items.filter(item =>
            lane.laneValue === String(item[lane.itemColumn]) && item);
    }
    return <ShrelloLane
      {...props}
      {...lane}
    />;
  });

  return(
    <ShrelloLaneContainer {...props}>
      { shrelloLanes }
    </ShrelloLaneContainer>
  );
}

export default class ShrelloList extends React.Component<IShrelloListProps, IShrelloListState> {
  private _pagedItems: PagedItemCollection<any>;
  private _items: IShrelloItemWithChildren[];
  private _helpItems: IShrelloItemWithChildren[];
  private _currentUser: SiteUserProps;
  private _selectedView: IViewObject;
  // private _viewFilter: string;

  constructor(props: IShrelloListProps) {
    super(props);
    this._selectedView = props.selectedView ||
    viewData.find(view => view.menuProps.name === viewName.myRequests);

    this.state = {
      items: props.items || [],
      shrelloItems: props.shrelloItems || [],
      supportDepartments: props.supportDepartments || [],
      supportTopics: props.supportTopics || [],
      currentUser: props.currentUser || null,
      selectedView: this._selectedView
    };
    this._handleItemClick = this._handleItemClick.bind(this);
    this._handleSupportTeamSelect = this._handleSupportTeamSelect.bind(this);
  }

  public async componentWillReceiveProps(props: IShrelloListProps): Promise<void> {
    let selectedView: IViewObject = props.selectedView
      || viewData.find(view =>
        view.menuProps.name === viewName.myRequests);
    this._pagedItems = props.pagedShrelloItems;

    if(this._pagedItems) {
      this._items = this._pagedItems
      .results
      .map(item => {
        if(item.AttachmentFiles.results && item.AttachmentFiles.results.length) {
          item.AttachmentFiles = item.AttachmentFiles.results;
        }
        return item as IShrelloItem;
      });
    }

    this.setState({
      items: this._items,
      shrelloItems: props.shrelloItems || [],
      pagedShrelloItems: this._pagedItems || undefined,
      supportDepartments: props.supportDepartments || [],
      supportTopics: props.supportTopics || [],
      currentUser: props.currentUser || null,
      selectedView: selectedView
    });
  }

  public async componentDidMount(): Promise<void> {
    this._currentUser = this.props.currentUser;
    this._pagedItems = this.props.pagedShrelloItems;

    if(this._pagedItems) {
      this._items = this._pagedItems
      .results
      .map(item => {
        if(item.AttachmentFiles.results && item.AttachmentFiles.results.length) {
          item.AttachmentFiles = item.AttachmentFiles.results;
        }
        return item as IShrelloItem;
      });
    }

    this.setState({
      items: this._items,
      pagedShrelloItems: this._pagedItems || undefined,
      shrelloItems: this.props.shrelloItems || [],
      supportDepartments: this.props.supportDepartments || [],
      supportTopics: this.props.supportTopics || [],
      currentUser: this._currentUser
    });
  }

  public render(): JSX.Element {
    return <ShrelloLanes
      {...this.state}
      {...this.props}
    />;
  }

  private _handleItemClick = (ev: React.MouseEvent<HTMLElement>, item: IShrelloItem): void => {
    this.props.handleItemClick(ev, item);
  }

  private _handleSupportTeamSelect = (ev: React.MouseEvent<HTMLElement>, laneDepartment: ISupportDepartmentItem): void => {
    this.props.onSelectTeamView(ev, laneDepartment);
  }
}