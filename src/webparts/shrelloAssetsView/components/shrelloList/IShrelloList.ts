import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { SiteUserProps, PagedItemCollection } from "@pnp/sp";
import { IconNames } from "@uifabric/icons";
import { IContextualMenuItem } from "office-ui-fabric-react";
import { IShrelloLaneProps } from "./ShrelloLane";
import {
  ISupportDepartmentItem,
  IShrelloItem,
  ISupportTopicItem
} from "../../models";

export enum viewName {
  allByTeam = "TASCs by Team",
  teamShrellos = "Team TASCs",
  myRequests = "My Requests",
  myTASCs = "My TASCs",
  byProject = "By Project",
  byDueDate = "By Due Date",
  byPriority = "By Priority",
  dashboard = "Dashboard"
}

export interface IViewContextualMenuItem extends IContextualMenuItem {
  name: viewName;
  filter?: (id?: number) => string;
}

export interface IViewObject {
  menuProps: IViewContextualMenuItem;
  lanes: IShrelloLaneProps[];
}

export const viewData: IViewObject[] = [
  {
    menuProps: {
      key: "allByTeam",
      name: viewName.allByTeam,
      filter: () => "",
      iconProps: {
        iconName: IconNames.ViewAll
      },
      ["data-automation-id"]: "allByTeamView",
      onClick: null
    },
    lanes: []
  },
  {
    menuProps: {
      key: "myRequests",
      name: viewName.myRequests,
      filter: (id: number) => `Requester eq ${id}`,
      iconProps: {
        iconName: IconNames.Unknown
      },
      ["data-automation-id"]: "myRequestsView",
      onClick: null
    },
    lanes: [
      {
        title: "I need help with..."
      },
      {
        title: "Submitted",
        laneValue: "Assigned",
        itemColumn: "Status"
      },
      {
        title: "In Progress",
        laneValue: "In Progress",
        itemColumn: "Status"
      },
      {
        title: "Complete",
        laneValue: "Complete",
        itemColumn: "Status"
      }
    ]
  },
  {
    menuProps: {
      key: "myShrellos",
      name: viewName.myTASCs,
      filter: (id: number) => `AssignedTo eq ${id}`,
      iconProps: {
        iconName: IconNames.TaskLogo
      },
      ["data-automation-id"]: "myShrellosView",
      onClick: null
    },
    lanes: [
      {
        title: "Backlog",
        laneValue: "Assigned",
        itemColumn: "Status"
      },
      {
        title: "Current Sprint",
        laneValue: "Planned",
        itemColumn: "Status"
      },
      {
        title: "In Progress",
        laneValue: "In Progress",
        itemColumn: "Status"
      },
      {
        title: "Ready for Approval",
        laneValue: "Ready for Review",
        itemColumn: "Status"
      },
      {
        title: "Complete",
        laneValue: "Complete",
        itemColumn: "Status"
      }
    ]
  },
  {
    menuProps: {
      key: "teamShrellos",
      name: viewName.teamShrellos,
      filter: (id: number) => `SupportDepartment eq ${id}`,
      iconProps: {
        iconName: IconNames.Teamwork
      },
      ["data-automation-id"]: "myShrellosView",
      onClick: null
    },
    lanes: [
      {
        title: "Backlog",
        laneValue: "Assigned",
        itemColumn: "Status"
      },
      {
        title: "Current Sprint",
        laneValue: "Planned",
        itemColumn: "Status"
      },
      {
        title: "In Progress",
        laneValue: "In Progress",
        itemColumn: "Status"
      },
      {
        title: "Ready for Approval",
        laneValue: "Ready for Review",
        itemColumn: "Status"
      },
      {
        title: "Complete",
        laneValue: "Complete",
        itemColumn: "Status"
      }
    ]
  },
];

export interface IShrelloItemWithChildren extends IShrelloItem {
  children?: IShrelloItemWithChildren[];
}

export interface IShrelloListProps {
  shrelloItems: IShrelloItem[];
  pagedShrelloItems?: PagedItemCollection<any>;
  supportDepartments: ISupportDepartmentItem[];
  supportTopics: ISupportTopicItem[];
  currentUser: SiteUserProps;
  context?: IWebPartContext;
  selectedView?: IViewObject;
  items?: IShrelloItem[];
  newShrelloItem?: IShrelloItem;
  updatedShrelloItem?: IShrelloItem;
  handleItemClick?: (ev: React.MouseEvent<HTMLElement>, item: IShrelloItem) => void;
  onSelectTeamView?: (ev: React.MouseEvent<HTMLElement>, laneDepartment: ISupportDepartmentItem) => void;
  getShrelloItems?: (
    filter?: string,
    orderby?: {prop:string, ascending?: boolean}
  ) => Promise<PagedItemCollection<any>>;
}

export interface IShrelloListState {
  shrelloItems?: IShrelloItem[];
  supportDepartments: ISupportDepartmentItem[];
  supportTopics: ISupportTopicItem[];
  currentUser: SiteUserProps;
  pagedShrelloItems?: PagedItemCollection<any>;
  items?: IShrelloItemWithChildren[];
  helpItems?: IShrelloItemWithChildren[];
  selectedView?: IViewObject;
  lanes?: JSX.Element[];
}