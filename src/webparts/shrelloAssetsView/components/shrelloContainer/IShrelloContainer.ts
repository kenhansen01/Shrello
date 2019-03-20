import { DisplayMode } from "@microsoft/sp-core-library";
import { SPRest, SiteUserProps, PagedItemCollection } from "@pnp/sp";
import { ICommandBarItemProps } from "office-ui-fabric-react/lib/CommandBar";

import { IWebPartContext } from "@microsoft/sp-webpart-base";
import {
  IShrelloItem,
  ISupportDepartmentItem,
  ISupportTopicItem
} from "../../models";
import { IViewObject } from "../shrelloList";

interface IShrelloContainerProps {
  sp: SPRest;
  context: IWebPartContext;
  contentArea: Element;
  currentUser: SiteUserProps;
  webPartDisplayMode: DisplayMode;
  configureStartCallback: () => void;
}

interface IShrelloContainerState {
  shrelloHeading: string;
  shrelloItems: IShrelloItem[];
  pagedShrelloItems?: PagedItemCollection<any>;
  supportDepartmentItems: ISupportDepartmentItem[];
  supportTopicItems: ISupportTopicItem[];
  showForm: boolean;
  fullScreen: boolean;
  shrelloCommands?: ICommandBarItemProps[];
  shrelloFarCommands?: ICommandBarItemProps[];
  contentArea: Element;
  currentUser: SiteUserProps;
  item?: IShrelloItem;
  newShrelloItem?: IShrelloItem;
  updatedShrelloItem?: IShrelloItem;
  selectedView?: IViewObject;
  selectedSupportDepartment?: ISupportDepartmentItem;
  shrelloListEntityType?: string;
}

export {
  IShrelloContainerProps,
  IShrelloContainerState
};