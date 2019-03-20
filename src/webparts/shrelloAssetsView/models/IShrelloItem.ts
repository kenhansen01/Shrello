import { IPrincipal } from "@pnp/spfx-controls-react/lib/common/SPEntities";

interface IShrelloItem {
  Id?: number;
  Title?: string;
  PercentComplete?: number;
  Status?: string;
  Priority?: string;
  Body?: string;
  TASCId?: string;
  TASCTypeCategory?: string;
  TASCComments?: string;
  RequesterId?: number;
  Requester?: IPrincipal;
  AssignedToId?: number[];
  AssignedTo?: IPrincipal[];
  WatchingId?: number[];
  Watching?: IPrincipal[];
  PredecessorsId?: number[];
  SupportDepartmentId?: number;
  SupportTopicId?: number;
  StartDate?: Date;
  DueDate?: Date;
  ParentID?: number;
  Attachments?: boolean;
  AttachmentFiles?: {FileName: string, ServerRelativeUrl: string}[];
}

export default IShrelloItem;