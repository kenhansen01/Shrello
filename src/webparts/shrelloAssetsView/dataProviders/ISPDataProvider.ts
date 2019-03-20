import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { List, SPRest, SiteUserProps } from "@pnp/sp";

import { ISPList } from "../models";
import { IPrincipal } from "@pnp/spfx-controls-react/lib/common/SPEntities";

interface ISPDataProvider {
  sp: SPRest;
  list: List;
  items: any[];
  listProps: ISPList;
  webPartContext: IWebPartContext;
  currentUser: SiteUserProps;
  getList(listName: string): Promise<List>;
  getItems<T>(filter?:string): Promise<T[]>;
  // getItemAttachments(itemId: number): Promise<{Title: string, ServerRelativeUrl: string}[]>;
  addItemAttachments(itemId: number, files: File[]): Promise<void>;
  createItem<T>(newItem: T, attachments?: File[]): Promise<T[]>;
  updateItem<T>(itemUpdated: T, itemId: number, attachments?: File[]): Promise<T[]>;
  deleteItem<T>(itemDeleted: T): Promise<T[]>;
  resolvePrincipal(itemId: number): Promise<IPrincipal>;
}

export default ISPDataProvider;