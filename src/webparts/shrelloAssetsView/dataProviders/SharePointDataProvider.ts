// import {
//   sp,
//   List,
//   ItemAddResult,
//   ItemUpdateResult,
//   SPRest,
//   SiteUserProps
// } from "@pnp/sp";
// import {
//   AttachmentFiles,
//   AttachmentFileInfo
// } from "@pnp/sp/src/attachmentfiles";
// import { SPBatch } from "@pnp/sp/src/batch";
// import { IWebPartContext } from "@microsoft/sp-webpart-base";

// import ISPDataProvider from "./ISPDataProvider";
// import { ISPList } from "../models";

// import { ProvisionLists } from "../../../assets/lists/ProvisionLists";
// import { PnPClientStorage } from "@pnp/common";
// import { IPrincipal } from "@pnp/spfx-controls-react/lib/common/SPEntities";

// export interface IWithAttachments {
//   Attachments?: boolean;
//   AttachedFiles?: AttachmentFiles;
//   AttachmentFiles?: any;
// }
// export default class SharePointDataProvider implements ISPDataProvider {

//   private _sp: SPRest;
//   private _list: List;
//   private _items: any[];
//   private _listProps: ISPList;
//   private _webPartContext: IWebPartContext;
//   private _currentUser: SiteUserProps;
//   private _pLists: ProvisionLists;

//   public set sp(value: SPRest) { this._sp = value; }
//   public get sp(): SPRest { return this._sp; }

//   public set list(value: List) { this._list = value; }
//   public get list(): List { return this._list; }

//   public set items(value: any[]) { this._items = value; }
//   public get items(): any[] { return this._items; }

//   public set listProps(value: ISPList) { this._listProps = value; }
//   public get listProps(): ISPList { return this._listProps; }

//   public set currentUser(value: SiteUserProps) { this._currentUser = value; }
//   public get currentUser(): SiteUserProps { return this._currentUser; }

//   public set webPartContext(value: IWebPartContext) { this._webPartContext = value; }
//   public get webPartContext(): IWebPartContext { return this._webPartContext; }

//   public async getList(listName: string): Promise<List> {
//     let list: List;
//     list = await sp.web.lists.getByTitle(listName);
//     // check for existence by getting title, on err provision the lists
//     await list.select("Title").get()
//       .catch(err => {
//         list = this._provisionAssets();
//       });

//     return list;
//   }

//   public async getItems<T>(filter?: string): Promise<T[]> {
//     return await this._getItems<T>(filter);
//   }

//   public async getCurrentUser(): Promise<SiteUserProps> {
//     return await this._sp.web.currentUser.get<SiteUserProps>();
//   }

//   public async addItemAttachments(itemId: number, files: File[]): Promise<void> {
//     return await this._addItemAttachments(itemId, files);
//   }

//   public async createItem<T>(
//     item: T,
//     attachments?: File[]
//   ): Promise<T[]> {
//     const batch: SPBatch = sp.web.createBatch();
//     const entityType: string = await this._list.getListItemEntityTypeFullName();

//     const batchPromises: Promise<{}>[] = [
//       this._createItem(batch, entityType, item, attachments),
//       this._getItemsBatched(batch)
//     ];

//     return this._resolveBatch<T>(batch, batchPromises);
//   }

//   public async updateItem<T>(
//     itemUpdated: T,
//     itemId: number,
//     attachments?: File[]
//   ): Promise<T[]> {
//     const batch: SPBatch = sp.web.createBatch();
//     const entityType: string = await this._list.getListItemEntityTypeFullName();

//     const batchPromises: Promise<{}>[] = [
//       this._updateItem(batch, entityType, itemUpdated, itemId, attachments),
//       this._getItemsBatched(batch)
//     ];

//     return this._resolveBatch<T>(batch, batchPromises);
//   }

//   public async deleteItem<T>(itemDeleted: (T & {Id:number})): Promise<T[]> {
//     const batch: SPBatch = sp.web.createBatch();

//     const batchPromises: Promise<{}>[] = [
//       this._deleteItem(batch, itemDeleted),
//       this._getItemsBatched(batch)
//     ];

//     return this._resolveBatch<T>(batch, batchPromises);
//   }

//   public async resolvePrincipal(userId: number): Promise<IPrincipal> {
//     return await this._resolvePrincipal(userId);
//   }

//   private async _getItems<T>(filter?: string): Promise<T[]> {
//     let items: any[] = [];
//     if (!filter) {
//       items = await this._list.items.expand("AttachmentFiles").getAll();
//     } else {
//       items = await this._list.items.expand("AttachmentFiles").filter(filter).getAll();
//     }
//     const tItems: T[] = items.map((item: T) => item);
//     return tItems;
//   }

//   private async _addItemAttachments(
//     itemId: number,
//     files: File[]
//   ): Promise<void> {
//     let fileInfos: AttachmentFileInfo[] = [];

//     for (let i: number = 0; i < files.length; i++) {
//       const f: File = files[i];
//       const fileInfo: AttachmentFileInfo = {
//         name: f.name,
//         content: await this._getFileArrayBuffer(f) as ArrayBuffer
//       };
//       fileInfos.push(fileInfo);
//     }

//     await this._list.items.getById(itemId)
//       .attachmentFiles.addMultiple(fileInfos);

//     // tslint:disable-next-line:no-string-literal
//     files.forEach(file => window.URL.revokeObjectURL(file["preview"]));
//   }

//   private _getFileArrayBuffer = (file: File): Promise<any> => {
//     const reader: FileReader = new FileReader();

//     return new Promise((resolve, reject) => {

//       reader.onload = () => resolve(reader.result);

//       reader.readAsArrayBuffer(file);
//     });
//   }

//   private async _getItemsBatched<T>(batch: SPBatch): Promise<T[]> {
//     let items: any[] = await this._list.items.expand("AttachmentFiles").inBatch(batch).getAll();
//     const tItems: T[] = items.map((item: T) => item);
//     return tItems;
//   }

//   private async _createItem<T>(
//     batch: SPBatch,
//     entityType: string,
//     item: T,
//     attachments?: File[]
//   ): Promise<ItemAddResult> {
//     const addItemResult: ItemAddResult = await this._list.items.inBatch(batch).add(item, entityType);
//     if (attachments) {
//       const itemId: number = (await addItemResult.item.select("Id").get()).Id;
//       await this._addItemAttachments(itemId, attachments);
//     }
//     return addItemResult;
//   }

//   private async _updateItem<T>(
//     batch: SPBatch,
//     entityType: string,
//     item: T,
//     itemId: number,
//     attachments?: File[]
//   ): Promise<ItemUpdateResult> {

//     if (attachments) {
//       // tslint:disable-next-line:no-string-literal
//       item["TASCComments"] = `${item["TASCComments"]} | Attachment(s) added.`;
//     }

//     const updateResult: ItemUpdateResult = await this._list.items.getById(itemId).inBatch(batch).update(item, "*", entityType);

//     if (attachments) {
//       await this._addItemAttachments(itemId, attachments);
//     }

//     return updateResult;
//   }

//   private async _deleteItem<T>(batch: SPBatch, item: (T & {Id:number})): Promise<{}> {
//     return await this._list.items.getById(item.Id).delete().then(async _ => await {});
//   }

//   private async _resolveBatch<T>(batch: SPBatch, promises: Promise<{}>[]): Promise<T[]> {
//     return await batch.execute()
//       .then(async _ => {
//         const pValues: {}[] = await Promise.all(promises);
//         return pValues[pValues.length -1] as T[];
//       });
//   }

//   private async _resolvePrincipal(userId: number): Promise<IPrincipal> {
//     const userInfo: any = await this._sp.web.lists.getByTitle("User Information List")
//     .items.getById(userId).select("Id","EMail","Department","JobTitle","SipAddress","Title","Picture").get();
//     const userPrincipal: IPrincipal = {
//       id: userInfo.Id,
//       email: userInfo.EMail,
//       department: userInfo.Department,
//       jobTitle: userInfo.JobTitle,
//       sip: userInfo.SipAddress,
//       title: userInfo.Title,
//       value: null,
//       picture: userInfo.Picture ? userInfo.Picture.Url : null
//     };
//     return userPrincipal;
//   }

//   private _provisionAssets(): List {
//     this._pLists = new ProvisionLists(this._sp);
//     return this._pLists.ShrelloList;
//   }
// }