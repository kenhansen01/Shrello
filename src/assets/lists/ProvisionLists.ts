// import * as core from "core-js";
import {
  List,
  SPRest,
  FieldTypes
} from "@pnp/sp";
import {
  Logger,
  ConsoleListener,
  LogLevel
} from "@pnp/logging";

import SupportDepartmentList from "./supportdepartment/SupportDepartmentList";
import SupportTopicList from "./supporttopic/SupportTopicList";
import ShrelloList from "./shrello/ShrelloList";
import { ListUtility, IListDef } from "../../utilities/ListUtility";
import { ILookupFieldDef } from "../fields/IFieldDefs";
import { View } from "@pnp/sp/src/views";

interface IViewFieldDef {
  fIName: string;
  vIndex: number;
}

export class ProvisionLists {
  public SupportDepartmentList: List;
  public SupportTopicList: List;
  public ShrelloList: List;

  private _sp: SPRest;
  private _listUtility: ListUtility;
  private _supportDepartmentLookupField: ILookupFieldDef;
  private _supportTopicLookupField: ILookupFieldDef;
  private _taskListDefaultView: IViewFieldDef[] = [
    { fIName:"LinkTitle", vIndex: 0 },
    { fIName:"Priority", vIndex: 1 },
    { fIName:"DueDate", vIndex: 2 },
    { fIName:"TASCTypeCategory", vIndex: 3 },
    { fIName:"Body", vIndex: 4 },
    { fIName:"TASCId", vIndex: 5 },
    { fIName:"SupportDepartment", vIndex: 6 },
    { fIName:"SupportTopic", vIndex: 7 },
    { fIName:"AssignedTo", vIndex: 8 },
    { fIName:"Status", vIndex: 9 },
    { fIName:"PercentComplete", vIndex: 10 },
    { fIName:"WorkHoursEstimate", vIndex: 11 },
    { fIName:"Requester", vIndex: 12 },
    { fIName:"Watching", vIndex: 13 },
    { fIName:"Modified", vIndex: 14 }
  ];

  constructor(sp: SPRest) {
    this._sp = sp;
    this._supportDepartmentLookupField = {
      title: "Support Department",
      fieldType: FieldTypes.Lookup,
      fieldSchema: "",
      lookupListId: "",
      lookupFieldName: "Title",
      viewIndex: 4,
      properties: {
        InternalName: "SupportDepartment",
        Required: true
      }
    };

    this._supportTopicLookupField = {
      title: "Support Topic",
      fieldType: FieldTypes.Lookup,
      fieldSchema: "",
      lookupListId: "",
      lookupFieldName: "Title",
      viewIndex: 5,
      properties: {
        InternalName: "SupportTopic",
        Required: true
      }
    };

    Logger.subscribe(new ConsoleListener());
    Logger.activeLogLevel = LogLevel.Info;

    this._listUtility = new ListUtility(sp);
    this.ensureLists().catch(err => Logger.write(err));
  }

  private async ensureLists(): Promise<void> {
    let dId: string;
    let tId: string;
    let tView: View;
    this.SupportDepartmentList = await this.ensureList(SupportDepartmentList);
    this.SupportTopicList = await this.ensureList(SupportTopicList);
    this.ShrelloList = await this.ensureList(ShrelloList);
    // store Support List ID
    dId = (await this.SupportDepartmentList.select("Id").get()).Id as string;
    tId = (await this.SupportTopicList.select("Id").get()).Id as string;
    this._supportDepartmentLookupField.lookupListId = dId;
    this._supportTopicLookupField.lookupListId = tId;

    await this._listUtility.ensureListField(this.SupportTopicList, this._supportDepartmentLookupField)
      .then(async _ => {
        let sFields: IViewFieldDef[] = SupportTopicList.fieldDefs.map((fDef, idx) => {
          return {
            fIName: fDef.properties.InternalName === "Title" ? "LinkTitle" : fDef.properties.InternalName,
            vIndex: idx + 1
          };
        }).concat([{fIName: this._supportDepartmentLookupField.properties.InternalName, vIndex: 0}]);
        await this.setView(this.SupportTopicList.defaultView, sFields);
      });
    await this._listUtility.ensureListField(this.ShrelloList, this._supportDepartmentLookupField)
      .then(async _ => await this._listUtility.ensureListField(this.ShrelloList, this._supportTopicLookupField))
      .then(async _ => await this.setView(this.ShrelloList.defaultView, this._taskListDefaultView));
  }

  public async ensureList(lDef: IListDef): Promise<List> {
    // get or create the list
    const list: List = await this._listUtility.ensureList(lDef);
    const allItemsView: View = await list.views.getByTitle(lDef.defaultViewName);
    let internalName: string;

    // get or create fields on list
    // for...of forces synchronous execution preventing 409
    for (let fDef of lDef.fieldDefs) {
      internalName = fDef.properties.InternalName === "Title"
        ? "LinkTitle"
        : fDef.properties.InternalName;
      await this._listUtility.ensureListField(list, fDef);
    }

    // the below would be better, but SP returns 409 errors for concurrent updates
    // const fieldPromises = lDef.fieldDefs.map(async fDef => {
    //   const lfield: Field = await this._listUtility.ensureListField(list, fDef);
    //   return lfield;
    // });
    // for (const fieldPromise of fieldPromises) {
    //   logger.writeJSON(await fieldPromise);
    // }

    return list;
  }

  public async setView(view: View, fieldDefs: {fIName: string, vIndex: number}[]): Promise<void> {
    view.fields.removeAll();
    for(let field of fieldDefs) {
      await view.fields.add(field.fIName);
    }
    for(let f of fieldDefs) {
      await view.fields.move(f.fIName, f.vIndex);
    }
  }
}