// import * as core from "core-js";
import {
  Web,
  SPRest,
  List,
  ListEnsureResult,
  FieldTypes,
  DateTimeFieldFormatType,
  CalendarType,
  Field,
  ChoiceFieldFormatType,
  UrlFieldFormatType,
  FieldUpdateResult,
  FieldCreationProperties,
  FieldAddResult
} from "@pnp/sp";
import { TypedHash } from "@pnp/common";

import {
  IBooleanFieldDef,
  ICalcFieldDef,
  IChoiceFieldDef,
  ICurrencyFieldDef,
  ICustomFieldDef,
  IDateTimeFieldDef,
  IFieldDef,
  ILookupFieldDef,
  IMultiChoiceFieldDef,
  IMultiLineTextFieldDef,
  INumberFieldDef,
  ITextFieldDef,
  IUrlFieldDef,
  IUserFieldDef
} from "../assets/fields/IFieldDefs";
import { Logger } from "@pnp/logging";

export interface IListDef {
  title: string;
  description?: string;
  template?: number;
  enableContentTypes?: boolean;
  additionalSettings?: TypedHash<string | number | boolean> & {
    EnableAttachments?: boolean,
    EnableFolderCreation?: boolean,
    EnableVersioning?: boolean,
    EnableMinorVersions?: boolean,
    ForceCheckout?: boolean,
    Hidden?: boolean,
    NoCrawl?: boolean,
    OnQuickLaunch?: boolean
  };
  fieldDefs?: (
    IBooleanFieldDef |
    ICalcFieldDef |
    IChoiceFieldDef |
    ICurrencyFieldDef |
    ICustomFieldDef |
    IDateTimeFieldDef |
    ILookupFieldDef |
    IMultiChoiceFieldDef |
    IMultiLineTextFieldDef |
    INumberFieldDef |
    ITextFieldDef |
    IUrlFieldDef |
    IUserFieldDef
  )[];
  defaultViewName?: string;
}

export class ListUtility {
  private web: Web;

  constructor(sp: SPRest) {
    this.web = sp.web;
  }

  private fldType<T>(fld: T): T { return fld; }

  public async ensureList(listDef: IListDef): Promise<List> {
    let ler: ListEnsureResult = await this.web.lists.ensure(
      listDef.title,
      listDef.description,
      listDef.template,
      listDef.enableContentTypes,
      listDef.additionalSettings);
    return ler.list;
  }

  public async ensureListField(
    list: List,
    fieldDef: IFieldDef |
      IBooleanFieldDef |
      ICalcFieldDef |
      IChoiceFieldDef |
      ICurrencyFieldDef |
      ICustomFieldDef |
      IDateTimeFieldDef |
      ILookupFieldDef |
      IMultiChoiceFieldDef |
      IMultiLineTextFieldDef |
      INumberFieldDef |
      ITextFieldDef |
      IUrlFieldDef |
      IUserFieldDef
  ): Promise<void> {

    let field: Field;
    let updateField: boolean = false;

    // first try to get the field
    try {
      if (fieldDef.properties && fieldDef.properties.InternalName) {
        field = list.fields.getByInternalNameOrTitle(<string>fieldDef.properties.InternalName);
      } else {
        field = list.fields.getByTitle(<string>fieldDef.title);
      }
      await field.get();
      updateField = true;
    }
    // if getting field fails, create it
    catch (err) {
      Logger.writeJSON(err);
      await this.AddField(list, fieldDef);
    }
    if(updateField) {
      await this.UpdateField(field, fieldDef);
      updateField = false;
    }
  }
  /**
   * AddField - Creates a new field, takes any of the field definitions from IFieldDefs
   * 
   * @param list - the list that will get the field.
   * 
   * @param fieldDef - the object with the field info to be updated.
   */
  public async AddField(
    list: List,
    fieldDef: IFieldDef |
      IBooleanFieldDef |
      ICalcFieldDef |
      IChoiceFieldDef |
      ICurrencyFieldDef |
      ICustomFieldDef |
      IDateTimeFieldDef |
      ILookupFieldDef |
      IMultiChoiceFieldDef |
      IMultiLineTextFieldDef |
      INumberFieldDef |
      ITextFieldDef |
      IUrlFieldDef |
      IUserFieldDef
  ): Promise<void> {
    
    let fAddResult: Promise<FieldAddResult>;
    let field: Field;

    try {
      switch (fieldDef.fieldType) {
        case FieldTypes.Boolean:
          var fdBool: IBooleanFieldDef = this.fldType<IBooleanFieldDef>(fieldDef as IBooleanFieldDef);
          fAddResult = list.fields.addBoolean(
            fdBool.properties.InternalName,
            fdBool.properties
          );
          break;
        case FieldTypes.Calculated:
          var fdCalc: ICalcFieldDef = this.fldType<ICalcFieldDef>(fieldDef as ICalcFieldDef);
          fAddResult = list.fields.addCalculated(
            fdCalc.properties.InternalName,
            fdCalc.formula,
            fdCalc.dateFormat,
            fdCalc.outputType,
            fdCalc.properties
          );
          break;
        case FieldTypes.Choice:
          var fdChoice: IChoiceFieldDef = this.fldType<IChoiceFieldDef>(fieldDef as IChoiceFieldDef);
          fAddResult = list.fields.addChoice(
            fdChoice.properties.InternalName,
            fdChoice.choices,
            fdChoice.format,
            fdChoice.fillIn,
            fdChoice.properties
          );
          break;
        case FieldTypes.Currency:
          var fdCurrency: ICurrencyFieldDef = this.fldType<ICurrencyFieldDef>(fieldDef as ICurrencyFieldDef);
          fAddResult = list.fields.addCurrency(
            fdCurrency.properties.InternalName,
            fdCurrency.minValue,
            fdCurrency.maxValue,
            fdCurrency.currencyLocalId,
            fdCurrency.properties
          );
          break;
        case FieldTypes.DateTime:
          var fdDate: IDateTimeFieldDef = this.fldType<IDateTimeFieldDef>(fieldDef as IDateTimeFieldDef);
          fAddResult = list.fields.addDateTime(
            fdDate.properties.InternalName,
            fdDate.displayFormat,
            fdDate.calendarType,
            fdDate.friendlyDisplayFormat,
            fdDate.properties
          );
          break;
        case FieldTypes.Lookup:
          var fdLookup: ILookupFieldDef = this.fldType<ILookupFieldDef>(fieldDef as ILookupFieldDef);
          fAddResult = list.fields.addLookup(
            fdLookup.properties.InternalName,
            fdLookup.lookupListId,
            fdLookup.lookupFieldName
          );
          break;
        case FieldTypes.MultiChoice:
          var fdMChoice: IMultiChoiceFieldDef =
            this.fldType<IMultiChoiceFieldDef>(fieldDef as IMultiChoiceFieldDef);
          fAddResult = list.fields.addMultiChoice(
            fdMChoice.properties.InternalName,
            fdMChoice.choices,
            fdMChoice.fillIn,
            fdMChoice.properties
          );
          break;
        case FieldTypes.Note:
          var fdMultiLine: IMultiLineTextFieldDef =
            this.fldType<IMultiLineTextFieldDef>(fieldDef as IMultiLineTextFieldDef);
          fAddResult = list.fields.addMultilineText(
            fdMultiLine.properties.InternalName,
            fdMultiLine.numberOfLines,
            fdMultiLine.richText,
            fdMultiLine.restrictedMode,
            fdMultiLine.appendOnly,
            fdMultiLine.allowHyperLink,
            fdMultiLine.properties
          );
          break;
        case FieldTypes.Number:
          var fdNumber: INumberFieldDef = this.fldType<INumberFieldDef>(fieldDef as INumberFieldDef);
          fAddResult = list.fields.addNumber(
            fdNumber.properties.InternalName,
            fdNumber.minValue,
            fdNumber.maxValue,
            fdNumber.properties
          );
          break;
        case FieldTypes.Text:
          var fdText: ITextFieldDef = this.fldType<ITextFieldDef>(fieldDef as ITextFieldDef);
          fAddResult = list.fields.addText(
            fdText.properties.InternalName,
            fdText.maxLength,
            fdText.properties
          );
          break;
        case FieldTypes.URL:
          var fdURL: IUrlFieldDef = this.fldType<IUrlFieldDef>(fieldDef as IUrlFieldDef);
          fAddResult = list.fields.addUrl(
            fdURL.properties.InternalName,
            fdURL.displayFormat,
            fdURL.properties
          );
          break;
        case FieldTypes.User:
          var fdUser: IUserFieldDef = this.fldType<IUserFieldDef>(fieldDef as IUserFieldDef);
          fAddResult = list.fields.addUser(
            fdUser.properties.InternalName,
            fdUser.selectionMode,
            fdUser.properties
          );
          break;
        default:
          var fdCustomField: ICustomFieldDef = this.fldType<ICustomFieldDef>(fieldDef as ICustomFieldDef);
          fAddResult = list.fields.add(
            fdCustomField.properties.InternalName,
            fdCustomField.fieldType,
            fdCustomField.properties
          );
          break;
      }
      field = (await fAddResult).field;
      await this.UpdateField(field, fieldDef);
    } catch (err) {
      Logger.writeJSON(err);
    }
  }

  /**
   * UpdateField - Updates existing field, takes any of the field definitions from IFieldDefs
   * 
   * @param field - the field to be updated
   *
   * @param fieldDef - the object with the field info to be updated.
   */
  public async UpdateField(
    field: Field,
    fieldDef: IFieldDef |
      IBooleanFieldDef |
      ICalcFieldDef |
      IChoiceFieldDef |
      ICurrencyFieldDef |
      ICustomFieldDef |
      IDateTimeFieldDef |
      ILookupFieldDef |
      IMultiChoiceFieldDef |
      IMultiLineTextFieldDef |
      INumberFieldDef |
      ITextFieldDef |
      IUrlFieldDef |
      IUserFieldDef
  ): Promise<void> {

    const cField: any = await field.get();

    let fldUpdateResult: FieldUpdateResult;
    let fType: string;
    let props: FieldCreationProperties & {
      FieldTypeKind?: number,
      InternalName?: string,
      Sortable?: boolean,
      Choices?: { results: string[] }
    } = fieldDef.properties;
    // props.StaticName = undefined;
    props.Title = fieldDef.title;

    try {
      switch (fieldDef.fieldType) {
        case FieldTypes.Boolean:
          fType = "SP.Field";
          break;
        case FieldTypes.Calculated:
          var fdCalc: ICalcFieldDef = this.fldType<ICalcFieldDef>(fieldDef as ICalcFieldDef);
          props.Formula = fdCalc.formula || cField["Formula"];
          props.DateFormat = fdCalc.dateFormat || cField["DateFormat"];
          props.OutputType = fdCalc.outputType || cField["OutputType"];
          fType = "SP.FieldCalculated";
          break;
        case FieldTypes.Choice:
          var fdChoice: IChoiceFieldDef = this.fldType<IChoiceFieldDef>(fieldDef as IChoiceFieldDef);
          props.Choices = { results: fdChoice.choices } || cField["Choices"];
          props.FillInChoice = fdChoice.fillIn || cField["FillInChoice"];
          props.EditFormat = fdChoice.format || cField["EditFormat"];
          fType = "SP.FieldChoice";
          break;
        case FieldTypes.Currency:
          var fdCurrency: ICurrencyFieldDef = this.fldType<ICurrencyFieldDef>(fieldDef as ICurrencyFieldDef);
            props.MinimumValue = fdCurrency.minValue || cField["MinimumValue"];
            props.MaximumValue = fdCurrency.maxValue || cField["MaximumValue"];
            props.CurrencyLocaleId = fdCurrency.currencyLocalId || cField["CurrencyLocaleId"];
            fType = "SP.FieldCurrency";
          break;
        case FieldTypes.DateTime:
          var fdDate: IDateTimeFieldDef = this.fldType<IDateTimeFieldDef>(fieldDef as IDateTimeFieldDef);
          props.DisplayFormat = fdDate.displayFormat || cField["DisplayFormat"];
          props.DateTimeCalendarType = fdDate.calendarType || cField["DateTimeCalendarType"];
          props.FriendlyDisplayFormat = fdDate.friendlyDisplayFormat || cField["FriendlyDisplayFormat"];
          fType = "SP.FieldDateTime";
          break;
        case FieldTypes.Lookup:
          var fdLookup: ILookupFieldDef = this.fldType<ILookupFieldDef>(fieldDef as ILookupFieldDef);
          // props.LookupList = fdLookup.lookupListId || cField["LookupList"];
          props.LookupField = fdLookup.lookupFieldName || cField["LookupField"];
          fType = "SP.FieldLookup";
          break;
        case FieldTypes.MultiChoice:
          var fdMChoice: IMultiChoiceFieldDef =
            this.fldType<IMultiChoiceFieldDef>(fieldDef as IMultiChoiceFieldDef);
          props.Choices = { results: fdMChoice.choices } || cField["Choices"];
          props.FillInChoice = fdMChoice.fillIn || cField["FillInChoice"];
          fType = "SP.FieldMultiChoice";
          break;
        case FieldTypes.Note:
          const fdMultiLine: IMultiLineTextFieldDef =
            this.fldType<IMultiLineTextFieldDef>(fieldDef as IMultiLineTextFieldDef);
          const richTextAttr = /\bRichTextMode\s*=\s*"([^"]*)"/g.exec(<string>cField["SchemaXml"]);
          const rTxtToReplace: string = richTextAttr ? richTextAttr[0] : `/>`;
          const rTxtReplacement: string = richTextAttr ? `RichTextMode="FullHtml"` : `RichTextMode="FullHtml" />`;
          props.NumberOfLines = fdMultiLine.numberOfLines || cField["NumberOfLines"];
          props.RichText = fdMultiLine.richText || cField["RichText"];
          props.RestrictedMode = fdMultiLine.restrictedMode || cField["RestrictedMode"];
          props.AppendOnly = fdMultiLine.appendOnly || cField["AppendOnly"];
          props.AllowHyperlink = fdMultiLine.allowHyperLink || cField["AllowHyperlink"];
          props.SchemaXml = props.RestrictedMode ? cField["SchemaXml"] : (<string>cField["SchemaXml"]).replace(rTxtToReplace, rTxtReplacement);
          fType = "SP.FieldMultiLineText";
          break;
        case FieldTypes.Number:
          var fdNumber: INumberFieldDef = this.fldType<INumberFieldDef>(fieldDef as INumberFieldDef);
          props.MinimumValue = fdNumber.minValue || cField["MinimumNumber"];
          props.MaximumValue = fdNumber.maxValue || cField["MaximumNumber"];
          fType = "SP.FieldNumber";
          break;
        case FieldTypes.Text:
          var fdText: ITextFieldDef = this.fldType<ITextFieldDef>(fieldDef as ITextFieldDef);
          props.MaxLength = fdText.maxLength || cField["MaxLength"];
          fType = "SP.FieldText";
          break;
        case FieldTypes.URL:
          var fdURL: IUrlFieldDef = this.fldType<IUrlFieldDef>(fieldDef as IUrlFieldDef);
          props.DisplayFormat = fdURL.displayFormat || cField["DisplayFormat"];
          fType = "SP.FieldUrl";
          break;
        case FieldTypes.User:
          var fdUser: IUserFieldDef = this.fldType<IUserFieldDef>(fieldDef as IUserFieldDef);
          props.SelectionMode = fdUser.selectionMode || cField["SelectionMode"];
          fType = "SP.FieldUser";
          break;
      }
      fldUpdateResult = await field.update(props, fType);
      Logger.writeJSON(fldUpdateResult);
    } catch (err) {
      Logger.writeJSON(err);
    }
  }
}