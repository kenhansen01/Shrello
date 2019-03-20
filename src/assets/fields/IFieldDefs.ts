import {
  FieldCreationProperties,
  FieldTypes,
  DateTimeFieldFormatType,
  CalendarType,
  UrlFieldFormatType,
  FieldUserSelectionMode,
  ChoiceFieldFormatType } from "@pnp/sp";

export interface IFieldDef {
  title: string;
  fieldType: FieldTypes;
  fieldSchema?: string;
  viewIndex?: number;
  properties?: FieldCreationProperties & {
    InternalName?: string,
    Sortable?: boolean
  };
}

export interface ICustomFieldDef {
  title: string;
  fieldType: string;
  fieldSchema?: string;
  viewIndex?: number;
  properties: FieldCreationProperties & {
    FieldTypeKind: number,
    InternalName?: string,
    Sortable?: boolean
  };
}

export interface ITextFieldDef extends IFieldDef {
  maxLength?: number;
}

export interface ICalcFieldDef extends IFieldDef {
  formula: string;
  dateFormat: DateTimeFieldFormatType;
  outputType?: FieldTypes;
}

export interface IDateTimeFieldDef extends IFieldDef {
  displayFormat?: DateTimeFieldFormatType;
  calendarType?: CalendarType;
  friendlyDisplayFormat?: number;
}

export interface INumberFieldDef extends IFieldDef {
  minValue?: number;
  maxValue?: number;
}

export interface ICurrencyFieldDef extends INumberFieldDef {
  currencyLocalId?: number;
}

export interface IMultiLineTextFieldDef extends IFieldDef {
  numberOfLines?: number;
  richText?: boolean;
  restrictedMode?: boolean;
  appendOnly?: boolean;
  allowHyperLink?: boolean;
}

export interface IUrlFieldDef extends IFieldDef {
  displayFormat?: UrlFieldFormatType;
}

export interface IUserFieldDef extends IFieldDef {
  selectionMode?: FieldUserSelectionMode;
}

export interface ILookupFieldDef extends IFieldDef {
  lookupListId: string;
  lookupFieldName: string;
}

export interface IMultiChoiceFieldDef extends IFieldDef {
  choices: string[];
  fillIn?: boolean;
}

export interface IChoiceFieldDef extends IMultiChoiceFieldDef {
  format?: ChoiceFieldFormatType;
}

export interface IBooleanFieldDef extends IFieldDef {}
