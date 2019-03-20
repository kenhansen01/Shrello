import { FieldTypes, ChoiceFieldFormatType } from "@pnp/sp";
import {
  IBooleanFieldDef,
  ICalcFieldDef,
  IChoiceFieldDef,
  ICurrencyFieldDef,
  ICustomFieldDef,
  IDateTimeFieldDef,
  ILookupFieldDef,
  IMultiChoiceFieldDef,
  IMultiLineTextFieldDef,
  INumberFieldDef,
  ITextFieldDef,
  IUrlFieldDef,
  IUserFieldDef
} from "../../fields/IFieldDefs";
import { IconNames } from "@uifabric/icons";

const SupportDepartmentListFields: (
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
)[] = [
  <ITextFieldDef> {
    title: "Support Department",
    fieldType: FieldTypes.Text,
    properties: {
      InternalName: "Title"
    }
  },
  <IMultiLineTextFieldDef> {
    title: "Support Department Information",
    fieldType: FieldTypes.Note,
    numberOfLines: 8,
    richText: true,
    restrictedMode: false,
    allowHyperLink: true,
    properties: {
      InternalName: "SupportDepartmentInformation",
      Hidden: false,
      Required: true
    }
  },
  <IChoiceFieldDef> {
    title: "Icon Name",
    fieldType: FieldTypes.Choice,
    choices: [
      IconNames.Teamwork,
      IconNames.DeveloperTools,
      IconNames.ServerEnviroment,
      IconNames.Money
    ],
    fillIn: true,
    format: ChoiceFieldFormatType.Dropdown,
    properties: {
      InternalName: "IconName",
      Required: false
    }
  },
  <IBooleanFieldDef> {
    title: "On Form",
    fieldType: FieldTypes.Boolean,
    properties: {
      InternalName: "OnForm"
    }
  }
];

export default SupportDepartmentListFields;
