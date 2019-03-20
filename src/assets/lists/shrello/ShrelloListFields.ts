import {
  FieldTypes,
  ChoiceFieldFormatType,
  FieldUserSelectionMode,
  DateTimeFieldFormatType
} from "@pnp/sp";
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

let ShrelloListFields: (
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
  // the following should only update the display name or choice selections
  <ITextFieldDef> {
    title: "TASC Name",
    fieldType: FieldTypes.Text,
    viewIndex: 0,
    properties: {
      InternalName: "Title"
    }
  },
  <IChoiceFieldDef> {
    title: "TASC Status",
    fieldType: FieldTypes.Choice,
    viewIndex: 9,
    choices: [
      "Not Started",
      "Initiated",
      "Assigned",
      "Planned",
      "In Progress",
      "Ready for Review",
      "Update RunBook",
      "Ready for Release",
      "On Hold",
      "Complete",
      "Canceled"
    ],
    fillIn: false,
    format: ChoiceFieldFormatType.Dropdown,
    properties: {
      InternalName: "Status"
    }
  },
  <IChoiceFieldDef> {
    title: "TASC Priority",
    fieldType: FieldTypes.Choice,
    choices: [
      "1 - Top",
      "2 - High",
      "3 - Normal",
      "4 - Low"
    ],
    fillIn: false,
    format: ChoiceFieldFormatType.Dropdown,
    viewIndex: 7,
    properties: {
      InternalName: "Priority"
    }
  },
  <IMultiLineTextFieldDef> {
    title: "TASC Request",
    fieldType: FieldTypes.Note,
    numberOfLines: 8,
    restrictedMode: false,
    richText: true,
    allowHyperLink: true,
    viewIndex: 6,
    properties: {
      InternalName: "Body",
      Hidden: false,
      Required: true,
      NumberOfLines: 8
    }
  },
  // the following are custom fields
  <IChoiceFieldDef> {
    title: "TASC Type/Category",
    fieldType: FieldTypes.Choice,
    viewIndex: 8,
    choices: [
      "Unknown",
      "Access Request",
      "Bug",
      "Enhancement",
      "Suggestion",
      "Data Issue",
      "Service Request",
      "Other"
    ],
    fillIn: false,
    format: ChoiceFieldFormatType.Dropdown,
    properties: {
      InternalName: "TASCTypeCategory"
    }
  },
  <ITextFieldDef> {
    title: "TASC ID",
    fieldType: FieldTypes.Text,
    viewIndex: 10,
    properties: {
      InternalName: "TASCId",
      Indexed: true,
      EnforceUniqueValues: true,
      Required: false,
      Sortable: true
    }
  },
  <IMultiLineTextFieldDef> {
    title: "TASC Comments",
    fieldType: FieldTypes.Note,
    numberOfLines: 8,
    richText: true,
    restrictedMode: false,
    appendOnly: true,
    allowHyperLink: true,
    viewIndex: 11,
    properties: {
      InternalName: "TASCComments",
      Required: false
    }
  },
  <IUserFieldDef> {
    title: "Requester",
    fieldType: FieldTypes.User,
    selectionMode: FieldUserSelectionMode.PeopleOnly,
    viewIndex: 12,
    properties: {
      InternalName: "Requester",
      Presence: true,
      Sortable: true
    }
  },
  <IUserFieldDef> {
    title: "Watching",
    fieldType: FieldTypes.User,
    selectionMode: FieldUserSelectionMode.PeopleOnly,
    viewIndex: 12,
    properties: {
      InternalName: "Watching",
      AllowMultipleValues: true,
      Presence: true,
      Sortable: true
    }
  },
  <INumberFieldDef> {
    title: "Work Hours Estimate",
    fieldType: FieldTypes.Number,
    minValue: 0,
    viewIndex: 13,
    properties: {
      InternalName: "WorkHoursEstimate"
    }
  },
  <INumberFieldDef> {
    title: "Actual Hours",
    fieldType: FieldTypes.Number,
    minValue: 0,
    viewIndex: 14,
    properties: {
      InternalName: "ActualHours"
    }
  },
  <IDateTimeFieldDef> {
    title: "Date Closed",
    fieldType: FieldTypes.DateTime,
    displayFormat: DateTimeFieldFormatType.DateOnly,
    viewIndex: 15,
    properties: {
      InternalName: "DateClosed"
    }
  }
];

export default ShrelloListFields;
