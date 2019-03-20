import {
  FieldTypes,
  FieldUserSelectionMode,
  UrlFieldFormatType
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

const SupportTopicListFields: (
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
    title: "Support Topic",
    fieldType: FieldTypes.Text,
    properties: {
      InternalName: "Title"
    }
  },
  // the following are custom fields
  <ITextFieldDef> {
    title: "TASC Abbreviation",
    fieldType: FieldTypes.Text,
    properties: {
      InternalName: "TASCAbbreviation",
      Required: true,
      EnforceUniqueValues: true,
      Indexed: true,
      Sortable: true
    }
  },
  <IMultiLineTextFieldDef> {
    title: "Support Topic Description",
    fieldType: FieldTypes.Note,
    numberOfLines: 8,
    richText: true,
    allowHyperLink: true,
    restrictedMode: false,
    properties: {
      InternalName: "SupportTopicDescription"
    }
  },
  <IMultiLineTextFieldDef> {
    title: "Language",
    fieldType: FieldTypes.Note,
    richText: false,
    properties: {
      InternalName: "Language"
    }
  },
  <ITextFieldDef> {
    title: "Classification",
    fieldType: FieldTypes.Text,
    properties: {
      InternalName: "Classification",
      Sortable: true
    }
  },
  <ITextFieldDef> {
    title: "Vendor",
    fieldType: FieldTypes.Text,
    properties: {
      InternalName: "Vendor",
      Sortable: true
    }
  },
  <ITextFieldDef> {
    title: "Vendor Contact",
    fieldType: FieldTypes.Text,
    properties: {
      InternalName: "VendorContact",
      Sortable: true
    }
  },
  <ITextFieldDef> {
    title: "Vendor Email",
    fieldType: FieldTypes.Text,
    properties: {
      InternalName: "VendorEmail",
      Sortable: true
    }
  },
  <IUrlFieldDef> {
    title: "Vendor Website",
    fieldType: FieldTypes.URL,
    displayFormat: UrlFieldFormatType.Hyperlink,
    properties: {
      InternalName: "VendorWebsite"
    }
  },
  <IBooleanFieldDef> {
    title: "On Request Form",
    fieldType: FieldTypes.Boolean,
    properties: {
      InternalName: "OnRequestForm"
    }
  },
  <IUserFieldDef> {
    title: "Responsible Supervisor",
    fieldType: FieldTypes.User,
    selectionMode: FieldUserSelectionMode.PeopleOnly,
    properties: {
      InternalName: "ResponsibleSupervisor",
      AllowMultipleValues: true,
      Presence: true,
      Sortable: true
    }
  },
  <IUserFieldDef> {
    title: "Primary Support",
    fieldType: FieldTypes.User,
    selectionMode: FieldUserSelectionMode.PeopleOnly,
    properties: {
      InternalName: "PrimarySupport",
      AllowMultipleValues: true,
      Presence: true,
      Sortable: true
    }
  },
  <IUserFieldDef> {
    title: "Secondary Support",
    fieldType: FieldTypes.User,
    selectionMode: FieldUserSelectionMode.PeopleOnly,
    properties: {
      InternalName: "SecondarySupport",
      AllowMultipleValues: true,
      Presence: true,
      Sortable: true
    }
  },
  <IUserFieldDef> {
    title: "Contact for Support",
    fieldType: FieldTypes.User,
    selectionMode: FieldUserSelectionMode.PeopleOnly,
    properties: {
      InternalName: "SupportContact",
      AllowMultipleValues: true,
      Presence: true,
      Sortable: true
    }
  },
  <IUserFieldDef> {
    title: "Access Request Contact",
    fieldType: FieldTypes.User,
    selectionMode: FieldUserSelectionMode.PeopleOnly,
    properties: {
      InternalName: "AccessContact",
      AllowMultipleValues: true,
      Presence: true,
      Sortable: true
    }
  },
  <INumberFieldDef> {
    title: "Ticket Count",
    fieldType: FieldTypes.Number,
    minValue: 0,
    properties: {
      InternalName: "TicketCount",
      DefaultValue: "0"
    }
  }
];

export default SupportTopicListFields;
