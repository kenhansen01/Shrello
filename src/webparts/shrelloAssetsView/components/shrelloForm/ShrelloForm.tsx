import * as React from "react";
import { IconNames } from "@uifabric/icons";
import {
  DefaultButton,
  PrimaryButton,
  IChoiceGroupOption,
  IDropdownOption,
  Icon,
  Label,
  ColorClassNames,
  IconButton,
  css
} from "office-ui-fabric-react";
import { Modal } from "office-ui-fabric-react/lib/Modal";
import { addDays } from "office-ui-fabric-react/lib/utilities/dateMath/DateMath";
import update from "immutability-helper";

import {
  ShrelloTitleField,
  ShrelloRequestField,
  ShrelloLabelsCallout,
  ShrelloSupportTopicField,
  ShrelloDueDateField,
  ShrelloAttachments,
  ShrelloMembersCalloutField,
  ShrelloCommentField
} from "./shrelloFormFields/ShrelloFormFields";
import {
  IShrelloItem,
  ISupportDepartmentItem,
  ISupportTopicItem
} from "../../models";
import {
  IShrelloLabel,
  Priorities,
  Categories,
  Statuses
} from "../../config/shrelloViewConfig";
import {
  IShrelloFormProps,
  IShrelloFormState
} from "./IShrelloForm";
import styles from "./ShrelloForm.module.scss";

function FormModal(props: any): JSX.Element {
  return (
    <div>
      <Modal
        { ...props }
        containerClassName={ styles.shrelloForm }
      >
        <div className={ styles.shrelloFormHeader }>TASC Request{
          !!props.item.Title &&
          ": " + props.item.Title
        }</div>
        { props.formBody }
        { props.formCallouts }
        { props.formActions}
      </Modal>
    </div>
  );
}

function FormBody(props: any): JSX.Element {
  const {
    item,
    supportDepartmentItem,
    priority,
    status,
    category,
    supportTopics,
    supportDepartment,
    onTitleChanged,
    onSupportTopicChange,
    onRequestChanged,
    onFileDrop,
    onCommentChanged,
    onSelectDueDate,
    showLabels,
    showMembers } = props;
  return (
    <div className={ styles.shrelloFormBody }>
      <div className="ms-Grid-row">
        <div className="ms-Grid-col ms-sm12 ms-lg9">
          <ShrelloTitleField
            title={ item.Title }
            onChanged={ onTitleChanged }
          />
          <div className={ styles.formLabels }>
            <Label>Labels:</Label>
            { item.SupportDepartmentId && supportDepartmentItem &&
              <span className={ styles.formLabel }>
                <Icon iconName={ supportDepartmentItem.IconName }/>
                <Label className={ styles.formLabelText }>{ supportDepartmentItem.Title }</Label>
              </span>
            }
            { item.Priority && priority &&
              <span className={ styles.formLabel }>
                <Icon
                  iconName={ priority.iconName }
                  className={ priority.iconColor }
                />
                <Label className={ styles.formLabelText }>{ priority.name }</Label>
              </span>
            }
            { item.Status && status &&
              <span className={ styles.formLabel }>
                <Icon
                  iconName={ status.iconName }
                  className={ status.iconColor }
                />
                <Label className={ styles.formLabelText }>{ status.name }</Label>
              </span>
            }
            { item.TASCTypeCategory && category &&
              <span className={ styles.formLabel }>
                <Icon
                  iconName={ category.iconName }
                  className={ category.iconColor }
                />
                <Label className={ styles.formLabelText }>{ category.name }</Label>
              </span>
            }
            <span className={ styles.formLabel }>
              <IconButton
                iconProps={{
                  iconName: IconNames.BoxAdditionSolid,
                  className: css(ColorClassNames.neutralLight, ColorClassNames.neutralDarkHover, styles.formLabelIcon)
                }}
                onClick={ props.onToggleLabels }
              />
              <Label className={ styles.formLabelText }>Add a Label</Label>
            </span>           
          </div>
          <ShrelloSupportTopicField
            item={ item }
            supportTopics={supportTopics}
            selectedDepartment={supportDepartment || null}
            onChanged={ onSupportTopicChange }
          />
          <Label>Full Request:</Label>
          <ShrelloRequestField
            item={ item }
            onSave={ onRequestChanged }
          />
          <ShrelloAttachments
            item={item}
            handleDrop={ onFileDrop }
          />
          <ShrelloCommentField
            item={item}
            onChanged={ onCommentChanged }
          />                
        </div>
        <div className="ms-Grid-col ms-sm12 ms-lg3">
          <h2>Add</h2>
          <DefaultButton
            text="Members"
            onClick={ props.onToggleMembers }
            className={styles.shrelloActions}
            iconProps={{iconName: IconNames.People}}
          />
          <DefaultButton
            text="TASC Labels"
            onClick={ props.onToggleLabels }
            className={styles.shrelloActions}
            iconProps={{iconName: IconNames.Label}}
          />
          <ShrelloDueDateField
            item={ item }
            label="Requested Due Date"
            onSelectDate={ props.onSelectDueDate }
          />
        </div>              
      </div>
    </div>
  );
}

function FormCallouts(props: any): JSX.Element {
  return (
   <div>
     { props.showLabels &&
        <ShrelloLabelsCallout {...props}
          onDismiss={ props.onDismissLabels }
        />
      }
      { props.showMembers &&
        <ShrelloMembersCalloutField {...props}
          onDismiss={ props.onDismissMembers }
        />
      }
   </div>
 );
}

function FormActions(props: any): JSX.Element {
  return (
    <div className={ styles.shrelloFormFooter }>
      <PrimaryButton
        className={ styles.shrelloFormAction }
        onClick={ props.onSubmitForm }
        text={ props.isNew ? "Submit Request" : "Update Item" }
      />
      <DefaultButton
        className={ styles.shrelloFormAction }
        onClick={ props.onCancelForm }
        text="Cancel"
      />
    </div>
  );
}

function FormParts(props: any): JSX.Element {
  return (
    <FormModal {...props}
      formBody={ <FormBody {...props} /> }
      formCallouts={ <FormCallouts {...props} /> }
      formActions={ <FormActions {...props} /> }
    />
  );
}

export default class ShrelloForm extends React.Component<IShrelloFormProps, IShrelloFormState> {

  private _buttonClicked: HTMLElement = null;
  private _newFiles: boolean = false;

  constructor(props: IShrelloFormProps) {
    super(props);
    let fItem: IShrelloItem = props.item;
    if (fItem && fItem.DueDate) {
      fItem.DueDate = new Date(fItem.DueDate);
    }
    if (fItem && fItem.StartDate) {
      fItem.StartDate = new Date(fItem.StartDate);
    }
    const supportTopics: ISupportTopicItem[] = props.supportTopics.length && fItem && fItem.SupportDepartmentId
      ? props.supportTopics.filter(topic =>
        topic.SupportDepartmentId === fItem.SupportDepartmentId)
      : [];

    const supportDepartment: ISupportDepartmentItem = props.supportDepartments.find(d => d.Id === props.item.SupportDepartmentId);

    this.state = {
      supportDepartment: supportDepartment.Title,
      supportTopic: null,
      showForm: props.showForm,
      showLabels: false,
      showMembers: false,
      supportTopics: supportTopics,
      item: fItem ? fItem : {
        Title: undefined,
        SupportDepartmentId: 1,
        TASCTypeCategory: "Unknown",
        Priority: "3 - Normal",
        DueDate: addDays((new Date(Date.now())), 14)
      },
      isNew: !fItem ? true : !fItem.Title ? true : false
    };
  }

  public componentWillReceiveProps(props: IShrelloFormProps): void {
    let supportDepartment: ISupportDepartmentItem = null;
    let supportTopics: ISupportTopicItem[] = props.supportTopics;
    if (props.item) {
      if (props.item.DueDate) {
        props.item.DueDate = new Date(props.item.DueDate);
      }
      if (props.item.StartDate) {
        props.item.StartDate = new Date(props.item.StartDate);
      }
      supportDepartment = props.supportDepartments.find(d => d.Id === props.item.SupportDepartmentId);

      supportTopics = props.supportTopics.filter(topic =>
        topic.SupportDepartmentId === props.item.SupportDepartmentId);
    }

    const sDepartmentTitle: string = supportDepartment ? supportDepartment.Title : null;

    this.setState({
      supportDepartment: sDepartmentTitle,
      showForm: props.showForm,
      supportTopics: supportTopics,
      item: props.item ? props.item : {
        Title: undefined,
        SupportDepartmentId: 1,
        TASCTypeCategory: "Unknown",
        SupportTopicId: supportTopics[0].Id,
        Priority: "3 - Normal",
        DueDate: addDays((new Date(Date.now())), 14)
      },
      isNew: !props.item ? true : !props.item.Title ? true : false
    });
  }

  // public componentDidMount(): void {}

  private _onTitleChanged = (newValue:any): void => {
    const updatedItem: IShrelloItem = update(this.state.item, { Title: {$set: newValue} });
    this.setState({ item: updatedItem });
  }

  private _onCommentChanged = (newValue:any): void => {
    const updatedItem: IShrelloItem = update(this.state.item, { TASCComments: {$set: newValue} });
    this.setState({ item: updatedItem });
  }

  private _onRequestChanged = (rqHtml: string): void => {
    const updatedItem: IShrelloItem = update(this.state.item, { Body: {$set: rqHtml} });
    this.setState({ item: updatedItem });
  }

  private _toggleLabels = (ev: React.MouseEvent<HTMLButtonElement>) => {
    this._buttonClicked = ev.target as HTMLElement;
    this.setState({
      showLabels: !this.state.showLabels
    });
  }

  private _dismissLabels = (ev: React.MouseEvent<HTMLButtonElement>) => {
    this._buttonClicked = ev.target as HTMLElement;
    this.setState({
      showLabels: false
    });
  }

  private _toggleMembers = (ev: React.MouseEvent<HTMLButtonElement>) => {
    this._buttonClicked = ev.target as HTMLElement;
    this.setState({
      showMembers: !this.state.showMembers
    });
  }

  private _dismissMembers = (ev: React.MouseEvent<HTMLButtonElement>) => {
    this._buttonClicked = ev.target as HTMLElement;
    this.setState({
      showMembers: false
    });
  }

  private _onUpdateMembers = (ev: React.MouseEvent<HTMLButtonElement>, item: IShrelloItem) => {
    let updatedItem: IShrelloItem = update(this.state.item, {
      RequesterId: { $set: item.RequesterId },
      Requester: { $set: item.Requester },
      AssignedToId: { $set: item.AssignedToId },
      AssignedTo: { $set: item.AssignedTo },
      WatchingId: { $set: item.WatchingId },
      Watching: { $set: item.Watching }
    });
    this.setState({ item: updatedItem, showMembers: false });
    // this._dismissMembers(ev);
  }

  private _onSupportDepartmentChange = (ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption): void => {
    const supportTopics: ISupportTopicItem[] = this.props.supportTopics.filter(topic =>
      topic.SupportDepartmentId === parseInt(option.key, 10));
    const updatedItem: IShrelloItem = update(this.state.item, { SupportDepartmentId: { $set: parseInt(option.key, 10) } });
    this.setState({
      supportTopics: supportTopics,
      supportDepartment: option.text,
      item: updatedItem
    });
  }

  private _onSupportTopicChange = (option: IDropdownOption): void => {
    const oKey: string = option.key as string;
    const updatedItem: IShrelloItem = update(this.state.item, { SupportTopicId: { $set: parseInt(oKey, 10) } });
    this.setState({ item: updatedItem });
  }

  private _onPriorityChange = (ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption): void => {
    const updatedItem: IShrelloItem = update(this.state.item, { Priority: { $set: option.key } });
    this.setState({ item: updatedItem });
  }

  private _onStatusChange = (ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption): void => {
    const updatedItem: IShrelloItem = update(this.state.item, { Status: { $set: option.key } });
    this.setState({ item: updatedItem });
  }

  private _onCategoryChange = (ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption): void => {
    const updatedItem: IShrelloItem = update(this.state.item, { TASCTypeCategory: { $set: option.key } });
    this.setState({ item: updatedItem });
  }

  private _onSelectDueDate = (date: Date | null | undefined): void => {
    const updatedItem: IShrelloItem = update(this.state.item, { DueDate: { $set: date } });
    this.setState({ item: updatedItem });
  }

  private _onFileDrop = (files: File[]) => {
    this.setState({ files });
    this._newFiles = true;
  }

  private _submitForm = async (): Promise<void> => {
    const files: File[] = this._newFiles ? this.state.files : null;
    this.state.isNew
      ? await this.props.onAddShrelloItem(this.state.item, files)
      : await this.props.onUpdateShrelloItem(this.state.item, files);
    this._newFiles = false;
    this._cancelForm();
  }

  private _closeForm = (): void => {
    this.props.onCloseForm();
  }

  private _cancelForm = (): void => {
    const uItem: IShrelloItem = update(this.state.item, {
      Id: { $set: undefined },
      Title: { $set: undefined },
      PercentComplete: { $set: undefined },
      Status: { $set: undefined },
      Priority: { $set: Priorities[1].name },
      Body: { $set: undefined },
      TASCId: { $set: undefined },
      TASCTypeCategory: { $set: Categories[0].name },
      TASCComments: { $set: undefined },
      RequesterId: { $set: undefined },
      Requester: { $set: undefined },
      AssignedToId: { $set: undefined },
      AssignedTo: { $set: undefined },
      WatchingId: { $set: undefined },
      Watching: { $set: undefined },
      PredecessorsId: { $set: undefined },
      SupportDepartmentId: { $set: 1 },
      SupportTopicId: { $set: undefined },
      StartDate: { $set: undefined },
      DueDate: { $set: undefined },
      ParentID: { $set: undefined }
    });
    this.setState({ item: uItem });
    this._closeForm();
  }

  public render(): JSX.Element {
    require("./ShrelloFormStyles.css");
    const { supportDepartments } = this.props;
    const { showForm, showLabels, showMembers, item, isNew, supportTopics, supportDepartment } = this.state;
    const supportDepartmentItem: ISupportDepartmentItem = supportDepartments.find(d => d.Title === supportDepartment);
    const priority: IShrelloLabel = Priorities.find(p => p.name === item.Priority);
    const category: IShrelloLabel = Categories.find(c => c.name === item.TASCTypeCategory);
    const status: IShrelloLabel = Statuses.find(s => s.name === item.Status);

    return (
      <FormParts
        isBlocking={ true }
        isOpen={ showForm }
        onDismiss={ this._cancelForm }
        item={ item }
        isNew={ isNew }
        supportDepartment={ supportDepartment }
        supportDepartmentItem={ supportDepartmentItem }
        supportDepartments={ supportDepartments }
        priority={ priority }
        category={ category }
        status={ status }
        supportTopics={ supportTopics }
        showLabels={ showLabels }
        showMembers={ showMembers }
        onTitleChanged={ this._onTitleChanged }
        onRequestChanged={ this._onRequestChanged }
        onSupportTopicChange={ this._onSupportTopicChange }
        onFileDrop={ this._onFileDrop }
        onCommentChanged={ this._onCommentChanged }
        onSelectDueDate={ this._onSelectDueDate }
        onSupportDepartmentChange={ this._onSupportDepartmentChange }
        onPriorityChange={ this._onPriorityChange }
        onCategoryChange={ this._onCategoryChange }
        onStatusChange={ this._onStatusChange }
        buttonElement={ this._buttonClicked }
        onToggleLabels={ this._toggleLabels }
        onDismissLabels={ this._dismissLabels }
        onToggleMembers={ this._toggleMembers }
        onDismissMembers={ this._dismissMembers }
        onUpdateMembers={ this._onUpdateMembers }
        onSubmitForm={ this._submitForm }
        onCancelForm={ this._cancelForm }
        context={ this.props.context }
      />
    );
  }
}