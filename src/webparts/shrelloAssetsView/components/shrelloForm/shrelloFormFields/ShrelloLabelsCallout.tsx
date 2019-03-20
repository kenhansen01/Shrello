import * as React from "react";
import {
  Callout,
  ICalloutProps
} from "office-ui-fabric-react/lib/Callout";
import { IChoiceGroupOption } from "office-ui-fabric-react/lib/ChoiceGroup";

import {
  ShrelloSupportDepartmentField,
  ShrelloPriorityField,
  ShrelloCategoryField,
  ShrelloStatusField
} from "./ShrelloFormFields";
import {
  IShrelloItem,
  ISupportDepartmentItem
} from "../../../models";
import styles from "../ShrelloForm.module.scss";
import { IconButton } from "office-ui-fabric-react";
import { IconNames } from "@uifabric/icons";

export interface IShrelloLabelsCalloutProps extends ICalloutProps {
  item?: IShrelloItem;
  supportDepartments: ISupportDepartmentItem[];
  buttonElement: HTMLElement;
  onSupportDepartmentChange: (
    ev: React.SyntheticEvent<HTMLElement>,
    option: IChoiceGroupOption
  ) => void;
  onPriorityChange: (
    ev: React.SyntheticEvent<HTMLElement>,
    option: IChoiceGroupOption
  ) => void;
  onStatusChange: (
    ev: React.SyntheticEvent<HTMLElement>,
    option: IChoiceGroupOption
  ) => void;
  onCategoryChange: (
    ev: React.SyntheticEvent<HTMLElement>,
    option: IChoiceGroupOption
  ) => void;
}

export class ShrelloLabelsCallout extends React.Component<IShrelloLabelsCalloutProps> {
  public constructor(props: IShrelloLabelsCalloutProps) {
    super(props);
    this.state = {
      isCalloutVisible: true
    };
    this._handleSupportDepartmentChange = this._handleSupportDepartmentChange.bind(this);
    this._handlePriorityChange = this._handlePriorityChange.bind(this);
    this._handleStatusChange = this._handleStatusChange.bind(this);
    this._handleCategoryChange = this._handleCategoryChange.bind(this);
  }

  private _handleSupportDepartmentChange = (
    ev: React.SyntheticEvent<HTMLElement>,
    option: IChoiceGroupOption
  ): void => this.props.onSupportDepartmentChange(ev, option);

  private _handlePriorityChange = (
    ev: React.SyntheticEvent<HTMLElement>,
    option: IChoiceGroupOption
  ): void => this.props.onPriorityChange(ev, option);

  private _handleStatusChange = (
    ev: React.SyntheticEvent<HTMLElement>,
    option: IChoiceGroupOption
  ): void => this.props.onStatusChange(ev, option);

  private _handleCategoryChange = (
    ev: React.SyntheticEvent<HTMLElement>,
    option: IChoiceGroupOption
  ): void => this.props.onCategoryChange(ev, option);

  public render(): JSX.Element {
    const { buttonElement, onDismiss, item, supportDepartments } = this.props;

    return (
      <div>
        <Callout
          role={ "alertdialog" }
          ariaLabelledBy={ "shrello-labels" }
          target={ buttonElement }
          onDismiss={ onDismiss }
          className={styles.shrelloActionCallout}
        >
          <IconButton
            iconProps={{ iconName: IconNames.ChromeClose }}
            onClick={ onDismiss }
            className={ styles.calloutDismiss }
          />
          <ShrelloSupportDepartmentField
            item={ item }
            supportDepartments={supportDepartments}
            onChange={ this._handleSupportDepartmentChange }
          />
          <hr/>
          <ShrelloPriorityField
            item={item}
            onChange={this._handlePriorityChange}
          />
          <hr/>
          <ShrelloStatusField
            status={item.Status}
            onChange={this._handleStatusChange}
          />
          <hr/>
          <ShrelloCategoryField
            item={item}
            onChange={this._handleCategoryChange}
          />
        </Callout>
      </div>
    );
  }
}