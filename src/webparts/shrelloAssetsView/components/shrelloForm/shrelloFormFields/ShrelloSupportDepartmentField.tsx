import * as React from "react";
import {
  ChoiceGroup,
  IChoiceGroupProps,
  IChoiceGroupOption,
  IChoiceGroupState
} from "office-ui-fabric-react/lib/ChoiceGroup";

import {
  IShrelloItem,
  ISupportDepartmentItem
} from "../../../models";

export interface IShrelloSupportDepartmentProps extends IChoiceGroupProps {
  item: IShrelloItem;
  supportDepartments: ISupportDepartmentItem[];
  choices?: IChoiceGroupOption[];
}

export interface IShrelloSupportDepartmentState extends IChoiceGroupState {
  options: IChoiceGroupOption[];
}

export default class ShrelloSupportDepartmentField extends React.Component<IShrelloSupportDepartmentProps, IShrelloSupportDepartmentState> {
  private _supportDepartments: ISupportDepartmentItem[] = [];
  private _options: IChoiceGroupOption[] = [];

  constructor(props: IShrelloSupportDepartmentProps) {
    super(props);

    this.state = {
      options: this._options,
      keyChecked: null
    };

    this._handleChange = this._handleChange.bind(this);
  }

  private _makeDepartmentOptions = (departments: ISupportDepartmentItem[]): IChoiceGroupOption[] => {
    const cDept: ISupportDepartmentItem[] = departments
      .filter(d => d.Title !== "How to use TASC.");
    return cDept.map(department => {
      const cOption: IChoiceGroupOption = {
        key: String(department.Id),
        text: department.Title,
        iconProps: department.IconName ? { iconName: department.IconName } : undefined
      };
      return cOption;
    });
  }

  private _handleChange = (ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption): void => this.props.onChange(ev, option);

  public componentDidMount(): void {
    this._supportDepartments = this.props.supportDepartments;
    this._options = this._makeDepartmentOptions(this._supportDepartments);
    this.setState({ options: this._options });
  }

  public render(): JSX.Element {
    const { item } = this.props;
    const { options } = this.state;
    const supportDepartmentId: string = item.SupportDepartmentId.toString();

    return (
      <ChoiceGroup
        label="Support Department:"
        id="SupportDepartmentSelect"
        options={ options }
        selectedKey={ supportDepartmentId }
        onChange={ this._handleChange }
      />
    );
  }
}