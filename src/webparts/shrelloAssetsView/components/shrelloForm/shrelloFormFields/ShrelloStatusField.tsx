import * as React from "react";
import {
  ChoiceGroup,
  IChoiceGroupProps,
  IChoiceGroupOption,
  IChoiceGroupState
} from "office-ui-fabric-react/lib/ChoiceGroup";

// import { IShrelloItem } from "../../../models";
import { Statuses } from "../../../config/shrelloViewConfig";

export interface IShrelloStatusProps extends IChoiceGroupProps {
  status: string;
}

export default class ShrelloStatusField extends React.Component<IShrelloStatusProps, IChoiceGroupState> {
  private _options: IChoiceGroupOption[] = [];

  constructor(props: IShrelloStatusProps) {
    super(props);

    this._options = this._makeStatusOptions();

    this.state = {
      keyChecked: null
    };

    this._handleChange = this._handleChange.bind(this);
  }

  private _makeStatusOptions = (): IChoiceGroupOption[] => {
    return Statuses.map(status => {
      const sOption: IChoiceGroupOption = {
        key: String(status.name),
        text: status.text || status.name,
        iconProps: status.iconName ? { iconName: status.iconName, className: status.iconColor } : undefined
      };
      return sOption;
    });
  }

  private _handleChange = (ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption): void => this.props.onChange(ev, option);

  public render(): JSX.Element {
    const { status } = this.props;

    return (
      <ChoiceGroup
        label="Status:"
        id="StatusSelect"
        options={ this._options }
        selectedKey={ status }
        onChange={ this._handleChange }
      />
    );
  }
}