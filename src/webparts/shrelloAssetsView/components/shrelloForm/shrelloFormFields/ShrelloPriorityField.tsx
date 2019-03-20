import * as React from "react";
import {
  ChoiceGroup,
  IChoiceGroupProps,
  IChoiceGroupOption,
  IChoiceGroupState
} from "office-ui-fabric-react/lib/ChoiceGroup";

import { IShrelloItem } from "../../../models";
import { Priorities } from "../../../config/shrelloViewConfig";

export interface IShrelloPriorityProps extends IChoiceGroupProps {
  item: IShrelloItem;
}

export default class ShrelloPriorityField extends React.Component<IShrelloPriorityProps, IChoiceGroupState> {
  private _options: IChoiceGroupOption[] = [];

  constructor(props: IShrelloPriorityProps) {
    super(props);

    this._options = this._makePriorityOptions();

    this.state = {
      keyChecked: null
    };

    this._handleChange = this._handleChange.bind(this);
  }

  private _makePriorityOptions = (): IChoiceGroupOption[] => {
    return Priorities.map(priority => {
      const pOption: IChoiceGroupOption = {
        key: String(priority.name),
        text: priority.text || priority.name,
        iconProps: priority.iconName ? { iconName: priority.iconName, className: priority.iconColor } : undefined
      };
      return pOption;
    });
  }

  private _handleChange = (ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption): void => this.props.onChange(ev, option);

  public render(): JSX.Element {
    const { item } = this.props;

    return (
      <ChoiceGroup
        label="Priority:"
        id="PrioritySelect"
        options={ this._options }
        selectedKey={ item.Priority }
        onChange={ this._handleChange }
      />
    );
  }
}