import * as React from "react";
import {
  ChoiceGroup,
  IChoiceGroupProps,
  IChoiceGroupOption,
  IChoiceGroupState
} from "office-ui-fabric-react/lib/ChoiceGroup";

import { IShrelloItem } from "../../../models";
import { Categories } from "../../../config/shrelloViewConfig";

export interface IShrelloCategoryProps extends IChoiceGroupProps {
  item: IShrelloItem;
}

export default class ShrelloCategoryField extends React.Component<IShrelloCategoryProps, IChoiceGroupState> {
  private _options: IChoiceGroupOption[] = [];

  constructor(props: IShrelloCategoryProps) {
    super(props);

    this._options = this._makeCategoryOptions();

    this.state = {
      keyChecked: null
    };
    this._handleChange = this._handleChange.bind(this);
  }

  private _makeCategoryOptions = (): IChoiceGroupOption[] => {
    return Categories.map(category => {
      const cOption: IChoiceGroupOption = {
        key: category.name,
        text: category.text || category.name,
        iconProps: category.iconName ? { iconName: category.iconName, className: category.iconColor } : undefined
      };
      return cOption;
    });
  }

  private _handleChange = (ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption): void => this.props.onChange(ev, option);

  public render(): JSX.Element {
    const { item } = this.props;

    return (
      <ChoiceGroup
        label="Category:"
        id="CategorySelect"
        options={ this._options }
        selectedKey={ item.TASCTypeCategory }
        onChange={ this._handleChange }
      />
    );
  }
}