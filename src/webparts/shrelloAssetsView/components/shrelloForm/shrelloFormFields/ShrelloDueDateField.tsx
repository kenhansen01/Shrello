import * as React from "react";
import {
  DatePicker,
  IDatePickerProps,
  IDatePickerState,
  DayOfWeek
} from "office-ui-fabric-react/lib/DatePicker";

import { IShrelloItem } from "../../../models";
import { DayPickerStrings } from "../../../config/shrelloViewConfig";

export interface IShrelloDueDateProps extends IDatePickerProps {
  item: IShrelloItem;
}

export default class ShrelloDueDateField extends React.Component<IShrelloDueDateProps, IDatePickerState> {
  constructor(props: IShrelloDueDateProps) {
    super(props);
    this._handleChange = this._handleChange.bind(this);
  }

  private _handleChange = (date: Date | null | undefined): void => {
    this.props.onSelectDate(date);
  }

  public render(): JSX.Element {
    const { item, label } = this.props;
    return (
      <DatePicker
        label={ label ? label : "Requested Due Date:" }
        isRequired={ false }
        allowTextInput={ true }
        firstDayOfWeek={ DayOfWeek.Sunday }
        strings={ DayPickerStrings }
        value={ item.DueDate }
        showMonthPickerAsOverlay={ true }
        onSelectDate={ this._handleChange }
        placeholder="Enter a due date."
      />
    );
  }

}
