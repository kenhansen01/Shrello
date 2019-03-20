import * as React from "react";
import { ITextFieldState, ITextFieldProps, TextField } from "office-ui-fabric-react/lib/TextField";

import { IShrelloItem } from "../../../models";
import styles from "../ShrelloForm.module.scss";

export interface IShrelloCommentProps extends ITextFieldProps {
  item: IShrelloItem;
}

export default class ShrelloCommentField extends React.Component<IShrelloCommentProps, ITextFieldState> {
  constructor(props: IShrelloCommentProps) {
    super(props);
    this._handleChange = this._handleChange.bind(this);
    this.state = { isFocused: false };
  }

  private _handleChange = (newValue: any): void => this.props.onChanged(newValue);

  public render(): JSX.Element {
    return (
      <TextField
        multiline={true}
        borderless={!this.state.isFocused}
        onChanged={ this._handleChange }
        onFocus={()=> this.setState({isFocused: true})}
        onBlur={()=> this.setState({isFocused: false})}
        placeholder="Add a comment..."
        inputClassName={styles.shrelloCommentInput}
        className={styles.shrelloCommentField}
      />
    );
  }
}