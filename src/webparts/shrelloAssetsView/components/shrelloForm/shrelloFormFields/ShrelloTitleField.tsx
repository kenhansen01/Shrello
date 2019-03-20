import * as React from "react";
import {
  ITextFieldState,
  ITextFieldProps,
  TextField
} from "office-ui-fabric-react/lib/TextField";

// import { IShrelloItem } from "../../../models";
import styles from "../ShrelloForm.module.scss";

export interface IShrelloTitleProps extends ITextFieldProps {
  title: string;
}

export default class ShrelloTitleField extends React.Component<IShrelloTitleProps, ITextFieldState> {
  constructor(props: IShrelloTitleProps) {
    super(props);
    this._handleChange = this._handleChange.bind(this);
    this.state = { isFocused: false };
  }

  private _handleChange = (newValue: any): void => this.props.onChanged(newValue);

  public componentDidMount(): void {
    this.setState({ isFocused: false });
  }

  public render(): JSX.Element {
    const { title } = this.props;
    return (
      <TextField
        label=""
        required={ true }
        value={ title }
        placeholder="Enter a title..."
        borderless={!this.state.isFocused}
        onChanged={ this._handleChange }
        onFocus={()=> this.setState({isFocused: true})}
        onBlur={()=> this.setState({isFocused: false})}
        className={styles.shrelloTitle}
        inputClassName={styles.shrelloTitleInput}
      />
    );
  }
}