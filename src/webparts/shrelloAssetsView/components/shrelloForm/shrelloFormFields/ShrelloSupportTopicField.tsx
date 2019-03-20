import * as React from "react";
import {
  IDropdownProps,
  IDropdownState,
  Dropdown,
  IDropdownOption,
  DropdownMenuItemType
} from "office-ui-fabric-react/lib/Dropdown";

import {IShrelloItem, ISupportTopicItem} from "../../../models";

export interface IShrelloSupportTopicProps extends IDropdownProps {
  item: IShrelloItem;
  supportTopics: ISupportTopicItem[];
  selectedDepartment?: string;
  choices?: IDropdownOption[];
}

export interface IShrelloSupportTopicState extends IDropdownState {
  choices: IDropdownOption[];
  selectedDepartment: string;
  selectedTopicKey?: string;
}

export default class ShrelloSupportTopicField extends React.Component<IShrelloSupportTopicProps, IShrelloSupportTopicState> {
  private _choices: IDropdownOption[] = [];

  constructor(props: IShrelloSupportTopicProps) {
    super(props);

    this.state = {
      choices: this._choices,
      selectedDepartment: props.selectedDepartment
    };

    this._handleChange = this._handleChange.bind(this);
  }

  private _makeSupportOptions = (topics: ISupportTopicItem[], selectedDepartment: string): IDropdownOption[] => {
    let supportTopicOptions: IDropdownOption[] = [];
    supportTopicOptions.push({
      key: "Header",
      text: selectedDepartment,
      itemType: DropdownMenuItemType.Header
    });

    topics.map(topic => {
      const dOption: IDropdownOption = {
        key: String(topic.Id),
        text: topic.Title
      };
      supportTopicOptions.push(dOption);
    });

    return supportTopicOptions;
  }

  private _handleChange = (option: IDropdownOption): void => this.props.onChanged(option);

  public componentDidMount(): void {
    this._choices = this._makeSupportOptions(this.props.supportTopics, this.props.selectedDepartment);
    const selectedChoice: IDropdownOption = this._choices.find(topic =>
      topic.key as string === String(this.props.item.SupportTopicId)) || this._choices[1];
    const selectedKey: string = selectedChoice.key as string;
    this.setState({
      choices: this._choices,
      selectedTopicKey: selectedKey
    });
    this._handleChange(selectedChoice);
  }

  public componentWillReceiveProps(props: IShrelloSupportTopicProps): void {
    let choicesUpdated: boolean = false;
    if (this.state.selectedDepartment !== props.selectedDepartment) {
      this.setState({ selectedDepartment: props.selectedDepartment});
    }
    this._choices = this._makeSupportOptions(props.supportTopics, props.selectedDepartment);
    choicesUpdated = this.state.choices.some((choice, idx) => (choice.key !== this._choices[idx].key));

    const selectedChoice: IDropdownOption = this._choices.find(topic =>
      topic.key as string === String(this.props.item.SupportTopicId)) || this._choices[1];
    const selectedKey: string = selectedChoice.key as string;
    if (choicesUpdated) {
      this.setState({ choices: this._choices });
    }
    if (this.state.selectedTopicKey !== selectedKey) {
      this.setState({ selectedTopicKey: selectedKey });
      this._handleChange(selectedChoice);
    }
  }

  public render(): JSX.Element {
    const { choices, selectedTopicKey } = this.state;

    return (
      <Dropdown
        label="Support Topic:"
        id="SupportTopicSelect"
        options={ choices }
        selectedKey={ selectedTopicKey }
        onChanged={ this._handleChange }
      />
    );
  }
}