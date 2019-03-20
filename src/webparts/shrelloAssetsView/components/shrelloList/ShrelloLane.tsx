import * as React from "react";
import { Link } from "office-ui-fabric-react";
import ShrelloCard from "./ShrelloCard";
import { IShrelloItem, ISupportDepartmentItem } from "../../models";
import styles from "./ShrelloList.module.scss";

export interface IShrelloLaneProps {
  title: string;
  itemColumn?: string;
  laneValue?: string;
  items?: IShrelloItem[];
  supportDepartments?: ISupportDepartmentItem[];
  laneDepartment?: ISupportDepartmentItem;
  handleItemClick?: (ev: React.MouseEvent<HTMLElement>, item: IShrelloItem) => void;
  onSelectTeamView?: (ev: React.MouseEvent<HTMLElement>, laneDepartment: ISupportDepartmentItem) => void;
}

export interface IShrelloLaneState {
  items: IShrelloItem[];
}

function ShrelloCardContainer(props): JSX.Element {
  return (
    <div className={ styles.shrelloLane }>
      <div className={ styles.shrelloLaneContent}>
        <div className={ styles.shrelloLaneHeader}>
          <div className={ styles.shrelloLaneHeaderTarget}></div>
          {
            props.laneDepartment &&
            <Link onClick={ props.onSelectTeamView }>
              <h2>{ props.title }</h2>
            </Link>
          }
          {
            !props.laneDepartment &&
            <h2>{ props.title }</h2>
          }          
        </div>
        <div className={ styles.shrelloLaneCards }>
          { props.children }
        </div>
      </div>
    </div>
  );
}

function ShrelloCards(props): JSX.Element {
  const { supportDepartments, items, handleItemClick } = props;
  const itemCards: JSX.Element[] = items.map(item => <ShrelloCard {...props}
    item={ item }
    supportDepartment={ supportDepartments.find(dept => dept.Id === item.SupportDepartmentId) }
    handleClick={ handleItemClick }
  />);
  return (
    <ShrelloCardContainer {...props}>
      { itemCards }
    </ShrelloCardContainer>
  );
}

export default class ShrelloLane extends React.Component<IShrelloLaneProps, IShrelloLaneState> {

  constructor(props:IShrelloLaneProps) {
    super(props);
    this.state = {
      items: props.items || []
    };
    this._handleItemClick = this._handleItemClick.bind(this);
    this._handleSupportTeamSelect = this._handleSupportTeamSelect.bind(this);
  }

  public componentWillReceiveProps(props: IShrelloLaneProps): void {
    this.setState({
      items: props.items || []
    });
  }

  public componentDidMount(): void {
    this.setState({
      items: this.props.items || []
    });
  }

  private _handleItemClick = (ev: React.MouseEvent<HTMLElement>, item: IShrelloItem): void => {
    this.props.handleItemClick(ev, item);
  }

  private _handleSupportTeamSelect = (ev: React.MouseEvent<HTMLElement>): void => {
    ev.preventDefault();
    this.props.onSelectTeamView(ev, this.props.laneDepartment);
  }

  public render(): JSX.Element {
    return (
      <ShrelloCards
        {...this.props}
        {...this.state}
        supportDepartments={ this.props.supportDepartments }
        handleItemClick={ this._handleItemClick }
        onSelectTeamView={ this._handleSupportTeamSelect }
      />
    );
  }
}