import * as React from "react";
import { css, Icon, TooltipHost } from "office-ui-fabric-react";
import { IconNames } from "office-ui-fabric-react/lib/Icons";

import {
  IShrelloItem,
  ISupportDepartmentItem
} from "../../models";
import {
  Priorities,
  IShrelloLabel,
  Categories,
  Statuses
} from "../../config/shrelloViewConfig";
import styles from "./ShrelloList.module.scss";

export interface IShrelloCardProps {
  item: IShrelloItem;
  supportDepartment: ISupportDepartmentItem;
  handleClick?: (ev: React.MouseEvent<HTMLElement>, item: IShrelloItem) => void;
}

export interface IShrelloCardState {
  item: IShrelloItem;
}

export default class ShrelloCard extends React.Component<IShrelloCardProps, IShrelloCardState> {

  constructor(props: IShrelloCardProps) {
    super(props);
    this.state = {
      item: props.item
    };
    this._handleClick = this._handleClick.bind(this);
  }

  private _handleClick = (ev: React.MouseEvent<HTMLElement>): void => {
    this.props.handleClick(ev, this.state.item);
  }

  public componentWillReceiveProps(props: IShrelloCardProps): void {
    this.setState({
      item: props.item
    });
  }

  public render(): JSX.Element {
    const { item } = this.state;
    const priority: IShrelloLabel = Priorities.find(p => p.name === item.Priority);
    const category: IShrelloLabel = Categories.find(c => c.name === item.TASCTypeCategory);
    const status: IShrelloLabel = Statuses.find(s => s.name === item.Status);
    return (
      <div className={ styles.shrelloCard } onClick={ this._handleClick }>
        <div className={ styles.shrelloCardDetails }>
          <div className={ styles.shrelloCardLabels }>
            { item.SupportDepartmentId && this.props.supportDepartment &&
              <span className={ styles.cardLabel }>
                <TooltipHost content={ this.props.supportDepartment.Title } calloutProps={{ gapSpace: 0 }}>
                  <Icon iconName={ this.props.supportDepartment.IconName }/>
                </TooltipHost>
              </span>
            }
            { item.Priority && priority &&
              <span className={ styles.cardLabel }>
                <TooltipHost content={ priority.name }>
                  <Icon
                    iconName={ priority.iconName }
                    className={ priority.iconColor }
                  />
                </TooltipHost>
              </span>
            }
            { item.Status && status &&
              <span className={ styles.cardLabel }>
                <TooltipHost content={ status.name }>
                  <Icon
                    iconName={ status.iconName }
                    className={ status.iconColor }
                  />
                </TooltipHost>
              </span>
            }
            { item.TASCTypeCategory && category &&
              <span>
                <TooltipHost content={ category.name }>
                  <Icon
                    iconName={ category.iconName }
                    className={ category.iconColor }
                  />
                </TooltipHost>
              </span>
            }
          </div>
          <span className={ styles.shrelloCardTitle}>
            {item.Title}
          </span>
          <div className={ styles.shrelloCardBadges }>
            { item.Body &&
              <div className={ styles.shrelloCardBadge }>
                <Icon iconName={ IconNames.List } className={ css(styles.shrelloCardBadgeIcon, styles.shrelloCardBadgeIconSmall) } />
              </div>
            }
            { item.Attachments && item.AttachmentFiles.length &&
              <div className={ styles.shrelloCardBadge }>
                <Icon iconName={ IconNames.Attach } className={ css(styles.shrelloCardBadgeIcon, styles.shrelloCardBadgeIconSmall) } />
                <span className={ styles.shrelloCardBadgeText }>{ item.AttachmentFiles.length.toString() }</span>
              </div>
            }
          </div>
        </div>
      </div>
    );
  }
}