import * as React from "react";
import {
  DefaultButton,
  ButtonType
 } from "office-ui-fabric-react/lib/Button";
 import { Icon } from "office-ui-fabric-react/lib/Icon";
 import styles from "./ConfigurationView.module.scss";
 import { IConfigurationViewProps, IConfigurationViewState } from "./IConfigurationView";

 export default class ConfigurationView extends React.Component<IConfigurationViewProps, IConfigurationViewState> {

  constructor(props: IConfigurationViewProps) {
     super(props);

     this.state = {
       inputValue: ""
     };

     this._handleConfigureButtonClick = this._handleConfigureButtonClick.bind(this);
   }

   private _handleConfigureButtonClick(event?: React.MouseEvent<HTMLButtonElement>): void {
    this.props.onConfigure();
   }

   public render(): JSX.Element {
     return (
      <div className="Placeholder">
      <div className="Placeholder-container ms-Grid">
          <div className="Placeholder-head ms-Grid-row">
              <div className="ms-Grid-col ms-hiddenSm ms-md3"></div>
              <div className="Placeholder-headContainer ms-Grid-col ms-sm12 ms-md6">
                <Icon
                  iconName={ this.props.icon }
                  className="Placeholder-icon ms-fontSize-su ms-Icon"
                />
                <span className="Placeholder-text ms-fontWeight-light ms-fontSize-xxl">{this.props.iconText}</span>
              </div>                
              <div className="ms-Grid-col ms-hiddenSm ms-md3"></div>
          </div>
          <div className="Placeholder-description ms-Grid-row">
            <span className="Placeholder-descriptionText">{this.props.description}</span>
          </div>
          <div className="Placeholder-description ms-Grid-row">
              <DefaultButton
                className={ styles.configureButton }
                buttonType={ ButtonType.primary }
                ariaLabel={ this.props.buttonLabel }
                onClick={this._handleConfigureButtonClick}>
                {this.props.buttonLabel}
              </DefaultButton>
          </div>
      </div>
  </div>
     );
   }
 }