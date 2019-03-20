import * as React from "react";
import { IconNames } from "@uifabric/icons";
import {
  Persona,
  IPersonaProps,
  ActionButton,
  ColorClassNames,
  CompoundButton,
  IconButton,
  DefaultButton,
  PrimaryButton
} from "office-ui-fabric-react";
import { IContext } from "@pnp/spfx-controls-react/lib/common/Interfaces";
import { IPrincipal } from "@pnp/spfx-controls-react/lib/common/SPEntities";
import { Callout, ICalloutProps } from "office-ui-fabric-react/lib/Callout";
import { PeoplePicker } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import update from "immutability-helper";

import { IShrelloItem } from "../../../models";
import styles from "../ShrelloForm.module.scss";

export interface IShrelloMembersCalloutProps extends ICalloutProps {
  item?: IShrelloItem;
  buttonElement: HTMLElement;
  context: IContext;
  onUpdateMembers: (ev: React.MouseEvent<HTMLButtonElement>, item: IShrelloItem) => void;
}

export class ShrelloMembersCalloutField extends React.Component<IShrelloMembersCalloutProps, {
  isCalloutVisible: boolean,
  item?: IShrelloItem,
  requester?: IPersonaProps,
  assignedTo: IPersonaProps[],
  watching: IPersonaProps[]
}> {
  constructor(props:IShrelloMembersCalloutProps) {
    super(props);
    this.state = {
      isCalloutVisible: false,
      item: props.item,
      requester: !!props.item.Requester
        ? this._makePersona(props.item.Requester)
        : undefined,
      assignedTo: !!props.item.AssignedTo
        ? props.item.AssignedTo.map(p => this._makePersona(p))
        : [],
      watching: !!props.item.Watching
        ? props.item.Watching.map(w => this._makePersona(w))
        : []
    };
    this._onUpdateMembers = this._onUpdateMembers.bind(this);
  }

  private _onUpdateMembers = (ev: React.MouseEvent<HTMLButtonElement>) =>
    this.props.onUpdateMembers(ev, this.state.item);

  private _makePersona = (person: IPrincipal): IPersonaProps => {
    const personaProps: IPersonaProps = {
      id: person.id,
      imageUrl: person.picture,
      primaryText: person.title.split(",").reverse().join(" "),
      secondaryText: person.email,
      tertiaryText: person.jobTitle,
      onRenderSecondaryText: (): JSX.Element => {
        return (<ActionButton
          href={`mailto:${person.email}`}
          iconProps={{
            iconName: IconNames.Mail,
            className: ColorClassNames.blue
          }}
          style={{height: "auto"}}
        >
          {person.email}
        </ActionButton>);
      }
    };
    return personaProps;
  }
  private _makePrincipal = (persona: IPersonaProps): IPrincipal => {
    const principal: IPrincipal = {
      id: persona.id,
      picture: persona.imageUrl,
      title: persona.primaryText,
      email: persona.secondaryText,
      department: "",
      sip: "",
      jobTitle: persona.tertiaryText,
      value: null
    };
    return principal;
  }

  private _setRequester = (requester: any[]) => {
    const updateItem: IShrelloItem = update(this.state.item, {
      RequesterId: { $set: parseInt(requester[0].id, 10)},
      Requester: { $set: this._makePrincipal(requester[0]) }
    });
    this.setState({ item: updateItem });
  }

  private _setAssignedTo = (assignedTo: any[]) => {
    const existingId: number[] = this.state.item.AssignedToId || [];
    const existingPrincipals: IPrincipal[] = this.state.item.AssignedTo || [];
    const assignedPrincipals: IPrincipal[] = assignedTo.map(a => this._makePrincipal(a));

    const combinedIds: number[] = Array.from(new Set<number>(existingId.concat(assignedTo.map(a => parseInt(a.id, 10)))));
    const combinedPrincipals: IPrincipal[] = assignedPrincipals.concat(existingPrincipals).filter((obj, pos, arr) => {
      return arr.map(mapObj => mapObj.id).indexOf(obj.id) === pos;
    });

    const updateItem: IShrelloItem = update(this.state.item, {
      AssignedToId: { $set: combinedIds },
      AssignedTo: { $set: combinedPrincipals }
    });

    this.setState({ item: updateItem });
  }

  private _setWatching = (watching: any[]) => {
    const existingId: number[] = this.state.item.WatchingId || [];
    const existingPrincipals: IPrincipal[] = this.state.item.Watching || [];
    const assignedPrincipals: IPrincipal[] = watching.map(a => this._makePrincipal(a));

    const combinedIds: number[] = Array.from(new Set<number>(existingId.concat(watching.map(a => parseInt(a.id, 10)))));
    const combinedPrincipals: IPrincipal[] = assignedPrincipals.concat(existingPrincipals).filter((obj, pos, arr) => {
      return arr.map(mapObj => mapObj.id).indexOf(obj.id) === pos;
    });

    const updateItem: IShrelloItem = update(this.state.item, {
      WatchingId: { $set: combinedIds },
      Watching: { $set: combinedPrincipals }
    });

    this.setState({ item: updateItem });
  }

  private _persona = (person: IPrincipal): JSX.Element => {
    const personaProps: IPersonaProps = {
      imageUrl: person.picture,
      primaryText: person.title.split(",").reverse().join(" "),
      tertiaryText: person.jobTitle,
      onRenderSecondaryText: (): JSX.Element => {
        return <ActionButton
          href={`mailto:${person.email}`}
          iconProps={{
            iconName: IconNames.Mail,
            className: ColorClassNames.blue
          }}
          style={{height: "auto"}}
        >
          {person.email}
        </ActionButton>;
      }
    };
    return <Persona
      {...personaProps}
      className={styles.persona}
    />;
  }

  public render(): JSX.Element {
    const { buttonElement, onDismiss, context } = this.props;
    const { item } = this.state;

    return (
      <div>
        <Callout
          role={ "alertdialog" }
          ariaLabelledBy={ "shrello-members" }
          target={ buttonElement }
          onDismiss={ onDismiss }
          className={styles.shrelloActionCallout}
        >
          <IconButton
            iconProps={{ iconName: IconNames.ChromeClose }}
            onClick={ onDismiss }
            className={ styles.calloutDismiss }
          />
          {/* <CompoundButton
            iconProps={{iconName: IconNames.AddFriend}}
            onClick={this._onUpdateMembers}
            text="Update Members."
            description="Saves changes to members of this TASC."
            primary={true}
          /> */}
          <div className={styles.personaZone}>
            <h3>Requester:</h3>
            { !!item.Requester && item.Requester.id && this._persona(item.Requester) }
            <PeoplePicker
              context={ context }
              titleText="Requester"
              groupName="Trans.TechnicalSvcs.Members"
              personSelectionLimit={1}
              showtooltip={true}
              tooltipMessage="Change the Requester."
              selectedItems={ this._setRequester }
            />
            <hr/>
            <h3>Assigned To:</h3>
            { !!item.AssignedTo && item.AssignedTo.map(p => this._persona(p)) }
            <PeoplePicker
              context={ context }
              titleText="Assigned To"
              groupName="Trans.TechnicalSvcs.Members"
              personSelectionLimit={6}
              showtooltip={true}
              tooltipMessage="Add to the Assigned group."
              selectedItems={ this._setAssignedTo }
            />
            <hr/>
            <h3>Watching:</h3>
            { !!item.Watching && item.Watching.map(p => this._persona(p)) }
            <PeoplePicker
              context={ context }
              titleText="Watching"
              groupName="Trans.TechnicalSvcs.Members"
              personSelectionLimit={6}
              showtooltip={true}
              tooltipMessage="Add a watcher."
              selectedItems={ this._setWatching }
            />
          </div>
          <div className={ styles.shrelloFormFooter }>
            <PrimaryButton
              className={ styles.shrelloFormAction }
              onClick={ this._onUpdateMembers }
              text="Save Members"
            />
            <DefaultButton
              className={ styles.shrelloFormAction }
              onClick={ onDismiss }
              text="Cancel"
            />
          </div>
        </Callout>
      </div>
    );
  }
}