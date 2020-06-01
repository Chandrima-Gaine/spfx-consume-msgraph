import * as React from 'react';
import styles from './ConsumeMsGraph.module.scss';
import { IConsumeMsGraphProps } from './IConsumeMsGraphProps';
import { IConsumeMsGraphState } from './IConsumeMsGraphState';
import { escape } from '@microsoft/sp-lodash-subset';
import {  MSGraphClient } from "@microsoft/sp-http";
import { IUserItem } from './IUserItem';

import {
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode
} from 'office-ui-fabric-react';

// Configure the columns for the DetailsList component
let _usersListColumns = [
  {
    key: 'displayName',
    name: 'Display name',
    fieldName: 'displayName',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'mail',
    name: 'Mail',
    fieldName: 'mail',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'userPrincipalName',
    name: 'User Principal Name',
    fieldName: 'userPrincipalName',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true
  },
];

export default class ConsumeMsGraph extends React.Component<IConsumeMsGraphProps, IConsumeMsGraphState> {
  constructor(props: IConsumeMsGraphProps, state: IConsumeMsGraphState) {
    super(props);
    
    // Initialize the state of the component
    this.state = {
      users: []
    };
  }
    public render(): React.ReactElement<IConsumeMsGraphProps> {
      return (
        <div className={ styles.consumeMsGraph }>
          <div className={ styles.container }>
            <div className={ styles.row }>
              <div className={ styles.column }>
                <span className={ styles.title }>Microsoft Graph</span>
                <p className={ styles.subTitle }>Consume MS Graph with SPFX...</p>
                <p className={ styles.description }>Click the button to get User details...</p>
                <p className={ styles.form }>
                  <PrimaryButton 
                    text='All User Details' 
                    title='All User Details' 
                    onClick={ this.getUserDetails } 
                  />
                </p>
                {
                (this.state.users != null && this.state.users.length > 0) ?
                  <p className={ styles.form }>
                    <DetailsList
                        items={ this.state.users }
                        columns={ _usersListColumns }
                        setKey='set'
                        checkboxVisibility={ CheckboxVisibility.hidden }
                        selectionMode={ SelectionMode.none }
                        layoutMode={ DetailsListLayoutMode.fixedColumns }
                        compact={ true }
                    />
                  </p>
                  : null
                }
              </div>
            </div>
          </div>
        </div>
      );
    }
  
    private getUserDetails = () : void => {

      // Log the current operation
      console.log("Using getUserDetails() method");
    
      this.props.context.msGraphClientFactory
        .getClient()
        .then((graphClient: MSGraphClient) => {
          graphClient
            .api("users")
            .version("v1.0")
            .select("displayName,mail,userPrincipalName")
            .get((err, res) => {
    
              if (err) {
                console.error(err);
                return;
              }
    
              // Prepare the output array
              var users: Array<IUserItem> = new Array<IUserItem>();
    
              // Map the JSON response to the output array
              res.value.map((item: any) => {
                users.push( {
                  displayName: item.displayName,
                  mail: item.mail,
                  userPrincipalName: item.userPrincipalName,
                });
              });
    
              // Update the component state accordingly to the result
              this.setState(
                {
                  users: users,
                }
              );
            });
        });
      }
}
