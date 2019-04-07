import * as React from 'react';
import styles from './Searchuser.module.scss';
import { ISearchuserProps } from './ISearchuserProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { MSGraphClient } from '@microsoft/sp-http';
import * as strings from 'SearchuserWebPartStrings';
import { IUserItem } from './IUserItem';
import ISearchUserState from './ISearchUserState';
import * as AdaptiveCards from "adaptivecards";

import {
  autobind,
  TextField
} from 'office-ui-fabric-react';


export default class Searchuser extends React.Component<ISearchuserProps, ISearchUserState> {

  constructor(props: ISearchuserProps, state: ISearchUserState) {
    super(props);
    this.state = {
      users: [],
      searchFor: ''
    };
  }

  public render(): React.ReactElement<ISearchuserProps> {
    return (
      <div className={ styles.searchuser }>
        <div className={ styles.container }>
          <TextField
              label={strings.SearchFor}
              required={true}
              value={this.state.searchFor}
              onChanged={this.searchUsers}
              onGetErrorMessage={this.searchUsersError}
            />
            {
              <div id="appendDiv" />
            }
        </div>
      </div>
    );
  }
  public componentDidUpdate() {
    var adaptiveCard = new AdaptiveCards.AdaptiveCard();
    adaptiveCard.hostConfig = new AdaptiveCards.HostConfig({
      fontFamily: "Segoe UI, Helvetica Neue, sans-serif"
      // More host config options
    });
    var componentData = [];
    var stringified = JSON.stringify(themeData);
    if (this.state != null && this.state.users != null && this.state.users.length > 0) {
      this.state.users.forEach(function (user) {
        var thisUserData = stringified.replace("{userName}", user.displayName).replace("{mail}", user.mail).replace("{userPrincipalName}", user.userPrincipalName).replace("{mobilePhone}", user.mobilePhone);

        componentData.push(JSON.parse(thisUserData));
      });
    }
    var toParse = {
      "type": "AdaptiveCard",
      "body": componentData,
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "version": "1.0"
    };

    adaptiveCard.parse(toParse);
    var renderedCard = adaptiveCard.render();
    document.getElementById("appendDiv").innerHTML = "";
    document.getElementById("appendDiv").appendChild(renderedCard);
  }
  @autobind
  private searchUsers(newValue: string): void {

    // Update the component state accordingly to the current user's input
    this.setState({
      searchFor: newValue,
    });
    this.getUsers();
  }
  private card = null;
  private searchUsersError(value: string): string {
    // The search for text cannot contain spaces
    return (value == null || value.length == 0 || value.indexOf(" ") < 0)
      ? ''
      : `${strings.SearchForValidationErrorMessage}`;
  }
  @autobind
  private getUsers(): void {

    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        client
          .api("users")
          .version("v1.0")
          .select("displayName,mail,userPrincipalName,jobTitle,mobilePhone,officeLocation,preferredLanguage,surname")
          .filter(`(startswith(displayName,'${escape(this.state.searchFor)}'))`)
          .get((err, res) => {

            if (err) {
              console.error(err);
              return;
            }

            // Prepare the output array
            var users: Array<IUserItem> = new Array<IUserItem>();

            // Map the JSON response to the output array
            if (res != null && res != undefined) {
              res.value.map((item: any) => {

                users.push({
                  displayName: item.displayName,
                  mail: item.mail,
                  userPrincipalName: item.userPrincipalName,
                  mobilePhone: item.mobilePhone
                });
              });

              // Update the component state accordingly to the result
              this.setState(
                {
                  users: users,
                }
              );
            }
          });
      });
  }
}
var themeData = {
  "type": "Container",
  "items": [
    {
      "type": "ColumnSet",
      "columns": [
        {
          "type": "Column",
          "width": "auto",
          "items": [
            {
              "type": "Image",
              "style": "Person",
              "url": "/_vti_bin/DelveApi.ashx/people/profileimage?size=L&userId={userPrincipalName}",
              "size": "Small"
            }
          ]
        },
        {
          "type": "Column",
          "width": "auto",
          "items": [
            {
              "type": "TextBlock",
              "text": "{userName}",
              "weight": "bolder",
              "size": "medium"
            }
          ]
        }
      ]
    },

    {
      "type": "ColumnSet",
      "columns": [

        {
          "type": "Column",
          "items": [
            {
              "type": "FactSet",
              "facts": [
                {
                  "title": "Mail:",
                  "value": "{mail}"
                },
                {
                  "title": "Mobile:",
                  "value": "{mobilePhone}"
                }
              ]
            }
          ],
          "width": "stretch"
        }
      ]
    }
  ]
};
