import React from "react";
import { GraphFileBrowser } from "@microsoft/file-browser";

import moment from "moment";
import config from "./Config";
import { getFileData } from "./GraphService";

function formatDateTime(dateTime) {
  return moment
    .utc(dateTime)
    .local()
    .format("M/D/YY h:mm A");
}

export default class Files extends React.Component {
  constructor(props) {
    super(props);

    this.state = {
      events: []
    };
  }

  async componentDidMount() {
    try {
      // Get the user's access token
      this.accessToken = await window.msal.acquireTokenSilent({
        scopes: config.scopes
      });
      // Get the user's events
      var fileData = await getFileData(this.accessToken);
      // Update the array of events in state
      this.setState({ fileData: fileData.value });
    } catch (err) {
      this.props.showError("ERROR", JSON.stringify(err));
    }
  }

  render() {
    return (
      <div>
        <h1>Files</h1>
        <GraphFileBrowser
          getAuthenticationToken={this.getAuthenticationToken}
        />
      </div>
    );
  }

  async getAuthenticationToken() {
    var accessTokenRequest = { scopes: config.scopes };
    try {
      // Get the user's access token
      var accessToken = await window.msal
        .acquireTokenSilent(accessTokenRequest)
        .then(function(accessTokenResponse) {
          // Acquire token silent success
          // Call API with token
          return accessTokenResponse.accessToken;
        });
      return accessToken;
    } catch (err) {
      //Acquire token silent failure, and send an interactive request
      var accessToken = await window.msal
        .acquireTokenPopup(accessTokenRequest)
        .then(function(accessTokenResponse) {
          // Acquire token interactive success
          return accessTokenResponse.accessToken;
        });
      //console.log(error);
    }
  }
}
