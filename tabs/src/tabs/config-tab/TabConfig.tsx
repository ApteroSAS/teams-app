import React from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { AccountAuth } from "../../components/profile/AccountAuth";
import {Input} from "@fluentui/react-northstar";//https://fluentsite.z22.web.core.windows.net/0.60.1
import { VERSION } from "../../Version";

/**
 * The 'Config' component is used to display your group tabs
 * user configuration options.  Here you will allow the user to
 * make their choices and once they are done you will need to validate
 * their choices and communicate that to Teams to enable the save button.
 */
class TabConfig extends React.Component {
  state={
    //url:"https://meet.aptero.co/3NKS9NY?embed_token=ae766e3d6bb6394d2aa38f69d5f1aad9&adapter=none"
    url:"https://meet.aptero.co/3S7yY8X/kpmg-eqho?embed_token=ed66cbb7892901957af4a85bd30f2cd1&adapter=none"
  }

  
  render() {
    // Initialize the Microsoft Teams SDK
    microsoftTeams.initialize();

    /**
     * When the user clicks "Save", save the url for your configured tab.
     * This allows for the addition of query string parameters based on
     * the settings selected by the user.
     */
    microsoftTeams.settings.registerOnSaveHandler((saveEvent) => {
      let config = {...this.state};
      const roomurl = new URL(config.url);
      roomurl.searchParams.append("msteams","true");
      config.url = roomurl.toString();
      const atobStr = btoa(encodeURIComponent(JSON.stringify(config)));
      const baseUrl = `https://${window.location.hostname}:${window.location.port}`;
      const url = new URL(baseUrl + "/index.html#/scene");
      url.searchParams.append("data",atobStr);
      microsoftTeams.settings.setSettings({
        suggestedDisplayName: "Virtual Meeting",
        entityId: "sceneConfig",
        contentUrl: url.toString(),
        websiteUrl: this.state.url,
      });
      saveEvent.notifySuccess();
    });

    /**
     * After verifying that the settings for your tab are correctly
     * filled in by the user you need to set the state of the dialog
     * to be valid.  This will enable the save button in the configuration
     * dialog.
     */
    microsoftTeams.settings.setValidityState(true);

    return (
      <div>
        <h1>Configuration</h1>
        <h2>Identity</h2>
        <AccountAuth />
        <h2>Room</h2>
        <Input
            label="Room URL"
            fluid
            value={this.state.url}
            onChange={(value)=>{
              this.setState({url:value})
            }}
        />
        <span id="version" style={{"margin":0,"top":"auto","left":"auto","right":"4px","bottom":"2px","fontSize":"0.7em","position":"fixed"}}>{VERSION}</span>
      </div>
    );
  }
}

export default TabConfig;
