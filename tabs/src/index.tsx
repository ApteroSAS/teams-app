import React from "react";
import ReactDOM from "react-dom";
import App from "./App";
import "./index.css";

import * as microsoftTeams from "@microsoft/teams-js";
import {ProviderState, SimpleProvider, TeamsProvider} from '@microsoft/mgt';

import { Providers } from '@microsoft/mgt';
import { TeamsUserCredential } from "@microsoft/teamsfx";
import { getScopes } from "./components/auth/AuthScope";

TeamsProvider.microsoftTeamsLib = microsoftTeams;


async function getAccessToken(scopes: string[]):Promise<string> {
    const credential = new TeamsUserCredential();
    return (await credential.getToken(getScopes()))?.token || "ERROR TOKEN";
}
  
async function login() {
    //login code
    Providers.globalProvider.setState(ProviderState.SignedIn)
    console.log("login")
}

async function logout() {
    // logout code
    console.log("logout")
}
  
Providers.globalProvider = new SimpleProvider(getAccessToken, login, logout);

ReactDOM.render(<App />, document.getElementById("root"));
