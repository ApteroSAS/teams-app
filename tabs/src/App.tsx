import React from "react";
// https://fluentsite.z22.web.core.windows.net/quick-start
import { Provider, teamsTheme, Loader } from "@fluentui/react-northstar";
import { HashRouter as Router, Redirect, Route } from "react-router-dom";
import { useTeamsFx } from "./components/profile/lib/useTeamsFx";
import "./App.css";
import TabConfig from "./tabs/config-tab/TabConfig";
import SceneTab from "./tabs/scene-tab/SceneTab";
import { TeamsUserCredential } from "@microsoft/teamsfx";
import { getScopes } from "./components/auth/AuthScope";

/**
 * The main app which handles the initialization and routing
 * of the app.
 */
export default function App() {
  const { theme, loading } = useTeamsFx();
  /*if(!loading){
    const credential = new TeamsUserCredential();
    try{
      credential.login(getScopes());
    }catch(err){
      // silent fail
    }
  }*/
  return (
    <Provider theme={theme || teamsTheme} styles={{ backgroundColor: "#eeeeee" }}>
      <Router>
        {loading ? (
          <Loader style={{ margin: 100 }} />
        ) : (
          <>
            <Route exact path="/scene" component={SceneTab} />
            <Route exact path="/config" component={TabConfig} />
          </>
        )}
      </Router>
    </Provider>
  );
}
