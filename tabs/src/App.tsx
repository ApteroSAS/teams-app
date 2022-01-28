import React from "react";
// https://fluentsite.z22.web.core.windows.net/quick-start
import { Provider, teamsTheme, Loader } from "@fluentui/react-northstar";
import { HashRouter as Router, Redirect, Route } from "react-router-dom";
import { useTeamsFx } from "./components/profile/lib/useTeamsFx";
import PersoTab from "./tabs/personal-tab/PersoTab";
import "./App.css";
import TabConfig from "./tabs/config-tab/TabConfig";
import SceneTab from "./tabs/scene-tab/SceneTab";

/**
 * The main app which handles the initialization and routing
 * of the app.
 */
export default function App() {
  const { theme, loading } = useTeamsFx();
  return (
    <Provider theme={theme || teamsTheme} styles={{ backgroundColor: "#eeeeee" }}>
      <Router>
        <Route exact path="/">
          <Redirect to="/perso" />
        </Route>
        {loading ? (
          <Loader style={{ margin: 100 }} />
        ) : (
          <>
            <Route exact path="/perso" component={PersoTab} />
            <Route exact path="/scene" component={SceneTab} />
            <Route exact path="/config" component={TabConfig} />
          </>
        )}
      </Router>
    </Provider>
  );
}
