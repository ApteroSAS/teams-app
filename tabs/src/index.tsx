import React from "react";
import ReactDOM from "react-dom";
import App from "./App";
import "./index.css";
import { lisenToAuthTokenrequest } from "./service/auth/AuthScope";
import { meetVsTeamsBridge } from "./service/auth/MeetVsTeamsBridge";
lisenToAuthTokenrequest();
meetVsTeamsBridge.start();
ReactDOM.render(<App />, document.getElementById("root"));
