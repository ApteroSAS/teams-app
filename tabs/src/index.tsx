import React from "react";
import ReactDOM from "react-dom";
import App from "./App";
import "./index.css";
import { lisenToAuthTokenrequest } from "./service/auth/Auth";
import { meetVsTeamsBridge } from "./service/auth/MeetVsTeamsBridge";
import { deepLinkService } from "./service/DeepLinkService";
lisenToAuthTokenrequest();
meetVsTeamsBridge.start();
deepLinkService.start();

ReactDOM.render(<App />, document.getElementById("root"));
