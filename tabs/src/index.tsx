import React from "react";
import ReactDOM from "react-dom";
import App from "./App";
import "./index.css";
import { lisenToAuthTokenrequest } from "./service/auth/Auth";
import { meetVsTeamsBridge } from "./service/auth/MeetVsTeamsBridge";
import { deepLinkService } from "./service/DeepLinkService";
import { VERSION } from "./Version";
lisenToAuthTokenrequest();
meetVsTeamsBridge.start();
deepLinkService.start();
console.log(VERSION);

ReactDOM.render(<App />, document.getElementById("root"));
