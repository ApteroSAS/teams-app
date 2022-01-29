import React, { useState } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import SidePanelConfigTab from "./SidePanelConfigTab";

export default function SceneTab() {
  const [frameContext, setFrameContext] = useState("loading");
  microsoftTeams.getContext((context) =>{
    console.log(context);
    // https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/access-teams-context
    //"frameContext": "The context where tab URL is loaded (for example, content, task, setting, remove, sidePanel)",
    setFrameContext(context.frameContext || "error");
  })

  let urlParams = new URL(window.location.href).searchParams;
  let config = JSON.parse(decodeURIComponent(window.atob( urlParams.get('data') || "{}" )));

  return (
    <>
    {frameContext==="loading" && <div>loading...</div>}
    {frameContext==="sidePanel" && <SidePanelConfigTab/>}
    {frameContext==="content" && 
        <iframe 
        title="scene"
        id="mainframe" 
        src={config.url}
        loading="eager" 
        style={{"width": "100%", "height": "100%","margin":0,"padding": 0,"border":0}}></iframe>}
    </>
  );
}
