import React from "react";
import * as microsoftTeams from "@microsoft/teams-js";

export default function SceneTab() {
  microsoftTeams.getContext((context) =>{
    console.log(context);
  })

  let urlParams = new URL(window.location.href).searchParams;
  let config = JSON.parse(decodeURIComponent(window.atob( urlParams.get('data') || "{}" )));

  return (
        <iframe 
        title="scene"
        id="mainframe" 
        src={config.url}
        loading="eager" 
        style={{"width": "100%", "height": "100%","margin":0,"padding": 0,"border":0}}></iframe>
  );
}
