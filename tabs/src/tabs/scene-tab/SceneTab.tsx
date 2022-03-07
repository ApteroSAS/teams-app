import React from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import SidePanelConfigTab from "./SidePanelConfigTab";


import ControlableProfilCard from "../../components/profile/ControlableProfilCard";
import { initAuth } from "../../service/auth/AuthScope";
import { Login } from "@microsoft/mgt-react";
import { meetVsTeamsBridge } from "../../service/auth/MeetVsTeamsBridge";


export default class SceneTab extends React.Component {
  
  state = {
    frameContext: "loading",
    auth: false,
  }
  scope: string[] = [];

  async componentDidMount() {
    await initAuth(()=>{
        console.error("consent needed");
    });
    
    microsoftTeams.getContext((context) =>{
      console.log(context);
      // https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/access-teams-context
      //"frameContext": "The context where tab URL is loaded (for example, content, task, setting, remove, sidePanel)",
      this.setState({frameContext:context.frameContext || "error"});
    })
    
  }
 

  render() {

    let urlParams = new URL(window.location.href).searchParams;
    let config = JSON.parse(decodeURIComponent(window.atob( urlParams.get('data') || "{}" )));

    return (
      <>
      <div style={{position:"absolute", top:"20px",left:"20px",backgroundColor:"white"}}>
        {/*<AccountAuth />*/}
        <Login loginCompleted={()=> setTimeout(()=>this.setState({auth:true}),2000)}/>
        {this.state.auth && <ControlableProfilCard/>}
      </div>
      {this.state.frameContext==="loading" && <div>loading...</div>}
      {this.state.frameContext==="sidePanel" && <SidePanelConfigTab/>}
      {(this.state.frameContext!=="loading" && this.state.frameContext!=="sidePanel") && 
          <iframe 
          title="scene"
          id="mainframe" 
          ref={(ref)=>{
            meetVsTeamsBridge.setIframe(ref);
          }}
          src={config.url}
          loading="eager" 
          style={{"width": "100%", "height": "100%","margin":0,"padding": 0,"border":0}}></iframe>}
      </>
    );
  }
}
