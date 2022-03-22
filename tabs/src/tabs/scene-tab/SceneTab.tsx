import React from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import SidePanelConfigTab from "./SidePanelConfigTab";
import { meetVsTeamsBridge } from "../../service/auth/MeetVsTeamsBridge";
import { VERSION } from "../../Version";
import { AuthorizePopUp } from "../../components/profile/AuthorizePopUp";
import { Button, Popup } from "@fluentui/react-northstar";


export default class SceneTab extends React.Component {
  
  state = {
    frameContext: "loading"
  }
  scope: string[] = [];

  async componentDidMount() {
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
    let frameUrl = new URL(config.url)
    frameUrl.searchParams.set("verion", VERSION);//add a parameters to refresh cache on new version

    return (
      <>
      <div style={{position:"absolute", top:"70px",left:"50%",transform: "translate(-50%, -50%)",backgroundColor:"white"}}>
        <AuthorizePopUp />
      </div>
      {/*<div style={{position:"absolute", bottom:"20px",left:"20px",backgroundColor:"white"}}>
        <AccountAuth />
      </div>
        <div style={{position:"absolute", top:"20px",left:"20px",backgroundColor:"white"}}>
          <ControlableProfilCard/>
      </div>*/}
      {this.state.frameContext==="loading" && <div>loading...</div>}
      {this.state.frameContext==="sidePanel" && <SidePanelConfigTab/>}
      {(this.state.frameContext!=="loading" && this.state.frameContext!=="sidePanel") && 
          <iframe 
          title="scene"
          id="mainframe" 
          ref={(ref)=>{
            meetVsTeamsBridge.setIframe(ref);
          }}
          src={frameUrl.toString()}
          loading="eager" 
          style={{"width": "100%", "height": "100%","margin":0,"padding": 0,"border":0}}></iframe>}
      </>
    );
  }
}
