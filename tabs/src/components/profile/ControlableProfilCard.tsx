import React from "react";

import { PersonCard, MgtPersonCard } from '@microsoft/mgt-react';
import { getScopes } from "../../service/auth/AuthScope";
import * as microsoftTeams from "@microsoft/teams-js";
import { createMicrosoftGraphClient, TeamsUserCredential } from "@microsoft/teamsfx";
import { CHANGE_USER, meetVsTeamsBridge } from "../../service/auth/MeetVsTeamsBridge";
import { Button } from "@fluentui/react-northstar";

MgtPersonCard.config.sections.profile = true;
MgtPersonCard.config.sections.organization = true;
MgtPersonCard.config.sections.files = false;
MgtPersonCard.config.sections.mailMessages = false;
MgtPersonCard.config.useContactApis = true;

export default class ControlableProfilCard extends React.Component<any,any> {

    state = {
        user: null,
        expanded: false
    }

    emailUserCache:{[id:string]:string} = {}
    uidUserCache:{[id:string]:string} = {}

    componentDidMount() {
        meetVsTeamsBridge.eventListener.on(CHANGE_USER, async data =>{
            try{
                if(data.data.uid){
                    if(!this.uidUserCache[data.data.email]){
                        const credential = new TeamsUserCredential();
                        const graph = createMicrosoftGraphClient(credential, getScopes());
                        let profile:any = await (graph.api('users/')
                        .filter("id eq '" + data.data.uid + "'")
                        .get());
                        this.uidUserCache[profile.value[0].id]=profile.value[0]
                    }
                    this.setState({ user:this.uidUserCache[data.data.uid]  || "", expanded: false });
                }else if(data.data.email){
                    if(!this.emailUserCache[data.data.email]){
                        const credential = new TeamsUserCredential();
                        const graph = createMicrosoftGraphClient(credential, getScopes());
                        let profile:any = await (graph.api('users/')
                        .filter("mail eq '" + data.data.email + "'")
                        .get());
                        this.emailUserCache[data.data.email]=profile.value[0]
                    }
                    this.setState({ user:this.emailUserCache[data.data.email]  || "", expanded: false });
                }else{
                    this.setState({ user: null, expanded: false });
                }
            }catch(err){
                this.setState({ user: null, expanded: false });
            }
        });
    }

    //https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/deep-links#deep-linking-to-a-chat
    //microsoftTeams.executeDeepLink(/*deepLink*/);
    render() {
        return (<>
            {
                this.state.user && 
                <>
                <div style={{width:"100%",display:"flex"}}>
                    {/*<a href={"https://teams.microsoft.com/l/call/0/0?users="+((this.state.user) as any).mail} target="_blank" style={{textDecoration:"none"}} >
                        <Button style={{flexGrow:1}}>Call</Button>
                    </a>
                    <a href={"https://teams.microsoft.com/l/chat/0/0?users="+((this.state.user) as any).mail} target="_blank" style={{textDecoration:"none"}} >
                        <Button style={{flexGrow:1}}>Chat</Button>
                    </a>*/}
                    <Button style={{flexGrow:1}} onClick={()=>{
                        microsoftTeams.executeDeepLink("https://teams.microsoft.com/l/call/0/0?users="+((this.state.user) as any).mail);
                        }} >Call</Button>
                    <Button style={{flexGrow:1}} onClick={()=>{
                        microsoftTeams.executeDeepLink("https://teams.microsoft.com/l/chat/0/0?users="+((this.state.user) as any).mail);
                        }} >Chat</Button>
                    <Button style={{flexGrow:1}} onClick={()=>{this.setState({ userId: "", expanded: false });}}>Close</Button>
                </div>
                <PersonCard userId={((this.state.user) as any).id} showPresence={true} isExpanded={this.state.expanded} expanded={(e) => { this.setState({ expanded: true }) }}></PersonCard>
                </>
            }</>
        )
    }
}