import React from "react";

import { PersonCard, MgtPersonCard } from '@microsoft/mgt-react';
import { useGraph } from "./lib/useGraph";
import { getScopes } from "../../service/auth/AuthScope";
import { createMicrosoftGraphClient, TeamsUserCredential } from "@microsoft/teamsfx";
import { meetVsTeamsBridge } from "../../service/auth/MeetVsTeamsBridge";

MgtPersonCard.config.sections.profile = true;
MgtPersonCard.config.sections.organization = true;
MgtPersonCard.config.sections.files = false;
MgtPersonCard.config.sections.mailMessages = false;
MgtPersonCard.config.useContactApis = true;

export default class ControlableProfilCard extends React.Component {

    state = {
        //userId: "abb64646-d3e2-473d-8d13-0b6faff23605",
        userId: "",
        expanded: false
    }

    componentDidMount() {
        meetVsTeamsBridge.eventListener.on("ms-teams-change-selected-user", async data =>{
            try{
                const credential = new TeamsUserCredential();
                const graph = createMicrosoftGraphClient(credential, getScopes());
                let profile:any = await (graph.api('users/')
                .filter("mail eq '" + data.data.email + "'")
                .get());
                this.setState({ userId:profile.value[0].id  || "", expanded: false });
            }catch(err){
                this.setState({ userId: "", expanded: false });
            }
        });
    }

    render() {

        let urlParams = new URL(window.location.href).searchParams;
        let config = JSON.parse(decodeURIComponent(window.atob(urlParams.get('data') || "{}")));

        return (<>
            {
                this.state.userId && <PersonCard userId={this.state.userId} showPresence={true} isExpanded={this.state.expanded} expanded={(e) => { this.setState({ expanded: true }) }}></PersonCard>
            }</>
        )
    }
}