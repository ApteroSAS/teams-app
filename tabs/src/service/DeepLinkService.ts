import { meetVsTeamsBridge, PROCESS_DEEPLINK } from "./auth/MeetVsTeamsBridge";
import * as microsoftTeams from "@microsoft/teams-js";

export class DeepLinkService{
    start(){
        meetVsTeamsBridge.eventListener.on(PROCESS_DEEPLINK,(data)=>{
            try{
                const link = data.data.link;
                microsoftTeams.executeDeepLink(link,(status,reason)=>{
                    console.log(status,reason);
                });
            }catch(err){
                console.error(err);
            }
        })
    }
}

export const deepLinkService = new DeepLinkService();