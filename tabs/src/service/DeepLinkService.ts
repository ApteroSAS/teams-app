import { meetVsTeamsBridge, PROCESS_DEEPLINK } from "./auth/MeetVsTeamsBridge";

export class DeepLinkService{
    start(){
        meetVsTeamsBridge.eventListener.on(PROCESS_DEEPLINK,(data)=>{
            const link = data.link;
            microsoftTeams.executeDeepLink(link);
        })
    }
}

export const deepLinkService = new DeepLinkService();