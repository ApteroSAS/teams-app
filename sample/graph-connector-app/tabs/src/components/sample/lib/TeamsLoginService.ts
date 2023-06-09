import {
    AuthenticationConfiguration,
    createMicrosoftGraphClient, createMicrosoftGraphClientWithCredential,
    IdentityType,
    LogLevel,
    setLogFunction,
    setLogLevel,
    TeamsFx, TeamsUserCredential
} from "@microsoft/teamsfx";
import {MicrosoftGraphService} from "./MicrosoftGraphService";
import {Client, GraphError} from "@microsoft/microsoft-graph-client";

export const startLoginPageUrl = process.env.REACT_APP_START_LOGIN_PAGE_URL; // window.location.origin + "/auth-start.html";
export const MICROSOFT_APP_AUTHORITY = "https://login.microsoftonline.com/apteroco.onmicrosoft.com";
export const MICROSOFT_APP_ID = process.env.REACT_APP_CLIENT_ID;

declare const window : any;
//To test and revoke application auth
// https://myapplications.microsoft.com/
// https://portal.office.com/account/?ref=MeControl#apps

export class TeamsLoginService {
    teamsFx: TeamsFx |null = null;
    teamsUserCredential: TeamsUserCredential |null = null;

    constructor(
        private microsoftGraphService: MicrosoftGraphService,
    ) {
        console.info("app id : " + MICROSOFT_APP_ID + " / " + MICROSOFT_APP_AUTHORITY);
        console.info("auth page : " + startLoginPageUrl);
    }

    //////////////
    /////TeamsFx
    ////////////
    private getTeamsFx(): TeamsFx {
        if (!this.teamsFx) {
            if (process.env.NODE_ENV === "development") {
                setLogLevel(LogLevel.Warn);
                setLogFunction((leve, message) => {
                    console.warn(message);
                });
            }
            //https://docs.microsoft.com/en-us/microsoftteams/platform/toolkit/teamsfx-sdk
            this.teamsFx = new TeamsFx(IdentityType.User, {
                authorityHost:MICROSOFT_APP_AUTHORITY,
                initiateLoginEndpoint: startLoginPageUrl,
                //simpleAuthEndpoint: teamsfxEndpoint,//https://github.com/OfficeDev/TeamsFx/wiki/How-Authentication-Works-in-TeamsFx-Tab-Template
                clientId: MICROSOFT_APP_ID, //process.env.REACT_APP_CLIENT_ID;
            });
        }
        return this.teamsFx;
    }

    private async triggerScopeAuthPopUpTeamsFx(scopes:string[]): Promise<TeamsFx> {
        const teamsFx = this.getTeamsFx();
        console.log("login TeamsFx :" , scopes);
        await teamsFx.login(scopes); //Trigger authorization popup
        let credential = teamsFx.getCredential();//checks credential
        console.log(await credential.getToken(scopes));//checks credential
        return teamsFx;
    }

    private async initGraphTeamsFx(scope: string[], silent: boolean = true): Promise<{ graph: Client, scope: string[] }> {
        const teamsFx = this.getTeamsFx();
        try {
            const graph = createMicrosoftGraphClient(teamsFx, scope);
            await this.microsoftGraphService.init(graph);
            return {graph, scope};
        } catch (err:any) {
            if (err.toString().indexOf("please login first") !== -1) {
                console.info("app not authorized - no graph support - re-login");
                console.info("app id : " + MICROSOFT_APP_ID + " / " + MICROSOFT_APP_AUTHORITY);
                console.info("auth page : " + startLoginPageUrl);
                if (!silent) {
                    await this.triggerScopeAuthPopUpTeamsFx(scope);//trigger login popup
                    const graph = createMicrosoftGraphClient(teamsFx, scope);
                    await this.microsoftGraphService.init(graph);
                    return {graph, scope};
                }else {
                    throw err;
                }
            } else {
                throw err;
            }
        }
    }
    //////////////
    /////TeamsUserCredential
    ////////////
    private getTeamsUserCredential(): TeamsUserCredential {
        if (!this.teamsUserCredential) {
            const authConfig: AuthenticationConfiguration = {
                clientId: MICROSOFT_APP_ID,
                initiateLoginEndpoint: startLoginPageUrl,
                authorityHost:MICROSOFT_APP_AUTHORITY,
            };
            this.teamsUserCredential = new TeamsUserCredential(authConfig);
        }
        return this.teamsUserCredential;
    }

    private async triggerScopeAuthPopUpTeamsUserCredential(scopes:string[]): Promise<TeamsUserCredential> {
        const teamsUserCredential = this.getTeamsUserCredential();
        console.log("login teamsUserCredential :" , scopes);
        await teamsUserCredential.login(scopes); //Trigger authorization popup
        console.log(await teamsUserCredential.getToken(scopes));//checks credential
        return teamsUserCredential;
    }
    async initGraphTeamsUserCredential<T>(scope: string[], silent: boolean = true): Promise<{ graph: Client, scope: string[] }> {
        let credential: TeamsUserCredential = this.getTeamsUserCredential();
        try {
            const graph = createMicrosoftGraphClientWithCredential(credential, scope);
            await this.microsoftGraphService.init(graph);
            return {graph, scope};
        } catch (err: unknown) {
            console.info("app not authorized - no graph support - re-login");
            console.info("scopes : " + scope);
            console.info("app id : " + MICROSOFT_APP_ID + " / " + MICROSOFT_APP_AUTHORITY);
            console.info("auth page : " + startLoginPageUrl);//you need to change the Appid if you are using localhost
            if (!silent) {
                // Silently fail for user didn't consent error
                await this.triggerScopeAuthPopUpTeamsUserCredential(scope);
                const graph = createMicrosoftGraphClientWithCredential(credential, scope);
                await this.microsoftGraphService.init(graph);
                return {graph, scope};
            } else {
                throw err;
            }
        }
    }

}
