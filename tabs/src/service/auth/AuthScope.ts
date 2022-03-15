import { Providers, ProviderState } from '@microsoft/mgt-react';
import { TeamsFxProvider } from "@microsoft/mgt-teamsfx-provider";
import * as microsoftTeams from "@microsoft/teams-js";
import { loadConfiguration, TeamsUserCredential, ResourceType, LogLevel, setLogLevel, setLogFunction } from "@microsoft/teamsfx";
import { AUTH_TOKEN, meetVsTeamsBridge, RELOAD, REQUEST_AUTH_TOKEN, REQUEST_USER_INFO, USER_INFO } from './MeetVsTeamsBridge';
//https://github.com/OfficeDev/TeamsFx/wiki/How-Authentication-Works-in-TeamsFx-Tab-Template

let credential: TeamsUserCredential | null; //Due to the limitation of msal, multiple instances of TeamsUserCredential is not recommended.
var teamsfxEndpoint = process.env.REACT_APP_TEAMSFX_ENDPOINT;
var startLoginPageUrl = process.env.REACT_APP_START_LOGIN_PAGE_URL;
var functionEndpoint = process.env.REACT_APP_FUNC_ENDPOINT;
var clientId = process.env.REACT_APP_CLIENT_ID;

export function getCredential() : TeamsUserCredential{
    if(!credential){
        initTeamsFx("common");
    }
    if(!credential){
        throw new Error("");
    }
    return credential;
}

export function getScopes(): string[] {
    //const neededScopes = MgtPersonCard.getScopes();
    //must be the same as permision.json
    //import perm from "../../../../permissions.json"
    //return perm[0].delegated;
    //Try to avoid admin consent for authorization
    return [
        "Contacts.Read",
        "Presence.Read.All",
        "User.ReadBasic.All",
        "People.Read"
    ]
}

/*
const credential = new TeamsUserCredential(); 
const user = await credential.getUserInfo();
{"displayName":"Admin Aptero","objectId":"abb64646-d3e2-473d-8d13-0b6faff23605","preferredUserName":"admin@aptero.co"}
*/

export function lisenToAuthTokenrequest() {
    meetVsTeamsBridge.eventListener.on(REQUEST_USER_INFO, async data => {
        const credential = new TeamsUserCredential();
        const user = await credential.getUserInfo();
        meetVsTeamsBridge.sendIframeMessage(USER_INFO, user);
    });
    meetVsTeamsBridge.eventListener.on(REQUEST_AUTH_TOKEN, async data => {
        try {
            let token = await passiveAuth();
            meetVsTeamsBridge.sendIframeMessage(AUTH_TOKEN, { token: await token.getToken(getScopes()) })
        } catch (e) {
            console.error(e);
        }
    });
}

export async function passiveAuth():Promise<TeamsUserCredential> {
    console.log("initAuth");
    const cred = getCredential();
    await initGraphToolkit(cred, getScopes());
    Providers.globalProvider.setState(ProviderState.SignedIn);
    return cred;
}

export async function userInteractionAuth(consentNeeded: () => void):Promise<TeamsUserCredential> {
    console.log("initAuth");
    const cred = getCredential();
    //https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-flow-tab
    microsoftTeams.authentication.authenticate();
    await initGraphToolkit(cred, getScopes());
    await checkIsConsentNeeded(consentNeeded);
    return cred;
}

export async function logout(){
    if(Providers.globalProvider.state === ProviderState.SignedIn){
        if(Providers.globalProvider && Providers.globalProvider.logout){
            await Providers.globalProvider.logout();
        }
        Providers.globalProvider.setState(ProviderState.SignedOut);
        meetVsTeamsBridge.sendIframeMessage(RELOAD);
    }
}

export function initTeamsFx(authorityHost: string) {
    loadConfiguration({
        authentication: {
            initiateLoginEndpoint: process.env.REACT_APP_START_LOGIN_PAGE_URL,
            clientId: process.env.REACT_APP_CLIENT_ID
        }
    });

    /*if (process.env.NODE_ENV === "development") {
        setLogLevel(LogLevel.Verbose);
        setLogFunction((leve: LogLevel, message: string) => { console.log(message); });
    }*/
    if (process.env.NODE_ENV === "development") {
        setLogLevel(LogLevel.Warn);
        setLogFunction((leve: LogLevel, message: string) => { console.warn(message); });
    }
    loadConfiguration({
        authentication: {
            authorityHost,
            initiateLoginEndpoint: startLoginPageUrl,
            //simpleAuthEndpoint: teamsfxEndpoint,//https://github.com/OfficeDev/TeamsFx/wiki/How-Authentication-Works-in-TeamsFx-Tab-Template
            clientId: clientId,
        },
        resources: [
            {
                type: ResourceType.API,
                name: "default",
                properties: {
                    endpoint: functionEndpoint,
                },
            },
        ],
    });
    credential = new TeamsUserCredential();
}

function initGraphToolkit(credential: TeamsUserCredential, scope: string[]) {
    const provider = new TeamsFxProvider(credential, scope)
    Providers.globalProvider = provider;
}

async function checkIsConsentNeeded(consentNeeded: () => void) {
    try {
        await credential?.getToken(getScopes());
        Providers.globalProvider.setState(ProviderState.SignedIn);
    } catch (error) {
        Providers.globalProvider.setState(ProviderState.SignedOut);
        consentNeeded();
    }
}