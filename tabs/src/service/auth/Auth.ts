import { Providers, ProviderState } from '@microsoft/mgt-react';
import { TeamsFxProvider } from "@microsoft/mgt-teamsfx-provider";
import * as microsoftTeams from "@microsoft/teams-js";
import { loadConfiguration, TeamsUserCredential, ResourceType, LogLevel, setLogLevel, setLogFunction } from "@microsoft/teamsfx";
import { AUTH_TOKEN, meetVsTeamsBridge, RELOAD, REQUEST_AUTH_TOKEN, REQUEST_USER_INFO, USER_INFO } from './MeetVsTeamsBridge';
import EventEmitter from "eventemitter3";

//https://github.com/OfficeDev/TeamsFx/wiki/How-Authentication-Works-in-TeamsFx-Tab-Template

let credential: TeamsUserCredential | null; //Due to the limitation of msal, multiple instances of TeamsUserCredential is not recommended.
let startLoginPageUrl = process.env.REACT_APP_START_LOGIN_PAGE_URL;
let functionEndpoint = process.env.REACT_APP_FUNC_ENDPOINT;
let clientId = process.env.REACT_APP_CLIENT_ID;

export const LOGGED_BASIC_INFO = "ms-auth-basic-info";
export const TOKEN_RECEIVED = "ms-auth-basic-info";
let events = new EventEmitter();
let passiveAuthInProgress = false;

export function getAuthEvents(){
    return events;
}

export function getCredential(): TeamsUserCredential {
    if (!credential) {
        initTeamsFx("common");
    }
    if (!credential) {
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
        // Teams app permission
        "Contacts.Read", //use for visit card
        "Presence.Read.All", //use for visit card
        "User.ReadBasic.All", // Use for basic email, call and chat call to action, Avatar image and name - Mandatory
        "People.Read", //use for visit card
        //Permission for the web part
        "profile",//Avatar image and name - Mandatory
        "openid",//Avatar image and name - Mandatory
        "Files.Read",//Use to add content in the 3d space (pdf, img, mp4 etc...)
        "Files.Read.All"//Use to add content in the 3d space (pdf, img, mp4 etc...)
    ]
}

/*
const credential = new TeamsUserCredential(); 
const user = await credential.getUserInfo();
{"displayName":"Admin Aptero","objectId":"abb64646-d3e2-473d-8d13-0b6faff23605","preferredUserName":"admin@aptero.co"}
*/

export function lisenToAuthTokenrequest() {
    meetVsTeamsBridge.eventListener.on(REQUEST_USER_INFO, async data => {
        try {
            const user = await getCredential().getUserInfo();
            meetVsTeamsBridge.sendIframeMessage(USER_INFO, user);        
        } catch (e) {
            console.error(e);
        }
    });
    meetVsTeamsBridge.eventListener.on(REQUEST_AUTH_TOKEN, async data => {
        try {
            await passiveAuth();
            const token = getCredential();
            meetVsTeamsBridge.sendIframeMessage(AUTH_TOKEN, { token: await token.getToken(getScopes()) })
        } catch (e) {
            console.error(e);
        }
    });
}

export async function passiveAuth(): Promise<boolean> {
    if(!passiveAuthInProgress){
        passiveAuthInProgress = true;
        try{
            console.log("passiveAuth");
            const cred = await initBasicInfo();
            await initGraphToolkit(cred, getScopes());
            Providers.globalProvider.setState(ProviderState.SignedIn);
            passiveAuthInProgress = false;
            return passiveAuthInProgress;
        }catch (e) {
            passiveAuthInProgress = false;
            throw e;
        }
    }
    return passiveAuthInProgress;
}

export async function userInteractionAuth(consentNeeded: () => void): Promise<TeamsUserCredential> {
    //To test and revok application auth => https://myapplications.microsoft.com/
    console.log("userInteractionAuth");
    const cred = await initBasicInfo();
    //await teamsAuth();
    await initGraphToolkit(cred, getScopes());
    await checkIsConsentNeeded(consentNeeded);
    return cred;
}

export async function logout() {
    if (Providers.globalProvider.state === ProviderState.SignedIn) {
        if (Providers.globalProvider && Providers.globalProvider.logout) {
            await Providers.globalProvider.logout();
        }
        Providers.globalProvider.setState(ProviderState.SignedOut);
        meetVsTeamsBridge.sendIframeMessage(RELOAD);
    }
}

async function initBasicInfo():Promise<TeamsUserCredential> {
    const cred = getCredential();
    await cred.getUserInfo();
    events.emit(LOGGED_BASIC_INFO)
    return cred;
}

async function teamsAuth(){
    return new Promise((resolve, reject) =>{
        try {
            
            const loginUrl = new URL(startLoginPageUrl || "ERROR");
            //loginUrl.searchParams.set("clientId", clientId);
            //loginUrl.searchParams.set("scope", getScopes());
            //loginUrl.searchParams.set("loginHint", clientId);

            //https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-tab-aad
            microsoftTeams.authentication.authenticate({
                url: loginUrl.toString(),
                width: 600,
                height: 535,
                successCallback: function (result) {
                    console.log(result);
                    resolve(result);
                },
                failureCallback: function (reason) {
                    console.error(reason);
                    reject(reason);
                }
            });
        } catch (e) {
            console.error(e);
        }
    })
}

function initTeamsFx(authorityHost: string) {
    /*loadConfiguration({
        authentication: {
            initiateLoginEndpoint: process.env.REACT_APP_START_LOGIN_PAGE_URL,
            clientId: process.env.REACT_APP_CLIENT_ID
        }
    });*/
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
        events.emit(AUTH_TOKEN);
    } catch (error) {
        Providers.globalProvider.setState(ProviderState.SignedOut);
        consentNeeded();
    }
}