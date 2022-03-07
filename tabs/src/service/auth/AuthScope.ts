import { Providers, ProviderState } from '@microsoft/mgt-react';
import { TeamsFxProvider } from "@microsoft/mgt-teamsfx-provider";
import { loadConfiguration, TeamsUserCredential } from "@microsoft/teamsfx";
import { AUTH_TOKEN, meetVsTeamsBridge, REQUEST_AUTH_TOKEN } from './MeetVsTeamsBridge';

let credential: TeamsUserCredential | null = null;

export function lisenToAuthTokenrequest(){
    meetVsTeamsBridge.eventListener.on(REQUEST_AUTH_TOKEN,async data => {
        let token = await initAuth(()=>{});
        meetVsTeamsBridge.sendIframeMessage(AUTH_TOKEN,{token : await token.getToken(getScopes())})
    });
}

export async function initAuth(consentNeeded: () => void) {
    await initTeamsFx();
    if (!credential) {
        throw new Error();
    }
    await initGraphToolkit(credential, getScopes());
    await checkIsConsentNeeded(consentNeeded);
    return credential;
}

export function initGraphToolkit(credential: TeamsUserCredential, scope: string[]) {
    const provider = new TeamsFxProvider(credential, scope)
    Providers.globalProvider = provider;
}

export function initTeamsFx() {
    loadConfiguration({
        authentication: {
            initiateLoginEndpoint: process.env.REACT_APP_START_LOGIN_PAGE_URL,
            clientId: process.env.REACT_APP_CLIENT_ID
        }
    });
    credential = new TeamsUserCredential();
}

export async function checkIsConsentNeeded(consentNeeded: () => void) {
    try {
        await credential?.getToken(getScopes());
        Providers.globalProvider.setState(ProviderState.SignedIn);
    } catch (error) {
        Providers.globalProvider.setState(ProviderState.SignedOut);
        consentNeeded();
    }
}

export function getScopes(): string[] {
    //const neededScopes = MgtPersonCard.getScopes();
    //must be the same as permision.json
    //import perm from "../../../../permissions.json"
    //return perm[0].delegated;
    return [
        "Contacts.Read",
        "Presence.Read.All",
        "User.Read.All",
        "People.Read.All"
    ]
}