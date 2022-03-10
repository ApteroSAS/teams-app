import EventEmitter from "eventemitter3";

export const INIT_OK = "ms-teams-init-ok";
export const REQUEST_AUTH_TOKEN = "ms-teams-request-auth-token";
export const AUTH_TOKEN = "ms-teams-auth-token";
export const CHANGE_USER = "ms-teams-change-selected-user";
export const CHANGE_SETTINGS_URL = "ms-teams-change-settings-url";
export const DECLARE_FRAME_SIZE = "ms-teams-declare-frame-size";

export class MeetVsTeamsBridge {

    eventListener = new EventEmitter();
    iframe: HTMLIFrameElement | null = null;

    start() {
        window.addEventListener("message", event => {
            try {
                if (event.data) {
                    var data = JSON.parse(event.data);
                    if (data.type) {
                        this.eventListener.emit(data.type, data)
                    }
                }
            } catch (err) {
                // ignore error
            }
        });
    }

    setIframe(iframe: HTMLIFrameElement | null) {
        //this can comes from react and thus beware of spam
        if(this.iframe != iframe){
            this.iframe = iframe;
            if (this.iframe) {
                this.sendIframeMessage(INIT_OK, {});
            }
        }
    }

    sendIframeMessage(type: string, data: any) {
        if (this.iframe && this.iframe.contentWindow) {
            this.iframe.contentWindow.postMessage(JSON.stringify({ type, data }), "*");
        }
    }

}

export const meetVsTeamsBridge = new MeetVsTeamsBridge();



//on auth token
//on auth token request