import EventEmitter from "eventemitter3";


export class MeetVsTeamsBridge{

    eventListener = new EventEmitter();
    
    start(){
        window.addEventListener("message", event => {
            if(event.data) {
                var data = JSON.parse(event.data);
                if (data.type) {
                this.eventListener.emit(data.type, data)
                }
            }
        });
    }

    sendIframeMessage(type:string, data:any) {
        window.postMessage(JSON.stringify({ type, data }), "*");
    }

}

export const meetVsTeamsBridge = new MeetVsTeamsBridge();



//on auth token
//on auth token request