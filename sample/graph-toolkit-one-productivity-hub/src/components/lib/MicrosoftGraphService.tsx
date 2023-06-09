import { Buffer } from "buffer";
import { Client } from "@microsoft/microsoft-graph-client";

Buffer.from("anything", "base64");

export class MicrosoftGraphService {
  client: Client;
  setReady = ()=>{};
  ready = new Promise((resolve) => {this.setReady = resolve as any});

  async init(client:Client) {
    if(!client){
      throw new Error("MicrosoftGraphService : init : client is null");
    }
    console.log("MicrosoftGraphService : init");
    this.client = client;
    await this.getUserDetails();//this is used to trigger token access verification will throw error if not authorized
    this.setReady();
  }

  public async awaitReady(){
    return await this.ready;
  }

  getClient() {
    return this.client;
  }

  isInitialized() {
    return this.client;
  }

  async getUserDetails() {
    return await this.client.api("/me").get();
  }

  async getUserDetailsBetaV() {
    return await this.client.api("/me/profile").version("beta").get();
  }

  async getUserOrganization() {
    //https://graph.microsoft.com/v1.0/organization
    //check if the  user account is a work account (MSA For personal account and AAD for work account)
    try {
      const organization = await this.client.api("/organization").get();
      return organization.value[0].displayName;
    } catch (e) {
      return null;
    }
  }

  blobToBase64(blob): Promise<string> {
    return new Promise((resolve, _) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result as string);
      reader.readAsDataURL(blob);
    });
  }

  async getUserPhotoAsBase64(): Promise<string> {
    try {
      let profile = await this.client.api("https://graph.microsoft.com/v1.0/me/photo/$value").get();
      let b64 = await this.blobToBase64(profile);
      return b64;
    } catch (e) {
      console.log("profile photo not found");
    }
    return "";
  }
}
