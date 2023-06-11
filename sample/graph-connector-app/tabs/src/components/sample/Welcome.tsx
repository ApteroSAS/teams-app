import { useContext } from "react";
import { Button, Image, Spinner } from "@fluentui/react-components";
import "./Welcome.css";
import { useGraphWithCredential } from "@microsoft/teamsfx-react";
import { Scopes } from "./constants";
import { TeamsFxContext } from "../Context";
import { MicrosoftGraphService } from "./lib/MicrosoftGraphService";
import { TeamsLoginService } from "./lib/TeamsLoginService";
import { IdentityType, TeamsFx, TeamsUserCredential } from "@microsoft/teamsfx";
import config from "./config";

export function Welcome() {
  /*const { teamsUserCredential } = useContext(TeamsFxContext);
  const { loading, error, data, reload } = useGraphWithCredential(
    async (graph, teamsUserCredential, scope) => {
      // Call graph api directly to get user profile information
      const profile = await graph.api("/me").get();
      return { profile };
    },
    { scope: Scopes, credential: teamsUserCredential }
  );*/
  const loading = false;
  const data = true;
  const error = false;
  const reload = ()=>{}

  return (
    <>
      {loading && (
        <div>
          <Spinner style={{ margin: 100 }} />
        </div>
      )}
      {!loading && data && (
        <div className="welcome page">
          <div className="narrow page-padding">
            <Image src="hello.png" />
            {/*<h1 className="center">Hello, {data.profile.displayName}!</h1>*/}
            <p className="center">
              Let's build your first custom Microsoft Graph connector.<br/>
              <Button onClick={
                async ()=>{
                  console.log("Graph API call");
                  const microsoftGraphService = new MicrosoftGraphService();
                  const tls = new TeamsLoginService(microsoftGraphService);
                  tls.initGraphTeamsUserCredential(Scopes);
                  await microsoftGraphService.awaitReady()
                  console.log(await microsoftGraphService.getUserDetails());
                }}>Graph Call (see console)</Button><br/>
                <Button onClick={
                  async ()=>{
                    console.log("TeamsFx Login");
                    const teamsfx = new TeamsFx(IdentityType.User,{
                      initiateLoginEndpoint: config.initiateLoginEndpoint!,
                      clientId: config.clientId!
                    });
                    await teamsfx.login(Scopes);
                    console.log("TeamsFx Login done");
                  }}>TeamsFx Login</Button><br/>
                  <Button onClick={
                    async ()=>{
                      console.log("teamsUserCredential Login");
                      const teamsUserCredential = new TeamsUserCredential({
                        initiateLoginEndpoint: config.initiateLoginEndpoint!,
                        clientId: config.clientId!
                      });
                      await teamsUserCredential.login(Scopes);
                      console.log("teamsUserCredential Login done");
                    }}>TeamsUserCredential Login</Button>
            </p>
          </div>
        </div>
      )}
      {!loading && !data && (
        <div className="auth">
          <h2>Welcome to Graph Connector App!</h2>
          <Button appearance="primary" disabled={loading} onClick={reload}>
            Start
          </Button>
        </div>
      )}
      {!loading && error && (
        <div className="error">
          Failed to get your profile. Please try again later. <br /> Details:{" "}
          {/*error.toString()*/}
        </div>
      )}
    </>
  );
}
