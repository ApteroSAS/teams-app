import { Image } from "@fluentui/react-northstar";
import "./Welcome.css";
import { Graph } from "../../components/profile/Graph";
import { useTeamsFx } from "../../components/profile/lib/useTeamsFx";
import { TeamsUserCredential } from "@microsoft/teamsfx";
import { useData } from "../../components/profile/lib/useData";

export function Welcome(props: { showFunction?: boolean; environment?: string }) {
  const { environment } = {
    environment: window.location.hostname === "localhost" ? "local" : "azure",
    ...props,
  };
  const friendlyEnvironmentName =
    {
      local: "local environment",
      azure: "Azure environment",
    }[environment] || "local environment";

  const { isInTeams } = useTeamsFx();
  const userProfile = useData(async () => {
    const credential = new TeamsUserCredential();
    return isInTeams ? await credential.getUserInfo() : undefined;
  })?.data;
  const userName = userProfile ? userProfile.displayName : "";
  return (
    <div className="welcome page">
      <div className="narrow page-padding">
        <Image src="hello.png" />
        <h1 className="center">Congratulations{userName ? ", " + userName : ""}!</h1>
        <p className="center">Your app is running in your {friendlyEnvironmentName}</p>
        <div className="sections">
          <div>
            <h2>
              Welcome to the <a href="https://www.aptero.co/">Aptero</a> teams application!
            </h2>
            <Graph />
          </div>
        </div>
      </div>
    </div>
  );
}
