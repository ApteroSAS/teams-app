import React from "react";
import { Button } from "@fluentui/react-northstar";
import { useGraph } from "./lib/useGraph";
import { ProfileCard } from "./ProfileCard";
import { Providers } from '@microsoft/mgt-element';
import { TeamsProvider } from '@microsoft/mgt-teams-provider';



export function AccountAuth() {
  const { loading, error, data, reload } = useGraph(
    async (graph) => {
      const profile = await graph.api("/me").get();
      let photoUrl = "";
      try {
        const photo = await graph.api("/me/photo/$value").get();
        photoUrl = URL.createObjectURL(photo);
      } catch {
        // Could not fetch photo from user's profile, return empty string as placeholder.
      }
      return { profile, photoUrl };
    },
    { scope: ["User.Read"] }
  );

        //https://docs.microsoft.com/en-us/graph/toolkit/components/person-card
  return (
    <div>
      {!data && 
      <div>
        <p>Click below to authorize this app to read your profile photo using Microsoft Graph.</p>
        <Button primary content="Authorize" disabled={loading} onClick={reload} />
      </div>
      }
      {loading && ProfileCard(true)}
      {!loading && error && (
        <div className="error">
          Failed to read your profile. Please try again later. <br /> Details: {error.toString()}
        </div>
      )}
      {!loading && data && ProfileCard(false, data)}
    </div>
  );
}
