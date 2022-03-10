import React from "react";
import { Button } from "@fluentui/react-northstar";
import { useGraph } from "./lib/useGraph";
import { ProfileCard } from "./ProfileCard";
import { getScopes } from "../../service/auth/AuthScope";



export function AccountAuthV2() {
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
    { scope: getScopes() }
  );
  console.log("user id :"+data?.profile.id);
  //https://docs.microsoft.com/en-us/graph/toolkit/components/person-card
  return (
    <div>
      {!data && 
      <div>
        <p>Click below to authorize this app.</p>
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
