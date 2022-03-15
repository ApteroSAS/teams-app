import React from "react";
import { Button, Flex } from "@fluentui/react-northstar";
import { useGraph } from "../../service/lib/useGraph";
import { getScopes, logout, passiveAuth, userInteractionAuth } from "../../service/auth/AuthScope";
import { Login } from "@microsoft/mgt-react";
import { LightProfileCard } from "./LightProfileCard";



export function AccountAuth(props: {
  loginInitiated?: (e: Event) => void;
  loginCompleted?: (e: Event) => void;
  loginFailed?: (e: Event) => void;
  logoutInitiated?: (e: Event) => void;
  logoutCompleted?: (e: Event) => void;
}
) {

  const { loading, error, data, reload } = useGraph(
    async (graph) => {
      await passiveAuth();
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

  const requestAuth = async ()=>{
    await userInteractionAuth(()=>{
      console.error("consent is needed");
    })
    reload();
  }

  console.log("user id :" + data?.profile.id);
  //https://docs.microsoft.com/en-us/graph/toolkit/components/person-card
  return (
    <>
      {!data &&
        <div style={{ margin: "8px" }} >
          {!loading && error && <p>Failed to authorize. Please try again. <br /> Details: {error.toString()}</p>}
          {!loading && !error && <p>Authorize this app to acces advanced features.</p>}
          <Flex gap="gap.small">
            <LightProfileCard />
            <Button primary content="Authorize" disabled={loading} onClick={requestAuth} style={{ height: "auto" }} />
          </Flex>
        </div>
      }
      {!loading && data && <Login {...props} 
      loginInitiated={()=>{
        requestAuth();
      }} 
      logoutInitiated={()=>{
        logout();
      }}/>}
    </>
  );
}
