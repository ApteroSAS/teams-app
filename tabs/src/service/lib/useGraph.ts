import { useData } from "./useData";
import { createMicrosoftGraphClient } from "@microsoft/teamsfx";
import { Client, GraphError } from "@microsoft/microsoft-graph-client";
import { getCredential, getScopes } from "../auth/Auth";

export function useGraph<T>(
  asyncFunc: (graph: Client) => Promise<T>,
  options?: { scope: string | string[] }
) {
  const { scope } = { scope: getScopes(), ...options };
  const initial = useData(async () => {
    try {
      const graph = createMicrosoftGraphClient(getCredential(), scope);
      return await asyncFunc(graph);
    } catch (err: unknown) {
      if (err instanceof GraphError && err.code?.includes("UiRequiredError")) {
        // Silently fail for user didn't login error
      } else {
        throw err;
      }
    }
  });

  const { data, error, loading, reload } = useData(
    async () => {
      await getCredential().login(scope);
      const graph = createMicrosoftGraphClient(getCredential(), scope);
      return await asyncFunc(graph);
    },
    { auto: false }
  );

  return data || error || loading
    ? { data, error, loading, reload }
    : {
        data: initial.data,
        error: initial.error,
        loading: initial.loading,
        reload,
      };
}
