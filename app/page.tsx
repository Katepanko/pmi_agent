import { getModelRegistry } from "./lib/models";
import { PMIWorkspace } from "./pmi-workspace";

export const dynamic = "force-dynamic";

export default function Home() {
  const models = getModelRegistry().map(({ key, displayName, provider, available, unavailableReason }) => ({
    key,
    displayName,
    provider,
    available,
    unavailableReason,
  }));

  return <PMIWorkspace initialModels={models} />;
}
