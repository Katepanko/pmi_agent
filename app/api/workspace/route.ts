import { authenticatedUserId, loadWorkspace, syncWorkspace, type WorkspaceSnapshot } from "../../lib/persistence";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    return Response.json(await loadWorkspace(authenticatedUserId(request)));
  } catch (error) {
    return Response.json({ error: error instanceof Error ? error.message : "Unable to load the workspace." }, { status: 503 });
  }
}

export async function PUT(request: Request) {
  try {
    const snapshot = (await request.json()) as WorkspaceSnapshot;
    if (!Array.isArray(snapshot.projects) || !Array.isArray(snapshot.chats)) return Response.json({ error: "Invalid workspace snapshot." }, { status: 400 });
    await syncWorkspace(authenticatedUserId(request), snapshot);
    return Response.json({ saved: true });
  } catch (error) {
    return Response.json({ error: error instanceof Error ? error.message : "Unable to save the workspace." }, { status: 503 });
  }
}
