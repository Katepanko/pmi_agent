import { authenticatedUserId, findArtifact } from "../../../lib/persistence";

export const dynamic = "force-dynamic";

type RuntimeEnv = { FILES?: R2Bucket };

export async function GET(request: Request, context: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await context.params;
    const artifact = await findArtifact(authenticatedUserId(request), id);
    if (!artifact) return Response.json({ error: "Artifact not found." }, { status: 404 });

    const { env } = await import("cloudflare:workers");
    const object = await (env as RuntimeEnv).FILES?.get(artifact.object_key);
    if (!object) return Response.json({ error: "The stored artifact file is unavailable." }, { status: 404 });

    return new Response(object.body, {
      headers: {
        "content-type": artifact.mime_type,
        "content-length": String(artifact.size_bytes),
        "content-disposition": `attachment; filename="${artifact.filename.replace(/["\r\n]/g, "_")}"`,
        "cache-control": "private, max-age=3600",
        "x-content-type-options": "nosniff",
      },
    });
  } catch (error) {
    return Response.json({ error: error instanceof Error ? error.message : "Unable to download the artifact." }, { status: 503 });
  }
}
