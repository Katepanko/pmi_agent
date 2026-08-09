import { buildGroundedPrompt, type SourceManifestItem } from "../../lib/pmi-prompt";
import { getProvider } from "../../lib/llm";
import { requireModel } from "../../lib/models";
import {
  buildPresentationPlanningPrompt,
  isPresentationRequest,
  parsePresentationModel,
  POWERPOINT_MIME,
  presentationFileName,
  renderPresentation,
  type PresentationModel,
} from "../../lib/presentation";
import { authenticatedUserId, loadLatestPresentation, savePresentationArtifact } from "../../lib/persistence";

export const dynamic = "force-dynamic";

type ChatBody = {
  modelKey?: string;
  message?: string;
  history?: Array<{ role: "user" | "assistant"; content: string }>;
  projectContext?: string;
  audience?: string;
  sources?: SourceManifestItem[];
  sourceRules?: string[];
  currentDraft?: string;
  chatId?: string;
  assistantMessageId?: string;
  projectId?: string | null;
  chatTitle?: string;
};

type RuntimeEnv = { FILES?: R2Bucket };

export async function POST(request: Request) {
  try {
    const body = (await request.json()) as ChatBody;
    if (!body.message?.trim()) return Response.json({ error: "A message is required." }, { status: 400 });

    const model = requireModel(body.modelKey ?? "openai-gpt56");
    const provider = getProvider(model.provider);
    const previous = body.chatId
      ? await loadLatestPresentation(authenticatedUserId(request), body.chatId).catch(() => null)
      : null;

    if (isPresentationRequest(body.message, Boolean(previous))) {
      if (!body.chatId || !body.assistantMessageId) {
        return Response.json({ error: "Chat and message IDs are required for presentation generation." }, { status: 400 });
      }
      const planningPrompt = buildPresentationPlanningPrompt({
        request: body.message,
        audience: body.audience ?? "Infer from request",
        projectContext: body.projectContext,
        sources: body.sources ?? [],
        history: body.history ?? [],
        currentPresentation: previous?.presentation as PresentationModel | null | undefined,
      });
      const planned = await provider.generate({
        model,
        system: planningPrompt,
        messages: [{ role: "user", content: body.message }],
        signal: request.signal,
      });
      const presentation = parsePresentationModel(planned, body.audience ?? "Management");
      const bytes = await renderPresentation(presentation);
      const userId = authenticatedUserId(request);
      const artifactId = crypto.randomUUID();
      const version = (previous?.version ?? 0) + 1;
      const filename = presentationFileName(presentation, version);
      const objectKey = `artifacts/${userId}/${body.chatId}/${artifactId}/${filename}`;
      const { env } = await import("cloudflare:workers");
      const bucket = (env as RuntimeEnv).FILES;
      if (!bucket) throw new Error("Artifact storage is unavailable: the FILES binding is not configured.");
      await bucket.put(objectKey, bytes, {
        httpMetadata: { contentType: POWERPOINT_MIME, contentDisposition: `attachment; filename="${filename}"` },
        customMetadata: { userId, chatId: body.chatId, messageId: body.assistantMessageId, version: String(version) },
      });
      const responseText = `${version > 1 ? "I updated" : "I created"} the ${presentation.title} PowerPoint presentation with ${presentation.slides.length} slides for ${presentation.audience}.`;
      await savePresentationArtifact({
        userId,
        chatId: body.chatId,
        messageId: body.assistantMessageId,
        projectId: body.projectId,
        chatTitle: body.chatTitle ?? presentation.title,
        audience: body.audience ?? presentation.audience,
        modelKey: body.modelKey ?? "openai-gpt56",
        message: responseText,
        artifactId,
        filename,
        mimeType: POWERPOINT_MIME,
        objectKey,
        sizeBytes: bytes.byteLength,
        version,
        slideCount: presentation.slides.length,
        presentation,
      });
      return Response.json({
        kind: "artifact",
        message: responseText,
        artifact: {
          id: artifactId,
          name: filename,
          format: "pptx",
          mimeType: POWERPOINT_MIME,
          url: `/api/artifacts/${artifactId}`,
          size: bytes.byteLength,
          slideCount: presentation.slides.length,
          version,
        },
      }, { headers: { "cache-control": "no-store" } });
    }

    const system = buildGroundedPrompt({
      projectContext: body.projectContext,
      audience: body.audience,
      sources: body.sources ?? [],
      sourceRules: body.sourceRules,
      currentDraft: body.currentDraft,
    });
    const stream = await provider.stream({
      model,
      system,
      messages: [...(body.history ?? []).slice(-20), { role: "user", content: body.message }],
      signal: request.signal,
    });

    return new Response(stream, {
      headers: {
        "content-type": "text/plain; charset=utf-8",
        "cache-control": "no-store",
        "x-accel-buffering": "no",
      },
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unable to generate the response.";
    return Response.json({ error: message }, { status: 503 });
  }
}
