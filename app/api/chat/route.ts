import { buildGroundedPrompt, type SourceManifestItem } from "../../lib/pmi-prompt";
import { getProvider } from "../../lib/llm";
import { requireModel } from "../../lib/models";
import {
  artifactTitle,
  detectArtifactRequest,
  parseArtifactModel,
  planArtifact,
  renderArtifact,
  validateArtifact,
  type ArtifactContentModel,
  type ArtifactFormat,
} from "../../lib/artifact";
import { authenticatedUserId, loadLatestArtifact, loadLatestArtifactModel, saveArtifact } from "../../lib/persistence";
import { getRuntimeBindings } from "../../lib/runtime-bindings";

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

export async function POST(request: Request) {
  try {
    const body = (await request.json()) as ChatBody;
    if (!body.message?.trim()) return Response.json({ error: "A message is required." }, { status: 400 });

    const model = requireModel(body.modelKey ?? "openai-gpt56");
    const provider = getProvider(model.provider);
    const userId = authenticatedUserId(request);
    const latest = body.chatId
      ? await loadLatestArtifact(userId, body.chatId).catch(() => null)
      : null;
    const requestedFormat = detectArtifactRequest(body.message, latest?.format);

    if (requestedFormat) {
      if (!body.chatId || !body.assistantMessageId) {
        return Response.json({ error: "Chat and message IDs are required for artifact generation." }, { status: 400 });
      }
      const previous = await loadLatestArtifactModel(userId, body.chatId, requestedFormat).catch(() => null);
      const planningPrompt = planArtifact({
        format: requestedFormat,
        request: body.message,
        audience: body.audience ?? "Infer from request",
        projectContext: body.projectContext,
        sources: body.sources ?? [],
        history: body.history ?? [],
        currentModel: previous?.model as ArtifactContentModel | null | undefined,
      });
      const planned = await provider.generate({
        model,
        system: planningPrompt,
        messages: [{ role: "user", content: body.message }],
        signal: request.signal,
      });
      const artifactModel = parseArtifactModel(requestedFormat, planned, body.audience ?? "Management");
      const artifactId = crypto.randomUUID();
      const version = (previous?.version ?? 0) + 1;
      const rendered = await renderArtifact({ format: requestedFormat, model: artifactModel, version, sources: body.sources ?? [] });
      await validateArtifact(rendered);
      const objectKey = `artifacts/${userId}/${body.chatId}/${artifactId}/${rendered.filename}`;
      const bucket = getRuntimeBindings().FILES;
      if (!bucket) throw new Error("Artifact storage is unavailable: the FILES binding is not configured.");
      await bucket.put(objectKey, rendered.bytes, {
        httpMetadata: { contentType: rendered.mimeType, contentDisposition: `attachment; filename="${rendered.filename}"` },
        customMetadata: { userId, chatId: body.chatId, messageId: body.assistantMessageId, version: String(version), format: rendered.format },
      });
      const formatName: Record<ArtifactFormat, string> = {
        pptx: "PowerPoint presentation", xlsx: "Excel workbook", docx: "Word document", pdf: "PDF report", html: "HTML dashboard",
      };
      const audience = "audience" in artifactModel ? artifactModel.audience : body.audience ?? "Management";
      const responseText = `${version > 1 ? "I updated" : "I created"} the ${artifactTitle(artifactModel)} ${formatName[requestedFormat]} with ${rendered.unitCount} ${rendered.unitLabel} for ${audience}.`;
      await saveArtifact({
        userId,
        chatId: body.chatId,
        messageId: body.assistantMessageId,
        projectId: body.projectId,
        chatTitle: body.chatTitle ?? artifactTitle(artifactModel),
        audience: body.audience ?? audience,
        modelKey: body.modelKey ?? "openai-gpt56",
        message: responseText,
        artifactId,
        filename: rendered.filename,
        mimeType: rendered.mimeType,
        objectKey,
        sizeBytes: rendered.bytes.byteLength,
        format: rendered.format,
        version,
        unitCount: rendered.unitCount,
        unitLabel: rendered.unitLabel,
        model: artifactModel,
        parentArtifactId: previous?.artifactId,
      });
      return Response.json({
        kind: "artifact",
        message: responseText,
        artifact: {
          id: artifactId,
          name: rendered.filename,
          format: rendered.format,
          mimeType: rendered.mimeType,
          url: `/api/artifacts/${artifactId}`,
          size: rendered.bytes.byteLength,
          unitCount: rendered.unitCount,
          unitLabel: rendered.unitLabel,
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
