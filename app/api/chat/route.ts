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
import { reconcileEvidence } from "../../lib/evidence";
import { enforceConflictVisibility } from "../../lib/conflict-guard";
import { authenticatedUserId, loadLatestArtifact, loadLatestArtifactModel, saveArtifact } from "../../lib/persistence";
import { getRuntimeBindings } from "../../lib/runtime-bindings";
import { generateStructuredModel } from "../../lib/structured-generation";
import { artifactStructuredOutput } from "../../lib/artifact-schema";
import { hasUnresolvedTemplateDirective, resolveTemplateReference, templateOutputFormat } from "../../lib/template";
import { analyzePresentationTemplate } from "../../lib/presentation-template";
import { analyzeReportTemplate } from "../../lib/report-template";
import {
  applyBlockEdits,
  assertLockedBlocksUnchanged,
  assertRenderedTextIntegrity,
  blockEditStructuredOutput,
  buildExistingContentDesignPrompt,
  buildScopedEditPrompt,
  existingContentDesignStructuredOutput,
  parseBlockEditResponse,
  parseExistingContentDesignPlan,
  resolveExistingContentRequest,
} from "../../lib/existing-content";
import { renderExistingContent } from "../../lib/renderers/existing-content";

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

    const sources = body.sources ?? [];
    const selectedTemplate = resolveTemplateReference(body.message, sources);
    if (!selectedTemplate && hasUnresolvedTemplateDirective(body.message)) {
      throw new Error("No uploaded file matches the @template reference. Attach the template, then use its @ button or type its exact filename.");
    }
    const evidenceSources = selectedTemplate ? sources.filter((source) => source.id !== selectedTemplate.sourceId) : sources;
    const userStatements = [...(body.history ?? []).filter((entry) => entry.role === "user").map((entry) => entry.content), body.message];
    const reconciliation = reconcileEvidence(evidenceSources, { userStatements, authorityRules: body.sourceRules });
    const userId = authenticatedUserId(request);
    if (selectedTemplate) {
      const templateObject = await getRuntimeBindings().FILES?.get(`staged/${userId}/${selectedTemplate.sourceId}`);
      if (!templateObject) throw new Error(`The selected template ${selectedTemplate.fileName} is no longer available. Re-attach it and retry.`);
      selectedTemplate.bytes = new Uint8Array(await templateObject.arrayBuffer());
      if (selectedTemplate.fileType === "pptx") {
        selectedTemplate.layoutModel = analyzePresentationTemplate(selectedTemplate.bytes) as unknown as Record<string, unknown>;
      } else {
        selectedTemplate.layoutModel = await analyzeReportTemplate(selectedTemplate.bytes, selectedTemplate.fileType);
      }
    }
    const latest = body.chatId
      ? await loadLatestArtifact(userId, body.chatId).catch(() => null)
      : null;
    const requestedFormat = detectArtifactRequest(body.message, latest?.format) ?? (selectedTemplate ? templateOutputFormat(selectedTemplate) : null);

    if (requestedFormat) {
      if (!body.chatId || !body.assistantMessageId) {
        return Response.json({ error: "Chat and message IDs are required for artifact generation." }, { status: 400 });
      }
      const previous = await loadLatestArtifactModel(userId, body.chatId, requestedFormat).catch(() => null);
      const existingContent = resolveExistingContentRequest({
        message: body.message,
        format: requestedFormat,
        history: body.history ?? [],
      });
      if (existingContent) {
        const model = requireModel(body.modelKey ?? "openai-gpt56");
        const provider = getProvider(model.provider);
        let finalBlocks = existingContent.blocks;
        if (existingContent.editInstruction && existingContent.editableBlockIds.length) {
          const editPrompt = buildScopedEditPrompt(existingContent.blocks, existingContent.editInstruction);
          const edits = await generateStructuredModel({
            provider,
            model,
            system: editPrompt,
            userMessage: existingContent.editInstruction,
            structuredOutput: blockEditStructuredOutput,
            parse: parseBlockEditResponse,
            outputLabel: "scoped locked-text edit",
            signal: request.signal,
          });
          finalBlocks = applyBlockEdits(existingContent.blocks, edits);
        }
        assertLockedBlocksUnchanged(existingContent.blocks, finalBlocks);
        const designPrompt = buildExistingContentDesignPrompt({
          blocks: finalBlocks,
          format: existingContent.format,
          request: body.message,
          templateDescription: selectedTemplate ? { fileName: selectedTemplate.fileName, fileType: selectedTemplate.fileType, layoutModel: selectedTemplate.layoutModel } : null,
        });
        const designPlan = await generateStructuredModel({
          provider,
          model,
          system: designPrompt,
          userMessage: body.message,
          structuredOutput: existingContentDesignStructuredOutput,
          parse: (raw) => parseExistingContentDesignPlan(raw, finalBlocks),
          outputLabel: "locked-content design plan",
          signal: request.signal,
        });
        const artifactId = crypto.randomUUID();
        const version = (previous?.version ?? 0) + 1;
        const rendered = await renderExistingContent({ format: existingContent.format, blocks: finalBlocks, version, plan: designPlan });
        assertRenderedTextIntegrity(finalBlocks, rendered.renderedTextBlocks);
        await validateArtifact(rendered, null);
        const objectKey = `artifacts/${userId}/${body.chatId}/${artifactId}/${rendered.filename}`;
        const bucket = getRuntimeBindings().FILES;
        if (!bucket) throw new Error("Artifact storage is unavailable: the FILES binding is not configured.");
        await bucket.put(objectKey, rendered.bytes, {
          httpMetadata: { contentType: rendered.mimeType, contentDisposition: `attachment; filename="${rendered.filename}"` },
          customMetadata: { userId, chatId: body.chatId, messageId: body.assistantMessageId, version: String(version), format: rendered.format, generationMode: existingContent.generationMode },
        });
        const title = finalBlocks.find((block) => block.kind === "heading")?.text ?? finalBlocks[0].text.slice(0, 72);
        const formatName: Record<string, string> = { pptx: "PowerPoint presentation", docx: "Word document", pdf: "PDF", html: "HTML file" };
        const responseText = `I ${version > 1 ? "updated" : "created"} the ${formatName[rendered.format]} from the referenced response${existingContent.editableBlockIds.length ? ", changing only the explicitly requested text" : " with its wording preserved"}.`;
        await saveArtifact({
          userId,
          chatId: body.chatId,
          messageId: body.assistantMessageId,
          projectId: body.projectId,
          chatTitle: body.chatTitle ?? title,
          audience: body.audience ?? "Management",
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
          model: { generationMode: existingContent.generationMode, blocks: finalBlocks, designPlan },
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
      const model = requireModel(body.modelKey ?? "openai-gpt56");
      const provider = getProvider(model.provider);
      const planningPrompt = planArtifact({
        format: requestedFormat,
        request: body.message,
        audience: body.audience ?? "Infer from request",
        projectContext: body.projectContext,
        sources: evidenceSources,
        history: body.history ?? [],
        currentModel: previous?.model as ArtifactContentModel | null | undefined,
        reconciliation,
        template: selectedTemplate,
      });
      const artifactModel = enforceConflictVisibility(await generateStructuredModel({
        provider,
        model,
        system: planningPrompt,
        userMessage: body.message,
        structuredOutput: artifactStructuredOutput(requestedFormat),
        parse: (raw) => parseArtifactModel(requestedFormat, raw, body.audience ?? "Management"),
        outputLabel: `${requestedFormat.toUpperCase()} content model`,
        signal: request.signal,
      }), reconciliation);
      const artifactId = crypto.randomUUID();
      const version = (previous?.version ?? 0) + 1;
      const rendered = await renderArtifact({ format: requestedFormat, model: artifactModel, version, sources: evidenceSources, template: selectedTemplate });
      await validateArtifact(rendered, selectedTemplate);
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
      const responseText = `${version > 1 ? "I updated" : "I created"} the ${artifactTitle(artifactModel)} ${formatName[requestedFormat]} with ${rendered.unitCount} ${rendered.unitLabel} for ${audience}${selectedTemplate ? `, using ${selectedTemplate.fileName} as the template` : ""}.`;
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

    const model = requireModel(body.modelKey ?? "openai-gpt56");
    const provider = getProvider(model.provider);
    const system = buildGroundedPrompt({
      projectContext: body.projectContext,
      audience: body.audience,
      sources: evidenceSources,
      sourceRules: body.sourceRules,
      currentDraft: body.currentDraft,
      reconciliation,
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
