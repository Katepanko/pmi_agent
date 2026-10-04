import { buildGroundedPrompt, type SourceManifestItem } from "../../lib/pmi-prompt";
import { getProvider } from "../../lib/llm";
import { requireModel } from "../../lib/models";
import {
  artifactTitle,
  detectArtifactRequest,
  isArtifactRevisionRequest,
  parseArtifactModel,
  planArtifact,
  renderArtifact,
  validateArtifact,
  type ArtifactContentModel,
  type ArtifactFormat,
} from "../../lib/artifact";
import { reconcileEvidence } from "../../lib/evidence";
import { enforceConflictVisibility } from "../../lib/conflict-guard";
import { authenticatedUserId, loadLatestArtifact, loadLatestArtifactContext, loadLatestArtifactModel, loadTemplateSource, saveArtifact } from "../../lib/persistence";
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
  type ExistingContentHistoryEntry,
} from "../../lib/existing-content";
import { renderExistingContent } from "../../lib/renderers/existing-content";
import {
  assertTranslatedModelIntegrity,
  buildArtifactTranslationPrompt,
  buildVisibleTextTranslationPrompt,
  detectTranslationRequest,
  extractHtmlVisibleText,
  extractOfficeVisibleText,
  parseVisibleTextTranslations,
  translatedArtifactFileName,
  translatedArtifactLineage,
  translateOfficeArtifact,
  translateHtmlArtifact,
  visibleTextTranslationOutput,
  type VisibleTextUnit,
} from "../../lib/artifact-translation";
import type { ArtifactTemplate } from "../../lib/template";
import { PDF_VISIBLE_LABELS } from "../../lib/renderers/pdf";
import {
  detectCommunicationLanguage,
  languageName,
  referencesPriorArtifact,
  resolveArtifactLanguageContext,
  type CommunicationLanguage,
} from "../../lib/language-context";

export const dynamic = "force-dynamic";

type ChatBody = {
  modelKey?: string;
  message?: string;
  history?: ExistingContentHistoryEntry[];
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

function isArtifactContentModel(value: unknown): value is ArtifactContentModel {
  return Boolean(value && typeof value === "object" && (Array.isArray((value as { slides?: unknown }).slides) || Array.isArray((value as { sections?: unknown }).sections)));
}

function translationBatches(units: VisibleTextUnit[]) {
  const batches: VisibleTextUnit[][] = [];
  let current: VisibleTextUnit[] = [];
  let characters = 0;
  for (const unit of units) {
    if (current.length && (current.length >= 80 || characters + unit.text.length > 18_000)) {
      batches.push(current);
      current = [];
      characters = 0;
    }
    current.push(unit);
    characters += unit.text.length;
  }
  if (current.length) batches.push(current);
  return batches;
}

async function hydrateStoredTemplate(userId: string, sourceId: string): Promise<ArtifactTemplate | null> {
  const stored = await loadTemplateSource(userId, sourceId);
  if (!stored) return null;
  const object = await getRuntimeBindings().FILES?.get(stored.objectKey);
  if (!object) return null;
  const bytes = new Uint8Array(await object.arrayBuffer());
  const template: ArtifactTemplate = {
    sourceId: stored.sourceId,
    fileName: stored.fileName,
    fileType: stored.fileType,
    status: stored.status,
    excerpt: stored.excerpt,
    metadata: stored.metadata,
    warnings: stored.warnings,
    bytes,
  };
  template.layoutModel = template.fileType === "pptx"
    ? analyzePresentationTemplate(bytes) as unknown as Record<string, unknown>
    : await analyzeReportTemplate(bytes, template.fileType);
  return template;
}

function translatedResponse(format: ArtifactFormat, language: ReturnType<typeof resolveArtifactLanguageContext>["outputLanguage"], communicationLanguage: CommunicationLanguage) {
  return communicationLanguage === "de"
    ? `Ich habe das vollständige ${format.toUpperCase()}-Artefakt auf ${languageName(language, "de")} übersetzt und eine neue Datei unter Beibehaltung des ursprünglichen Layouts und der Formatierung erstellt.`
    : `I translated the complete ${format.toUpperCase()} artifact into ${languageName(language, "en")} and generated a new file with the original layout and formatting preserved.`;
}

function generatedResponse(input: {
  communicationLanguage: CommunicationLanguage;
  updated: boolean;
  title: string;
  format: ArtifactFormat;
  unitCount: number;
  unitLabel: string;
  audience: string;
  templateFileName?: string;
}) {
  if (input.communicationLanguage === "de") {
    const formatPhrase = ({ pptx: "die PowerPoint-Präsentation", xlsx: "die Excel-Arbeitsmappe", docx: "das Word-Dokument", pdf: "den PDF-Bericht", html: "das HTML-Dashboard" } as const)[input.format];
    const unitLabel = ({ slides: "Folien", sheets: "Tabellenblätter", sections: "Abschnitte", pages: "Seiten" } as Record<string, string>)[input.unitLabel] ?? input.unitLabel;
    return `Ich habe ${formatPhrase} „${input.title}“ ${input.updated ? "aktualisiert" : "erstellt"}; die Datei enthält ${input.unitCount} ${unitLabel} für ${input.audience}${input.templateFileName ? ` und verwendet ${input.templateFileName} als Vorlage` : ""}.`;
  }
  const formatName = ({ pptx: "PowerPoint presentation", xlsx: "Excel workbook", docx: "Word document", pdf: "PDF report", html: "HTML dashboard" } as const)[input.format];
  return `${input.updated ? "I updated" : "I created"} the ${input.title} ${formatName} with ${input.unitCount} ${input.unitLabel} for ${input.audience}${input.templateFileName ? `, using ${input.templateFileName} as the template` : ""}.`;
}

export async function POST(request: Request) {
  let communicationLanguage: CommunicationLanguage = "en";
  try {
    const body = (await request.json()) as ChatBody;
    if (!body.message?.trim()) return Response.json({ error: "A message is required." }, { status: 400 });
    communicationLanguage = detectCommunicationLanguage(body.message);

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
    const latestContext = body.chatId
      ? await loadLatestArtifactContext(userId, body.chatId).catch(() => null)
      : null;
    const translationLanguage = detectTranslationRequest(body.message, Boolean(latestContext));
    const priorArtifactReference = Boolean(latestContext && referencesPriorArtifact(body.message));
    const modifiesExistingArtifact = Boolean(latestContext && (translationLanguage || priorArtifactReference || isArtifactRevisionRequest(body.message)));
    const languageContext = resolveArtifactLanguageContext({
      message: body.message,
      sourceArtifactLanguageCode: latestContext?.contentLanguage,
      sourceArtifactModel: latestContext?.model,
      modifiesExistingArtifact,
    });

    if (translationLanguage && latestContext) {
      if (!body.chatId || !body.assistantMessageId) {
        return Response.json({ error: "Chat and message IDs are required for artifact translation." }, { status: 400 });
      }
      const model = requireModel(body.modelKey ?? "openai-gpt56");
      const provider = getProvider(model.provider);
      const sourceModel = latestContext.model;
      const officeFormat = latestContext.format === "docx" || latestContext.format === "pptx" || latestContext.format === "xlsx"
        ? latestContext.format
        : null;
      let translatedModel: unknown = sourceModel;
      if (isArtifactContentModel(sourceModel)) {
        const translationPrompt = buildArtifactTranslationPrompt({
          format: latestContext.format,
          language: translationLanguage,
          sourceModel,
        });
        translatedModel = await generateStructuredModel({
          provider,
          model,
          system: translationPrompt,
          userMessage: body.message,
          structuredOutput: artifactStructuredOutput(latestContext.format),
          parse: (raw) => parseArtifactModel(latestContext.format, raw, "Management"),
          outputLabel: `${translationLanguage.name} artifact translation`,
          signal: request.signal,
        });
        assertTranslatedModelIntegrity(sourceModel, translatedModel as ArtifactContentModel, latestContext.format);
      } else if (!officeFormat && latestContext.format !== "html") {
        throw new Error("The latest generated artifact does not have a translatable report model. Regenerate it once, then retry the translation.");
      }

      const version = latestContext.version + 1;
      const lineage = translatedArtifactLineage(latestContext, translationLanguage.code);
      const inheritedTemplate = latestContext.templateSourceId
        ? await hydrateStoredTemplate(userId, latestContext.templateSourceId)
        : null;
      let rendered;
      if (officeFormat) {
        const sourceObject = await getRuntimeBindings().FILES?.get(latestContext.objectKey);
        if (!sourceObject) throw new Error(`The source artifact ${latestContext.filename} is no longer available.`);
        const sourceBytes = new Uint8Array(await sourceObject.arrayBuffer());
        const sourceUnits = extractOfficeVisibleText(sourceBytes, officeFormat);
        const translatedUnits: VisibleTextUnit[] = [];
        for (const batch of translationBatches(sourceUnits)) {
          const prompt = buildVisibleTextTranslationPrompt(batch, translationLanguage);
          const translatedBatch = await generateStructuredModel({
            provider,
            model,
            system: prompt,
            userMessage: `Translate all ${batch.length} visible text units into ${translationLanguage.name}.`,
            structuredOutput: visibleTextTranslationOutput,
            parse: (raw) => parseVisibleTextTranslations(raw, batch),
            outputLabel: `${translationLanguage.name} visible document text`,
            signal: request.signal,
          });
          translatedUnits.push(...translatedBatch);
        }
        const bytes = translateOfficeArtifact({
          bytes: sourceBytes,
          format: officeFormat,
          language: translationLanguage,
          sourceUnits,
          translatedUnits,
        });
        rendered = {
          format: latestContext.format,
          mimeType: latestContext.mimeType,
          filename: translatedArtifactFileName(latestContext.filename, translationLanguage.code, version),
          bytes,
          unitCount: latestContext.unitCount,
          unitLabel: latestContext.unitLabel,
        };
      } else if (latestContext.format === "html") {
        const sourceObject = await getRuntimeBindings().FILES?.get(latestContext.objectKey);
        if (!sourceObject) throw new Error(`The source artifact ${latestContext.filename} is no longer available.`);
        const sourceBytes = new Uint8Array(await sourceObject.arrayBuffer());
        const sourceUnits = extractHtmlVisibleText(sourceBytes);
        const translatedUnits: VisibleTextUnit[] = [];
        for (const batch of translationBatches(sourceUnits)) {
          const translatedBatch = await generateStructuredModel({
            provider,
            model,
            system: buildVisibleTextTranslationPrompt(batch, translationLanguage),
            userMessage: `Translate all ${batch.length} visible HTML text units into ${translationLanguage.name}.`,
            structuredOutput: visibleTextTranslationOutput,
            parse: (raw) => parseVisibleTextTranslations(raw, batch),
            outputLabel: `${translationLanguage.name} visible HTML text`,
            signal: request.signal,
          });
          translatedUnits.push(...translatedBatch);
        }
        rendered = {
          format: latestContext.format,
          mimeType: latestContext.mimeType,
          filename: translatedArtifactFileName(latestContext.filename, translationLanguage.code, version),
          bytes: translateHtmlArtifact(sourceBytes, sourceUnits, translatedUnits),
          unitCount: latestContext.unitCount,
          unitLabel: latestContext.unitLabel,
        };
      } else {
        let locale;
        if (latestContext.format === "pdf") {
          const labelUnits = Object.entries(PDF_VISIBLE_LABELS).map(([id, text]) => ({ id, text }));
          const prompt = buildVisibleTextTranslationPrompt(labelUnits, translationLanguage);
          const translatedLabels = await generateStructuredModel({
            provider,
            model,
            system: prompt,
            userMessage: `Translate all PDF labels into ${translationLanguage.name}.`,
            structuredOutput: visibleTextTranslationOutput,
            parse: (raw) => parseVisibleTextTranslations(raw, labelUnits),
            outputLabel: `${translationLanguage.name} PDF labels`,
            signal: request.signal,
          });
          locale = {
            languageCode: translationLanguage.code,
            direction: translationLanguage.direction,
            labels: Object.fromEntries(translatedLabels.map((entry) => [entry.id, entry.text])),
          };
        }
        rendered = await renderArtifact({
          format: latestContext.format,
          model: translatedModel as ArtifactContentModel,
          version,
          sources: [],
          template: inheritedTemplate,
          locale,
        });
        rendered.filename = translatedArtifactFileName(latestContext.filename, translationLanguage.code, version);
      }
      await validateArtifact(rendered, inheritedTemplate);
      const artifactId = crypto.randomUUID();
      const objectKey = `artifacts/${userId}/${body.chatId}/${artifactId}/${rendered.filename}`;
      const bucket = getRuntimeBindings().FILES;
      if (!bucket) throw new Error("Artifact storage is unavailable: the FILES binding is not configured.");
      await bucket.put(objectKey, rendered.bytes, {
        httpMetadata: { contentType: rendered.mimeType, contentDisposition: `attachment; filename="${rendered.filename}"` },
        customMetadata: {
          userId,
          chatId: body.chatId,
          messageId: body.assistantMessageId,
          version: String(version),
          format: rendered.format,
          operation: lineage.operation,
          contentLanguage: lineage.contentLanguage,
          sourceArtifactId: lineage.parentArtifactId,
          ...(lineage.templateSourceId ? { templateSourceId: lineage.templateSourceId } : {}),
        },
      });
      const responseText = translatedResponse(latestContext.format, translationLanguage, languageContext.communicationLanguage);
      await saveArtifact({
        userId,
        chatId: body.chatId,
        messageId: body.assistantMessageId,
        projectId: body.projectId,
        chatTitle: body.chatTitle ?? (isArtifactContentModel(translatedModel) ? artifactTitle(translatedModel) : latestContext.filename.replace(/\.[^.]+$/, "")),
        audience: translatedModel && typeof translatedModel === "object" && "audience" in translatedModel && typeof translatedModel.audience === "string" ? translatedModel.audience : body.audience ?? "Management",
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
        model: translatedModel,
        ...lineage,
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
    const requestedFormat = detectArtifactRequest(body.message, latest?.format) ?? (selectedTemplate ? templateOutputFormat(selectedTemplate) : null);

    if (requestedFormat) {
      if (!body.chatId || !body.assistantMessageId) {
        return Response.json({ error: "Chat and message IDs are required for artifact generation." }, { status: 400 });
      }
      const previous = await loadLatestArtifactModel(userId, body.chatId, requestedFormat).catch(() => null);
      const referencedArtifact = modifiesExistingArtifact && latestContext && isArtifactContentModel(latestContext.model)
        ? latestContext
        : null;
      const currentArtifactModel = referencedArtifact?.model
        ?? (modifiesExistingArtifact ? previous?.model : null);
      const inheritedTemplate = !selectedTemplate && modifiesExistingArtifact && latestContext?.templateSourceId
        ? await hydrateStoredTemplate(userId, latestContext.templateSourceId)
        : null;
      const effectiveTemplate = selectedTemplate
        ?? (inheritedTemplate && templateOutputFormat(inheritedTemplate) === requestedFormat ? inheritedTemplate : null);
      const parentArtifactId = referencedArtifact?.artifactId ?? (modifiesExistingArtifact ? previous?.artifactId : null);
      const operation = parentArtifactId
        ? referencedArtifact && referencedArtifact.format !== requestedFormat ? "convert" as const : "regenerate" as const
        : "generate" as const;
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
          templateDescription: effectiveTemplate ? { fileName: effectiveTemplate.fileName, fileType: effectiveTemplate.fileType, layoutModel: effectiveTemplate.layoutModel } : null,
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
        const responseText = languageContext.communicationLanguage === "de"
          ? `Ich habe ${formatName[rendered.format] === "Word document" ? "das Word-Dokument" : formatName[rendered.format] === "PowerPoint presentation" ? "die PowerPoint-Präsentation" : `die ${formatName[rendered.format]}`} aus der referenzierten Antwort ${version > 1 ? "aktualisiert" : "erstellt"}${existingContent.editableBlockIds.length ? " und nur den ausdrücklich angeforderten Text geändert" : "; der Wortlaut blieb erhalten"}.`
          : `I ${version > 1 ? "updated" : "created"} the ${formatName[rendered.format]} from the referenced response${existingContent.editableBlockIds.length ? ", changing only the explicitly requested text" : " with its wording preserved"}.`;
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
          parentArtifactId,
          contentLanguage: languageContext.outputLanguage.code,
          templateSourceId: effectiveTemplate?.sourceId,
          operation,
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
        currentModel: currentArtifactModel as ArtifactContentModel | null | undefined,
        reconciliation,
        template: effectiveTemplate,
        languageContext,
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
      let pdfLocale;
      if (requestedFormat === "pdf") {
        let labels: Record<string, string> | undefined;
        if (languageContext.outputLanguage.code !== "en") {
          const labelUnits = Object.entries(PDF_VISIBLE_LABELS).map(([id, text]) => ({ id, text }));
          const translatedLabels = await generateStructuredModel({
            provider,
            model,
            system: buildVisibleTextTranslationPrompt(labelUnits, languageContext.outputLanguage),
            userMessage: `Translate all PDF labels into ${languageContext.outputLanguage.name}.`,
            structuredOutput: visibleTextTranslationOutput,
            parse: (raw) => parseVisibleTextTranslations(raw, labelUnits),
            outputLabel: `${languageContext.outputLanguage.name} PDF labels`,
            signal: request.signal,
          });
          labels = Object.fromEntries(translatedLabels.map((entry) => [entry.id, entry.text]));
        }
        pdfLocale = { languageCode: languageContext.outputLanguage.code, direction: languageContext.outputLanguage.direction, labels };
      }
      let rendered = await renderArtifact({
        format: requestedFormat,
        model: artifactModel,
        version,
        sources: evidenceSources,
        template: effectiveTemplate,
        locale: pdfLocale,
      });
      if ((requestedFormat === "docx" || requestedFormat === "pptx" || requestedFormat === "xlsx") && languageContext.outputLanguage.code !== "en") {
        const sourceUnits = extractOfficeVisibleText(rendered.bytes, requestedFormat);
        const translatedUnits: VisibleTextUnit[] = [];
        for (const batch of translationBatches(sourceUnits)) {
          const translatedBatch = await generateStructuredModel({
            provider,
            model,
            system: buildVisibleTextTranslationPrompt(batch, languageContext.outputLanguage),
            userMessage: `Translate all ${batch.length} visible text units into ${languageContext.outputLanguage.name}.`,
            structuredOutput: visibleTextTranslationOutput,
            parse: (raw) => parseVisibleTextTranslations(raw, batch),
            outputLabel: `${languageContext.outputLanguage.name} visible document text`,
            signal: request.signal,
          });
          translatedUnits.push(...translatedBatch);
        }
        rendered = {
          ...rendered,
          bytes: translateOfficeArtifact({
            bytes: rendered.bytes,
            format: requestedFormat,
            language: languageContext.outputLanguage,
            sourceUnits,
            translatedUnits,
          }),
        };
      }
      if (requestedFormat === "html" && languageContext.outputLanguage.code !== "en") {
        const sourceUnits = extractHtmlVisibleText(rendered.bytes);
        const translatedUnits: VisibleTextUnit[] = [];
        for (const batch of translationBatches(sourceUnits)) {
          const translatedBatch = await generateStructuredModel({
            provider,
            model,
            system: buildVisibleTextTranslationPrompt(batch, languageContext.outputLanguage),
            userMessage: `Translate all ${batch.length} visible HTML text units into ${languageContext.outputLanguage.name}.`,
            structuredOutput: visibleTextTranslationOutput,
            parse: (raw) => parseVisibleTextTranslations(raw, batch),
            outputLabel: `${languageContext.outputLanguage.name} visible HTML text`,
            signal: request.signal,
          });
          translatedUnits.push(...translatedBatch);
        }
        rendered = { ...rendered, bytes: translateHtmlArtifact(rendered.bytes, sourceUnits, translatedUnits) };
      }
      await validateArtifact(rendered, effectiveTemplate);
      const objectKey = `artifacts/${userId}/${body.chatId}/${artifactId}/${rendered.filename}`;
      const bucket = getRuntimeBindings().FILES;
      if (!bucket) throw new Error("Artifact storage is unavailable: the FILES binding is not configured.");
      await bucket.put(objectKey, rendered.bytes, {
        httpMetadata: { contentType: rendered.mimeType, contentDisposition: `attachment; filename="${rendered.filename}"` },
        customMetadata: {
          userId,
          chatId: body.chatId,
          messageId: body.assistantMessageId,
          version: String(version),
          format: rendered.format,
          operation,
          contentLanguage: languageContext.outputLanguage.code,
          ...(parentArtifactId ? { sourceArtifactId: parentArtifactId } : {}),
          ...(effectiveTemplate?.sourceId ? { templateSourceId: effectiveTemplate.sourceId } : {}),
        },
      });
      const audience = "audience" in artifactModel ? artifactModel.audience : body.audience ?? "Management";
      const responseText = generatedResponse({
        communicationLanguage: languageContext.communicationLanguage,
        updated: Boolean(parentArtifactId),
        title: artifactTitle(artifactModel),
        format: requestedFormat,
        unitCount: rendered.unitCount,
        unitLabel: rendered.unitLabel,
        audience,
        templateFileName: effectiveTemplate?.fileName,
      });
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
        parentArtifactId,
        contentLanguage: languageContext.outputLanguage.code,
        templateSourceId: effectiveTemplate?.sourceId,
        operation,
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
      communicationLanguage: languageContext.communicationLanguage,
    });
    const stream = await provider.stream({
      model,
      system,
      messages: [...(body.history ?? []).slice(-20).map(({ role, content }) => ({ role, content })), { role: "user", content: body.message }],
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
    const message = error instanceof Error ? error.message : communicationLanguage === "de" ? "Die Antwort konnte nicht erstellt werden." : "Unable to generate the response.";
    return Response.json({ error: message }, { status: 503 });
  }
}
