import { extractFile, createCoverageCheck } from "../../lib/ingestion/pipeline";
import { authenticatedUserId } from "../../lib/persistence";
import type { ExtractedDocument } from "../../lib/ingestion/types";

export const dynamic = "force-dynamic";

const MAX_FILE_SIZE = 25 * 1024 * 1024;
const MAX_FILES = 20;
const MAX_EXCERPT = 20_000;

type RuntimeEnv = { FILES?: R2Bucket };

function isUpload(value: FormDataEntryValue): value is File {
  return typeof value === "object" && value !== null && "name" in value && "size" in value && "arrayBuffer" in value;
}

function safeName(value: string) {
  return value.replace(/[^a-zA-Z0-9._-]+/g, "-").slice(-120) || "upload";
}

function compact(document: ExtractedDocument) {
  return {
    fileId: document.fileId,
    fileName: document.fileName,
    fileType: document.fileType,
    mediaType: document.mediaType,
    sizeBytes: document.sizeBytes,
    status: document.status,
    metadata: document.metadata,
    rawText: document.rawText.slice(0, MAX_EXCERPT),
    structuredElements: document.structuredElements.slice(0, 50),
    extractionWarnings: document.extractionWarnings,
  };
}

export async function POST(request: Request) {
  const form = await request.formData();
  const files = form.getAll("files").filter(isUpload);
  if (!files.length) return Response.json({ error: "Attach at least one file." }, { status: 400 });
  if (files.length > MAX_FILES) return Response.json({ error: `A maximum of ${MAX_FILES} files is allowed per message.` }, { status: 413 });
  const oversized = files.find((file) => file.size > MAX_FILE_SIZE);
  if (oversized) return Response.json({ error: `${oversized.name} exceeds the 25 MB upload limit.` }, { status: 413 });

  const requestedId = String(form.get("fileId") ?? "");
  const userId = authenticatedUserId(request);
  const documents = await Promise.all(files.map(async (file, index) => {
    const fileId = files.length === 1 && requestedId ? requestedId : crypto.randomUUID();
    const objectKey = `staged/${userId}/${fileId}`;
    try {
      const { env } = await import("cloudflare:workers");
      await (env as RuntimeEnv).FILES?.put(objectKey, await file.arrayBuffer(), {
        httpMetadata: { contentType: file.type || "application/octet-stream" },
        customMetadata: { fileName: safeName(file.name), userId, uploadIndex: String(index) },
      });
    } catch {
      // Local tests may not provide R2. Extraction remains available in that environment.
    }
    const document = await extractFile(file, fileId);
    return { ...compact(document), objectKey };
  }));
  const coverage = createCoverageCheck(documents);
  return Response.json({ documents, coverage });
}
