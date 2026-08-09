import { extractFiles, createCoverageCheck } from "../../lib/ingestion/pipeline";

export const dynamic = "force-dynamic";

const MAX_FILE_SIZE = 25 * 1024 * 1024;
const MAX_FILES = 20;

export async function POST(request: Request) {
  const form = await request.formData();
  const files = form.getAll("files").filter((value): value is File => value instanceof File);
  if (!files.length) return Response.json({ error: "Attach at least one file." }, { status: 400 });
  if (files.length > MAX_FILES) return Response.json({ error: `A maximum of ${MAX_FILES} files is allowed per message.` }, { status: 413 });
  const oversized = files.find((file) => file.size > MAX_FILE_SIZE);
  if (oversized) return Response.json({ error: `${oversized.name} exceeds the 25 MB upload limit.` }, { status: 413 });

  const documents = await extractFiles(files);
  const coverage = createCoverageCheck(documents);
  return Response.json({ documents, coverage });
}
