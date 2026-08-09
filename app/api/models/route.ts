import { getModelRegistry } from "../../lib/models";

export const dynamic = "force-dynamic";

export async function GET() {
  return Response.json({ models: getModelRegistry() });
}
