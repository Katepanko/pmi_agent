import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

async function loadWorker() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}-${Math.random()}`);
  return (await import(workerUrl.href)).default;
}

async function request(path = "/", init) {
  const worker = await loadWorker();
  return worker.fetch(
    new Request(`http://localhost${path}`, init),
    {
      ASSETS: { fetch: async () => new Response("Not found", { status: 404 }) },
    },
    { waitUntil() {}, passThroughOnException() {} },
  );
}

test("server-renders the PMI consulting workspace", async () => {
  const response = await request();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);

  const html = await response.text();
  assert.match(html, /<title>PMI Agent \| Decision-ready integration reporting<\/title>/i);
  assert.match(html, /PMI AGENT/i);
  assert.match(html, />User</i);
  assert.match(html, /PMI Agent/i);
  assert.match(html, /Executive report/i);
  assert.match(html, /Every applicable source is checked before synthesis/i);
  assert.match(html, /Content preview first/i);
  assert.match(html, /Source coverage complete/i);
  assert.doesNotMatch(html, /codex-preview|Your site is taking shape|react-loading-skeleton/i);
});

test("exposes configured model display names without leaking secrets", async () => {
  const response = await request("/api/models");
  assert.equal(response.status, 200);
  const body = await response.json();
  assert.deepEqual(body.models.map((model) => model.key), [
    "anthropic-sonnet",
    "anthropic-opus",
    "openai-gpt55",
    "openai-gpt56",
    "openai-mini",
  ]);
  assert.ok(body.models.every((model) => !("apiKey" in model)));
  assert.ok(body.models.every((model) => model.available === false));
});

test("returns a clear provider configuration error rather than a fabricated answer", async () => {
  const response = await request("/api/chat", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      modelKey: "openai-gpt56",
      message: "Create a Board report.",
      sources: [],
    }),
  });
  assert.equal(response.status, 503);
  const body = await response.json();
  assert.match(body.error, /OPENAI_API_KEY/);
  assert.match(body.error, /OPENAI_GPT56_MODEL/);
});

test("extracts multiple first-message files with compact provenance and coverage", async () => {
  const form = new FormData();
  form.append("files", new File([
    "Milestone,Owner,Due Date,Status\nERP design,Alex,2026-08-20,Amber\nPayroll cutover,Sam,2026-09-15,Green\n",
  ], "Masterplan.csv", { type: "text/csv" }));
  form.append("files", new File([
    "<html><body><h1>Weekly update</h1><p>ERP testing is blocked by missing access.</p></body></html>",
  ], "Weekly_Update.html", { type: "text/html" }));

  const response = await request("/api/extract", { method: "POST", body: form });
  assert.equal(response.status, 200);
  const body = await response.json();
  assert.equal(body.documents.length, 2);
  assert.equal(body.coverage.coverageComplete, true);
  assert.equal(body.coverage.sourcesConsidered.length, 2);
  assert.match(body.documents[0].structuredElements[0].location, /Masterplan\.csv → row 2/);
  assert.deepEqual(body.documents[0].structuredElements[0].structured.values, {
    Milestone: "ERP design",
    Owner: "Alex",
    "Due Date": "2026-08-20",
    Status: "Amber",
  });
  assert.match(body.documents[1].rawText, /ERP testing is blocked/);
  assert.ok(body.documents.every((document) => document.rawText.length <= 20_000));
});

test("keeps nine uploads isolated so one response cannot fail the entire batch", async () => {
  const responses = await Promise.all(Array.from({ length: 9 }, async (_, index) => {
    const form = new FormData();
    form.append("fileId", `source-${index}`);
    form.append("files", new File([`Item,Status\nWorkstream ${index},Green\n`], `Source_${index}.csv`, { type: "text/csv" }));
    return request("/api/extract", { method: "POST", body: form });
  }));
  assert.ok(responses.every((response) => response.status === 200));
  const bodies = await Promise.all(responses.map((response) => response.json()));
  assert.deepEqual(bodies.map((body) => body.documents[0].fileId), Array.from({ length: 9 }, (_, index) => `source-${index}`));
});

test("keeps grounding, source coverage, and tenant boundaries explicit in source", async () => {
  const [prompt, schema, packageJson, exporter] = await Promise.all([
    readFile(new URL("../app/lib/pmi-prompt.ts", import.meta.url), "utf8"),
    readFile(new URL("../db/schema.ts", import.meta.url), "utf8"),
    readFile(new URL("../package.json", import.meta.url), "utf8"),
    readFile(new URL("../app/lib/report-export.ts", import.meta.url), "utf8"),
  ]);

  assert.match(prompt, /Consider every source/i);
  assert.match(prompt, /Never invent numbers/i);
  assert.match(prompt, /AI-generated recommendation — validation required/i);
  assert.match(prompt, /Complete applicable source manifest/);
  assert.match(schema, /userId: text\("user_id"\)\.notNull\(\)/);
  assert.match(schema, /projectId: text\("project_id"\)/);
  assert.match(schema, /reportDraftVersions/);
  assert.doesNotMatch(packageJson, /react-loading-skeleton/);
  for (const format of ["pptx", "pdf", "xlsx", "docx", "html"]) assert.match(exporter, new RegExp(`\\b${format}\\b`));
});
