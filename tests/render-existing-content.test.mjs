import assert from "node:assert/strict";
import test from "node:test";
import { unzipSync } from "fflate";
import { PDFDocument } from "pdf-lib";
import {
  applyBlockEdits,
  assertLockedBlocksUnchanged,
  assertRenderedTextIntegrity,
  defaultExistingContentDesignPlan,
  deriveLockedMetrics,
  existingContentDesignStructuredOutput,
  normalizedText,
  parseExistingContent,
  parseExistingContentDesignPlan,
  resolveExistingContentRequest,
} from "../app/lib/existing-content.ts";
import { renderExistingContent } from "../app/lib/renderers/existing-content.ts";

const history = (content) => [
  { role: "user", content: "Summarize the integration." },
  { role: "assistant", content },
];

function xmlText(bytes, path, separator = "\n") {
  const xml = new TextDecoder().decode(unzipSync(bytes)[path]);
  return [...xml.matchAll(/<(?:w:t|a:t)(?:\s[^>]*)?>([\s\S]*?)<\/(?:w:t|a:t)>/g)]
    .map((match) => match[1].replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'"))
    .join(separator);
}

function visibleHtml(html) {
  return html.replace(/<style[\s\S]*?<\/style>/gi, " ").replace(/<[^>]+>/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'");
}

test("Word export resolves this text and renders the prior answer without an LLM rewrite", async () => {
  const source = "Overall integration progress is 78%.\nERP go-live remains at risk.";
  const request = resolveExistingContentRequest({ message: "Generate a Word file with this text.", format: "docx", history: history(source) });
  assert.ok(request);
  assert.equal(request.generationMode, "render_existing_content");
  assert.deepEqual(request.editableBlockIds, []);
  assert.equal(request.blocks.length, 1);
  assert.equal(request.blocks[0].text, source);

  const rendered = await renderExistingContent({ format: "docx", blocks: request.blocks, version: 1 });
  assertRenderedTextIntegrity(request.blocks, rendered.renderedTextBlocks);
  const documentText = xmlText(rendered.bytes, "word/document.xml");
  assert.ok(normalizedText(documentText).includes(normalizedText(source)));
});

test("PowerPoint export keeps previous-answer wording while distributing blocks across slides", async () => {
  const source = "# Board update\n\nOverall integration progress is **78%**.\n\n## Risk\n\nERP go-live remains at risk.";
  const request = resolveExistingContentRequest({ message: "Turn your previous answer into a PowerPoint.", format: "pptx", history: history(source) });
  assert.ok(request);
  const rendered = await renderExistingContent({ format: "pptx", blocks: request.blocks, version: 1 });
  assert.deepEqual(rendered.renderedTextBlocks, ["Board update", "Overall integration progress is 78%.", "Risk", "ERP go-live remains at risk."]);
  const zip = unzipSync(rendered.bytes);
  const slideText = Object.keys(zip).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path)).sort().map((path) => xmlText(rendered.bytes, path, "")).join("\n");
  for (const block of request.blocks) assert.ok(slideText.includes(block.text));
});

test("Markdown becomes native heading, bullet, bold, and table formatting", async () => {
  const source = "## Integration status\n\n- Overall integration progress: **78%**\n- ERP go-live: **At Risk**\n\n| Metric | Value |\n| --- | ---: |\n| Budget | €2.50m |\n| Forecast | €2.61m |";
  const blocks = parseExistingContent(source);
  assert.deepEqual(blocks.map((block) => block.kind), ["heading", "bullet", "bullet", "table"]);
  assert.equal(blocks[1].runs.some((run) => run.bold && run.text === "78%"), true);
  const rendered = await renderExistingContent({ format: "docx", blocks, version: 1 });
  const documentXml = new TextDecoder().decode(unzipSync(rendered.bytes)["word/document.xml"]);
  assert.match(documentXml, /<w:pStyle w:val="Heading2"/);
  assert.match(documentXml, /<w:numPr>/);
  assert.match(documentXml, /<w:b\/>/);
  assert.ok((documentXml.match(/<w:tbl>/g) ?? []).length >= 2);
  assert.doesNotMatch(documentXml, /\*\*|##|\|\s*---/);
});

test("professional Word composition remains active for locked report text", async () => {
  const source = "# Project Aurora — Board Integration Summary\n\n## Executive summary\n\nOverall integration progress is **78%**. ERP go-live remains at risk.\n\n## Financial position\n\n| Metric | Value |\n| --- | ---: |\n| Budget | €2.50m |\n| Forecast | €2.61m |\n| Variance | +€0.11m |";
  const blocks = parseExistingContent(source);
  const rendered = await renderExistingContent({ format: "docx", blocks, version: 1 });
  const zip = unzipSync(rendered.bytes);
  const documentXml = new TextDecoder().decode(zip["word/document.xml"]);
  assert.ok(zip["word/styles.xml"]);
  assert.ok(zip["word/header1.xml"]);
  assert.match(documentXml, /<w:pStyle w:val="Heading2"/);
  assert.match(documentXml, /w:fill="000000"/);
  assert.match(documentXml, /w:fill="F1F6E4"/);
  assert.ok((documentXml.match(/<w:tbl>/g) ?? []).length >= 4, "title, executive callout, native data table, and KPI/chart table should be present");
  assertRenderedTextIntegrity(blocks, rendered.renderedTextBlocks);
});

test("locked quantitative facts can drive a supported visualization without inventing values", async () => {
  const blocks = parseExistingContent("Synergy target: €37.0m\nForecast: €28.7m\nGap: €8.3m");
  const metrics = deriveLockedMetrics(blocks);
  assert.deepEqual(metrics.map(({ label, value }) => ({ label, value })), [
    { label: "Synergy target", value: "€37.0m" },
    { label: "Forecast", value: "€28.7m" },
    { label: "Gap", value: "€8.3m" },
  ]);
  const plan = defaultExistingContentDesignPlan(blocks);
  assert.equal(plan.visualizations[0].type, "bar_chart");
  assert.equal(JSON.stringify(existingContentDesignStructuredOutput.schema).includes("replacement_text"), false);
  const rendered = await renderExistingContent({ format: "docx", blocks, version: 1, plan });
  const text = xmlText(rendered.bytes, "word/document.xml");
  for (const metric of metrics) {
    assert.match(text, new RegExp(metric.value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
    assert.ok(blocks.some((block) => block.text.includes(metric.label) && block.text.includes(metric.value)));
  }
  assert.match(text, /KEY METRICS/);
  assertRenderedTextIntegrity(blocks, rendered.renderedTextBlocks);
});

test("the design planner can place blocks but cannot return replacement narrative", () => {
  const blocks = parseExistingContent("## Executive summary\n\nOverall integration progress is 78%.");
  const valid = JSON.stringify({
    style: "consulting",
    placements: [
      { blockIds: ["block_1"], component: "section_heading", page: 1, column: "full", emphasis: "standard" },
      { blockIds: ["block_2"], component: "executive_message", page: 1, column: "full", emphasis: "high" },
    ],
    visualizations: [{ type: "kpi_cards", sourceBlockIds: ["block_2"], page: 1 }],
  });
  const plan = parseExistingContentDesignPlan(valid, blocks);
  assert.equal(plan.placements[1].component, "executive_message");
  assert.throws(() => parseExistingContentDesignPlan(JSON.stringify({ ...JSON.parse(valid), replacement_text: "Integration advanced to 78%." }), blocks), /forbidden/i);
  assert.throws(() => parseExistingContentDesignPlan(JSON.stringify({
    ...JSON.parse(valid),
    placements: [{ blockIds: ["block_1", "block_2"], component: "body", page: 1, column: "full", emphasis: "standard", text: "Paraphrased" }],
  }), blocks), /forbidden/i);
});

test("PDF export adds no analysis, recommendations, or rewritten headings", async () => {
  const source = "# Integration status\n\nNo new analysis.\n\n## Confirmed risk\n\nERP go-live remains at risk.";
  const request = resolveExistingContentRequest({ message: "Create a PDF from the text above.", format: "pdf", history: history(source) });
  assert.ok(request);
  const rendered = await renderExistingContent({ format: "pdf", blocks: request.blocks, version: 1 });
  assertRenderedTextIntegrity(request.blocks, rendered.renderedTextBlocks);
  assert.deepEqual(rendered.renderedTextBlocks, request.blocks.map((block) => block.text));
  assert.equal(rendered.renderedTextBlocks.some((text) => /recommendation|management insight/i.test(text)), false);
  assert.ok((await PDFDocument.load(rendered.bytes)).getPageCount() >= 1);
});

test("a section-specific edit unlocks only that section and preserves every other block", async () => {
  const source = "# Board report\n\n## Executive summary\n\nThis is the longer executive summary.\n\n## Risks\n\nERP go-live remains at risk.";
  const request = resolveExistingContentRequest({ message: "Turn this into Word and shorten only the executive summary.", format: "docx", history: history(source) });
  assert.ok(request);
  assert.deepEqual(request.editableBlockIds, ["block_3"]);
  const finalBlocks = applyBlockEdits(request.blocks, [{ id: "block_3", text: "Short executive summary." }]);
  assertLockedBlocksUnchanged(request.blocks, finalBlocks);
  assert.equal(finalBlocks.find((block) => block.text === "ERP go-live remains at risk.")?.locked, true);
  const rendered = await renderExistingContent({ format: "docx", blocks: finalBlocks, version: 1 });
  assertRenderedTextIntegrity(finalBlocks, rendered.renderedTextBlocks);
});

test("new report generation from uploaded files remains in content-generation mode", () => {
  const request = resolveExistingContentRequest({
    message: "Create a Board report from all uploaded files.",
    format: "docx",
    history: history("An earlier chat answer."),
  });
  assert.equal(request, null);
  assert.equal(resolveExistingContentRequest({
    message: "Create a Word report from this uploaded file.",
    format: "docx",
    history: history("An earlier chat answer."),
  }), null);
});

test("HTML export preserves every visible text block", async () => {
  const source = "# Board update\n\nExact wording with **€2.61m** and 20 August 2026.\n\n- Keep this recommendation unchanged.";
  const request = resolveExistingContentRequest({ message: "Export the previous response as HTML using the same content.", format: "html", history: history(source) });
  assert.ok(request);
  const rendered = await renderExistingContent({ format: "html", blocks: request.blocks, version: 1 });
  assertRenderedTextIntegrity(request.blocks, rendered.renderedTextBlocks);
  const html = new TextDecoder().decode(rendered.bytes);
  const visible = normalizedText(visibleHtml(html));
  for (const block of request.blocks) assert.ok(visible.includes(normalizedText(block.text)));
  assert.match(html, /<strong>€2\.61m<\/strong>/);
  assert.doesNotMatch(visible, /\*\*|##/);
});
