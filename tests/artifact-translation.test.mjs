import assert from "node:assert/strict";
import test from "node:test";
import { strFromU8, unzipSync } from "fflate";
import { PDFDocument } from "pdf-lib";
import {
  SUPPORTED_TRANSLATION_LANGUAGES,
  assertTranslatedModelIntegrity,
  buildArtifactTranslationPrompt,
  detectTranslationRequest,
  extractOfficeVisibleText,
  translatedArtifactFileName,
  translatedArtifactLineage,
  translateOfficeArtifact,
} from "../app/lib/artifact-translation.ts";
import { renderWordDocument } from "../app/lib/renderers/word.ts";
import { renderPdfReport, PDF_VISIBLE_LABELS } from "../app/lib/renderers/pdf.ts";
import { renderPresentation } from "../app/lib/presentation.ts";
import { fillPdfTemplate } from "../app/lib/report-template.ts";

const report = {
  title: "Integration status report",
  subtitle: "Decision-ready view",
  audience: "Steering Committee",
  reportingPeriod: "Week 39",
  executiveSummary: "Overall progress is stable; one decision is required.",
  sections: [{
    name: "Status",
    title: "Delivery remains on track",
    keyMessage: "The technology dependency needs attention.",
    type: "status",
    items: [{
      label: "Technology integration",
      value: "On track",
      detail: "Migration is proceeding to plan.",
      implication: "The cutover date remains achievable.",
      recommendation: "Confirm the accountable owner.",
      owner: "PMI lead",
      deadline: "Friday",
      status: "green",
      evidenceType: "fact",
      sourceRefs: ["source-1"],
    }],
    sourceNotes: ["Weekly status source"],
  }],
};

const presentation = {
  title: "Integration status report",
  subtitle: "Decision-ready view",
  audience: "Steering Committee",
  executiveSummary: "Overall progress is stable; one decision is required.",
  slides: [{
    title: "Delivery remains on track",
    kicker: "Executive status",
    keyMessage: "The technology dependency needs attention.",
    layout: "summary",
    items: [{ label: "Technology integration", detail: "Migration is proceeding to plan.", implication: "The cutover date remains achievable.", status: "green", evidenceType: "fact", sourceRefs: ["source-1"] }],
    sourceNotes: ["Weekly status source"],
  }],
};

function language(code) {
  const value = SUPPORTED_TRANSLATION_LANGUAGES.find((candidate) => candidate.code === code);
  assert.ok(value, `missing language ${code}`);
  return value;
}

function fakeTranslations(units, prefix) {
  return units.map((unit) => ({ id: unit.id, text: `${prefix}${unit.text}` }));
}

test("recognizes artifact translation and regeneration requests for every supported language", () => {
  for (const target of SUPPORTED_TRANSLATION_LANGUAGES) {
    assert.equal(detectTranslationRequest(`Now give me the ${target.name} version.`, true)?.code, target.code);
  }
  assert.equal(detectTranslationRequest("Translate it all into English and generate the report one more time.", true)?.code, "en");
  assert.equal(detectTranslationRequest("Now make the same report in Japanese.", true)?.code, "ja");
  assert.equal(detectTranslationRequest("Translate it into Chinese.", true)?.code, "zh-CN");
  assert.equal(detectTranslationRequest("Translate it into English.", false), null, "there must be a prior artifact to resolve 'it'");
});

test("translation planning is language-independent and protects exact report structure", () => {
  const prompt = buildArtifactTranslationPrompt({ format: "docx", language: language("ja"), sourceModel: report });
  assert.match(prompt, /not a rewrite/i);
  assert.match(prompt, /Do not summarize, shorten, expand, omit, reorder, or add/i);
  assert.match(prompt, /sourceRefs/);
  const translated = structuredClone(report);
  translated.title = "統合状況レポート";
  translated.sections[0].items[0].detail = "移行は計画どおりに進んでいます。";
  assert.doesNotThrow(() => assertTranslatedModelIntegrity(report, translated, "docx"));
  translated.sections[0].items.pop();
  assert.throws(() => assertTranslatedModelIntegrity(report, translated, "docx"), /item count changed/);
  const changedNumber = structuredClone(report);
  changedNumber.reportingPeriod = "Week 40";
  assert.throws(() => assertTranslatedModelIntegrity(report, changedNumber, "docx"), /numeric facts or dates changed/);
});

test("translates every Word-visible unit while preserving package structure and Deloitte branding", async () => {
  const source = await renderWordDocument(report, []);
  const sourceBytes = new Uint8Array(source.bytes);
  const units = extractOfficeVisibleText(sourceBytes, "docx");
  assert.ok(units.some((unit) => unit.id.startsWith("word/header")), "headers are included");
  assert.ok(units.some((unit) => unit.id.startsWith("word/footer")), "footers are included");

  const translated = translateOfficeArtifact({
    bytes: sourceBytes,
    format: "docx",
    language: language("ar"),
    sourceUnits: units,
    translatedUnits: fakeTranslations(units, "عربي "),
  });
  const outputUnits = extractOfficeVisibleText(translated, "docx");
  assert.equal(outputUnits.length, units.length);
  assert.ok(outputUnits.every((unit) => unit.text.startsWith("عربي ")), "all visible units are replaced");
  const sourceZip = unzipSync(sourceBytes);
  const outputZip = unzipSync(translated);
  assert.deepEqual(Object.keys(outputZip).sort(), Object.keys(sourceZip).sort(), "the original DOCX package structure is preserved");
  for (const path of Object.keys(sourceZip).filter((path) => path.startsWith("word/media/"))) assert.deepEqual(outputZip[path], sourceZip[path], `${path} branding asset changed`);
  const documentXml = strFromU8(outputZip["word/document.xml"]);
  assert.match(documentXml, /<w:bidi\/>/);
  assert.match(documentXml, /w:eastAsia="Arial"/);
});

test("supports consecutive Word translations and always derives from the latest generated artifact", async () => {
  const source = new Uint8Array((await renderWordDocument(report, [])).bytes);
  const firstUnits = extractOfficeVisibleText(source, "docx");
  const japanese = translateOfficeArtifact({ bytes: source, format: "docx", language: language("ja"), sourceUnits: firstUnits, translatedUnits: fakeTranslations(firstUnits, "日本語 ") });
  const secondUnits = extractOfficeVisibleText(japanese, "docx");
  assert.ok(secondUnits.every((unit) => unit.text.startsWith("日本語 ")));
  const chinese = translateOfficeArtifact({ bytes: japanese, format: "docx", language: language("zh-CN"), sourceUnits: secondUnits, translatedUnits: fakeTranslations(secondUnits, "中文 ") });
  assert.ok(extractOfficeVisibleText(chinese, "docx").every((unit) => unit.text.startsWith("中文 日本語 ")), "the second translation used the first translated artifact as its source");
  assert.equal(translatedArtifactFileName("SteerCo_Statusbericht_v2.docx", "zh-CN", 3), "SteerCo_Statusbericht_zh-CN_v3.docx");
  const englishLineage = translatedArtifactLineage({ artifactId: "german-report", templateSourceId: "deloitte-template" }, "en");
  const japaneseLineage = translatedArtifactLineage({ artifactId: "english-report", templateSourceId: englishLineage.templateSourceId }, "ja");
  assert.deepEqual(englishLineage, { parentArtifactId: "german-report", templateSourceId: "deloitte-template", operation: "translate", contentLanguage: "en" });
  assert.deepEqual(japaneseLineage, { parentArtifactId: "english-report", templateSourceId: "deloitte-template", operation: "translate", contentLanguage: "ja" });
});

test("translates complete PowerPoint text including notes while preserving slide layout assets", async () => {
  const source = await renderPresentation(presentation);
  const units = extractOfficeVisibleText(source, "pptx");
  assert.ok(units.some((unit) => unit.id.startsWith("ppt/notesSlides/")), "speaker notes are included");
  const translated = translateOfficeArtifact({ bytes: source, format: "pptx", language: language("ja"), sourceUnits: units, translatedUnits: fakeTranslations(units, "日本語 ") });
  const outputUnits = extractOfficeVisibleText(translated, "pptx");
  assert.equal(outputUnits.length, units.length);
  assert.ok(outputUnits.every((unit) => unit.text.startsWith("日本語 ")));
  const sourceZip = unzipSync(source);
  const outputZip = unzipSync(translated);
  assert.equal(Object.keys(outputZip).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path)).length, Object.keys(sourceZip).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path)).length);
  for (const path of Object.keys(sourceZip).filter((path) => path.startsWith("ppt/media/"))) assert.deepEqual(outputZip[path], sourceZip[path], `${path} branding asset changed`);
  assert.match(strFromU8(outputZip["ppt/slides/slide1.xml"]), /typeface="Yu Gothic"/);
});

test("renders PDFs with embedded Simplified Chinese, Japanese, and Arabic fonts and usable RTL layout", async () => {
  const cases = [
    ["zh-CN", "集成状态报告", "关键决策需要本周完成。", "总体进展保持稳定"],
    ["ja", "統合状況レポート", "重要な決定が今週必要です。", "全体の進捗は安定しています"],
    ["ar", "تقرير حالة الاندماج", "يجب اتخاذ القرار هذا الأسبوع.", "التقدم العام مستقر"],
  ];
  for (const [code, title, summary, sectionTitle] of cases) {
    const target = language(code);
    const model = structuredClone(report);
    model.title = title;
    model.executiveSummary = summary;
    model.sections[0].title = sectionTitle;
    const rendered = await renderPdfReport(model, [], { languageCode: code, direction: target.direction, labels: PDF_VISIBLE_LABELS });
    const pdf = await PDFDocument.load(rendered.bytes);
    assert.ok(pdf.getPageCount() >= 1, `${code} PDF has no pages`);
    assert.ok(rendered.bytes.byteLength > 10_000, `${code} font was not embedded`);
  }
});

test("regenerates an Arabic PDF from a prior Deloitte PDF template with its page geometry intact", async () => {
  const source = await renderPdfReport(report, []);
  const sourcePdf = await PDFDocument.load(source.bytes);
  const arabic = structuredClone(report);
  arabic.title = "تقرير حالة الاندماج";
  arabic.executiveSummary = "يجب اتخاذ القرار هذا الأسبوع.";
  arabic.sections[0].title = "التقدم العام مستقر";
  arabic.sections[0].items[0].detail = "تسير عملية الترحيل وفق الخطة.";
  const rendered = await fillPdfTemplate(source.bytes, arabic, {
    languageCode: "ar",
    direction: "rtl",
    labels: { implication: "الأثر", recommendation: "التوصية", ownerTiming: "المالك / التوقيت", by: "بحلول" },
  });
  const output = await PDFDocument.load(rendered.bytes);
  assert.equal(output.getPages()[0].getWidth(), sourcePdf.getPages()[0].getWidth());
  assert.equal(output.getPages()[0].getHeight(), sourcePdf.getPages()[0].getHeight());
  assert.ok(rendered.bytes.byteLength > 10_000);
});
