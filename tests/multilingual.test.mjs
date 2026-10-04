import assert from "node:assert/strict";
import test from "node:test";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { detectArtifactRequest } from "../app/lib/artifact-intent.ts";
import {
  detectTranslationRequest,
  extractHtmlVisibleText,
  extractOfficeVisibleText,
  translateHtmlArtifact,
  translateOfficeArtifact,
} from "../app/lib/artifact-translation.ts";
import { referencesExistingContent } from "../app/lib/existing-content.ts";
import {
  detectCommunicationLanguage,
  detectRequestedOutputLanguage,
  referencesPriorArtifact,
  resolveArtifactLanguageContext,
} from "../app/lib/language-context.ts";
import { buildGroundedPrompt } from "../app/lib/pmi-prompt.ts";
import { buildPresentationPlanningPrompt } from "../app/lib/presentation.ts";

test("German artifact requests map to the same formats as English requests", () => {
  assert.equal(detectArtifactRequest("Erstelle eine PowerPoint-Präsentation für das SteerCo."), "pptx");
  assert.equal(detectArtifactRequest("Generiere eine Excel-Arbeitsmappe."), "xlsx");
  assert.equal(detectArtifactRequest("Speichere das als Word-Datei."), "docx");
  assert.equal(detectArtifactRequest("Erstelle einen PDF-Bericht."), "pdf");
  assert.equal(detectArtifactRequest("Baue ein HTML-Dashboard."), "html");
  assert.equal(detectArtifactRequest("Nutze den letzten generierten Bericht.", "pptx"), "pptx");
});

test("communication language and artifact language are detected independently", () => {
  const germanRequestForEnglishArtifact = resolveArtifactLanguageContext({
    message: "Erstelle einen SteerCo-Bericht auf Englisch.",
    modifiesExistingArtifact: false,
  });
  assert.equal(germanRequestForEnglishArtifact.communicationLanguage, "de");
  assert.equal(germanRequestForEnglishArtifact.outputLanguage.code, "en");

  const englishRequestForGermanArtifact = resolveArtifactLanguageContext({
    message: "Create the same report in German.",
    sourceArtifactLanguageCode: "en",
    modifiesExistingArtifact: true,
  });
  assert.equal(englishRequestForGermanArtifact.communicationLanguage, "en");
  assert.equal(englishRequestForGermanArtifact.outputLanguage.code, "de");
});

test("new artifacts default to request language while revisions preserve source language", () => {
  assert.equal(resolveArtifactLanguageContext({
    message: "Erstelle eine neue Präsentation.",
    modifiesExistingArtifact: false,
  }).outputLanguage.code, "de");

  assert.equal(resolveArtifactLanguageContext({
    message: "Update the title and regenerate the file.",
    sourceArtifactLanguageCode: "de",
    modifiesExistingArtifact: true,
  }).outputLanguage.code, "de");
});

test("German and cross-language translation follow-ups resolve the prior artifact", () => {
  const cases = [
    ["Übersetze den gesamten Bericht ins Englische und generiere die Datei erneut.", "en"],
    ["Kannst du den ganzen Bericht jetzt auf Deutsch erstellen?", "de"],
    ["Erstelle mir die gleiche Präsentation auf Japanisch.", "ja"],
    ["Now give me the same report in Japanese.", "ja"],
  ];
  for (const [message, code] of cases) {
    assert.equal(referencesPriorArtifact(message), true, message);
    assert.equal(detectTranslationRequest(message, true)?.code, code, message);
  }
  assert.equal(detectTranslationRequest("Erstelle einen neuen Bericht auf Deutsch.", true), null);
});

test("output language detection retains the broader translation language set", () => {
  assert.equal(detectRequestedOutputLanguage("Erstelle die Präsentation auf Japanisch.")?.code, "ja");
  assert.equal(detectRequestedOutputLanguage("Generate the report in Arabic.")?.code, "ar");
  assert.equal(detectRequestedOutputLanguage("Übersetze die Datei ins vereinfachte Chinesisch.")?.code, "zh-CN");
});

test("German follow-ups can render a previous conversational response", () => {
  assert.equal(referencesExistingContent("Speichere das als Word-Datei."), true);
  assert.equal(referencesExistingContent("Nutze diesen Text für eine Präsentation."), true);
});

test("chat and artifact prompts carry separate language contracts", () => {
  const chatPrompt = buildGroundedPrompt({ sources: [], communicationLanguage: "de" });
  assert.match(chatPrompt, /Conversation language: German/);
  assert.match(chatPrompt, /affects only the conversational response/);

  const languageContext = resolveArtifactLanguageContext({
    message: "Erstelle eine Präsentation auf Englisch.",
    modifiesExistingArtifact: false,
  });
  const artifactPrompt = buildPresentationPlanningPrompt({
    request: "Erstelle eine Präsentation auf Englisch.",
    audience: "Steering Committee",
    sources: [],
    history: [],
    languageContext,
  });
  assert.match(artifactPrompt, /communication language is German/);
  assert.match(artifactPrompt, /exclusively in English \(en\)/);
  assert.match(artifactPrompt, /independent of both the conversation language/);
});

test("explicit communication-language requests do not become artifact-language defaults", () => {
  assert.equal(detectCommunicationLanguage("Bitte antworte mir auf Englisch."), "en");
  assert.equal(detectCommunicationLanguage("Please reply to me in German."), "de");
});

test("Excel and HTML visible system text can use the artifact output language", () => {
  const workbook = zipSync({
    "xl/sharedStrings.xml": strToU8('<?xml version="1.0"?><sst><si><t>Executive Summary</t></si><si><r><t>On </t></r><r><t>track</t></r></si></sst>'),
    "xl/worksheets/sheet1.xml": strToU8('<?xml version="1.0"?><worksheet><c t="inlineStr"><is><t>Management point</t></is></c></worksheet>'),
    "xl/styles.xml": strToU8('<?xml version="1.0"?><styleSheet><fonts><font><name val="Aptos"/></font></fonts></styleSheet>'),
  });
  const spreadsheetUnits = extractOfficeVisibleText(workbook, "xlsx");
  assert.deepEqual(spreadsheetUnits.map((unit) => unit.text), ["Executive Summary", "On track", "Management point"]);
  const german = detectRequestedOutputLanguage("auf Deutsch");
  assert.ok(german);
  const translatedWorkbook = unzipSync(translateOfficeArtifact({
    bytes: workbook,
    format: "xlsx",
    language: german,
    sourceUnits: spreadsheetUnits,
    translatedUnits: spreadsheetUnits.map((unit, index) => ({ ...unit, text: ["Management-Zusammenfassung", "Im Plan", "Management-Punkt"][index] })),
  }));
  assert.match(strFromU8(translatedWorkbook["xl/sharedStrings.xml"]), /Management-Zusammenfassung/);
  assert.match(strFromU8(translatedWorkbook["xl/worksheets/sheet1.xml"]), /Management-Punkt/);

  const html = new TextEncoder().encode("<!doctype html><html><head><style>.x{color:red}</style></head><body><h1>Executive Summary</h1><p>On track</p></body></html>");
  const htmlUnits = extractHtmlVisibleText(html);
  assert.deepEqual(htmlUnits.map((unit) => unit.text), ["Executive Summary", "On track"]);
  const translatedHtml = new TextDecoder().decode(translateHtmlArtifact(
    html,
    htmlUnits,
    htmlUnits.map((unit, index) => ({ ...unit, text: ["Management-Zusammenfassung", "Im Plan"][index] })),
  ));
  assert.match(translatedHtml, /<h1>Management-Zusammenfassung<\/h1>/);
  assert.match(translatedHtml, /\.x\{color:red\}/);
});
