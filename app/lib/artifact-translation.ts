import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import type { ArtifactContentModel, ArtifactFormat } from "./artifact.ts";
import {
  detectRequestedOutputLanguage,
  referencesPriorArtifact,
  type SupportedLanguage,
} from "./language-context.ts";

export { SUPPORTED_TRANSLATION_LANGUAGES, type SupportedLanguage } from "./language-context.ts";

export function detectTranslationRequest(message: string, hasPriorArtifact: boolean): SupportedLanguage | null {
  if (!hasPriorArtifact) return null;
  const normalized = message.normalize("NFKC").toLocaleLowerCase();
  const translationAction = /\b(?:translat(?:e|ed|ion)|übersetz\w*|uebersetz\w*)\b/iu.test(normalized);
  const translatedVersionCue = /\b(?:regenerat(?:e|ed)|version|same\s+(?:report|file|document|presentation|deck)|make\s+(?:it|this|the\s+same)|give\s+me|nochmals|erneut|wieder|fassung|gleich(?:e|en|er|es)?\s+(?:bericht|datei|dokument|präsentation|praesentation))\b/iu.test(normalized);
  if (!translationAction && !translatedVersionCue && !referencesPriorArtifact(message)) return null;
  return detectRequestedOutputLanguage(message);
}

export function buildArtifactTranslationPrompt(input: {
  format: ArtifactFormat;
  language: SupportedLanguage;
  sourceModel: ArtifactContentModel;
}) {
  return `You are translating an existing generated artifact into ${input.language.name} (${input.language.code}).

This is a faithful translation operation, not a rewrite or a new report:
- Translate every user-visible natural-language string in the supplied model, including titles, headings, labels, paragraphs, table content, captions, notes, headers, footers, and source notes.
- Preserve every fact, number, date, owner, status, recommendation, qualifier, and item. Do not summarize, shorten, expand, omit, reorder, or add information.
- Preserve the exact object structure and array lengths.
- Do not translate identifiers in sourceRefs or enum/control values such as status, evidenceType, layout, type, and templateSlide.
- Preserve proper names and trademarks unless they have a conventional ${input.language.name} form.
- Return the complete translated artifact model, never prose or Markdown.
${input.language.direction === "rtl" ? "- The target is right-to-left. Use natural Arabic wording; the renderer will apply RTL layout and fonts." : ""}

Artifact format: ${input.format}
Target language: ${input.language.name} (${input.language.locale})
Source artifact model: ${JSON.stringify(input.sourceModel)}`;
}

function invariant(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(`Translated artifact changed protected structure: ${message}`);
}

export function assertTranslatedModelIntegrity(source: ArtifactContentModel, translated: ArtifactContentModel, format: ArtifactFormat) {
  const numericLiterals = (value: unknown) => (JSON.stringify(value).match(/\d+(?:[.,]\d+)*/g) ?? []).map((literal) => literal.replace(",", ".")).sort();
  if (format === "pptx") {
    const before = source as Extract<ArtifactContentModel, { slides: unknown }>;
    const after = translated as Extract<ArtifactContentModel, { slides: unknown }>;
    invariant(Array.isArray(before.slides) && Array.isArray(after.slides), "slides are missing");
    invariant(before.slides.length === after.slides.length, "slide count changed");
    before.slides.forEach((slide, slideIndex) => {
      const candidate = after.slides[slideIndex];
      invariant(candidate && slide.items.length === candidate.items.length, `item count changed on slide ${slideIndex + 1}`);
      invariant(slide.layout === candidate.layout, `layout changed on slide ${slideIndex + 1}`);
      invariant(slide.templateSlide === candidate.templateSlide, `template mapping changed on slide ${slideIndex + 1}`);
      slide.items.forEach((item, itemIndex) => {
        const next = candidate.items[itemIndex];
        invariant(item.status === next.status, `status changed at slide ${slideIndex + 1}, item ${itemIndex + 1}`);
        invariant(item.evidenceType === next.evidenceType, `evidence type changed at slide ${slideIndex + 1}, item ${itemIndex + 1}`);
        invariant(JSON.stringify(item.sourceRefs ?? []) === JSON.stringify(next.sourceRefs ?? []), `source references changed at slide ${slideIndex + 1}, item ${itemIndex + 1}`);
      });
    });
    invariant(JSON.stringify(numericLiterals(source)) === JSON.stringify(numericLiterals(translated)), "numeric facts or dates changed");
    return;
  }
  const before = source as Extract<ArtifactContentModel, { sections: unknown }>;
  const after = translated as Extract<ArtifactContentModel, { sections: unknown }>;
  invariant(Array.isArray(before.sections) && Array.isArray(after.sections), "sections are missing");
  invariant(before.sections.length === after.sections.length, "section count changed");
  before.sections.forEach((section, sectionIndex) => {
    const candidate = after.sections[sectionIndex];
    invariant(candidate && section.items.length === candidate.items.length, `item count changed in section ${sectionIndex + 1}`);
    invariant(section.type === candidate.type, `section type changed in section ${sectionIndex + 1}`);
    section.items.forEach((item, itemIndex) => {
      const next = candidate.items[itemIndex];
      invariant(item.status === next.status, `status changed at section ${sectionIndex + 1}, item ${itemIndex + 1}`);
      invariant(item.evidenceType === next.evidenceType, `evidence type changed at section ${sectionIndex + 1}, item ${itemIndex + 1}`);
      invariant(JSON.stringify(item.sourceRefs ?? []) === JSON.stringify(next.sourceRefs ?? []), `source references changed at section ${sectionIndex + 1}, item ${itemIndex + 1}`);
    });
  });
  invariant(JSON.stringify(numericLiterals(source)) === JSON.stringify(numericLiterals(translated)), "numeric facts or dates changed");
}

export type VisibleTextUnit = { id: string; text: string };
type OfficeArtifactFormat = "docx" | "pptx" | "xlsx";

const DOCX_VISIBLE_XML = /^word\/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/;
const DOCX_CHART_XML = /^word\/charts\/chart\d+\.xml$/;
const PPTX_VISIBLE_XML = /^ppt\/(?:slides\/slide\d+|notesSlides\/notesSlide\d+)\.xml$/;
const PPTX_CHART_XML = /^ppt\/charts\/chart\d+\.xml$/;
const XLSX_SHARED_XML = /^xl\/sharedStrings\.xml$/;
const XLSX_WORKSHEET_XML = /^xl\/worksheets\/sheet\d+\.xml$/;
const XLSX_DRAWING_XML = /^xl\/drawings\/drawing\d+\.xml$/;
const XLSX_CHART_XML = /^xl\/charts\/chart\d+\.xml$/;

function decodeXml(value: string) {
  return value.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
}

function encodeXml(value: string) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

function decodeHtml(value: string) {
  return decodeXml(value).replace(/&#39;/g, "'").replace(/&nbsp;/g, "\u00a0");
}

function encodeHtml(value: string) {
  return encodeXml(value).replace(/\u00a0/g, "&nbsp;");
}

export function extractHtmlVisibleText(bytes: Uint8Array): VisibleTextUnit[] {
  const html = new TextDecoder().decode(bytes);
  const units: VisibleTextUnit[] = [];
  let excluded = false;
  let index = 0;
  for (const token of html.match(/<[^>]*>|[^<]+/g) ?? []) {
    if (/^<(?:style|script)\b/i.test(token)) excluded = true;
    if (!excluded && !token.startsWith("<") && decodeHtml(token).trim()) units.push({ id: `html#text#${index++}`, text: decodeHtml(token) });
    if (/^<\/(?:style|script)\b/i.test(token)) excluded = false;
  }
  return units;
}

export function translateHtmlArtifact(bytes: Uint8Array, sourceUnits: VisibleTextUnit[], translatedUnits: VisibleTextUnit[]) {
  const translations = new Map(translatedUnits.map((unit) => [unit.id, unit.text]));
  const html = new TextDecoder().decode(bytes);
  let excluded = false;
  let index = 0;
  let replaced = 0;
  const translated = (html.match(/<[^>]*>|[^<]+/g) ?? []).map((token) => {
    if (/^<(?:style|script)\b/i.test(token)) excluded = true;
    let output = token;
    if (!excluded && !token.startsWith("<") && decodeHtml(token).trim()) {
      const replacement = translations.get(`html#text#${index++}`);
      if (replacement !== undefined) {
        output = encodeHtml(replacement);
        replaced++;
      }
    }
    if (/^<\/(?:style|script)\b/i.test(token)) excluded = false;
    return output;
  }).join("");
  invariant(replaced === sourceUnits.length, `only ${replaced} of ${sourceUnits.length} HTML text units were replaced`);
  return new TextEncoder().encode(translated);
}

function textFromNodes(xml: string, tag: "w:t" | "a:t") {
  const pattern = tag === "w:t" ? /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g : /<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g;
  return decodeXml([...xml.matchAll(pattern)].map((match) => match[1]).join(""));
}

function textFromBareNodes(xml: string) {
  return decodeXml([...xml.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((match) => match[1]).join(""));
}

export function extractOfficeVisibleText(bytes: Uint8Array, format: OfficeArtifactFormat): VisibleTextUnit[] {
  const files = unzipSync(bytes);
  const units: VisibleTextUnit[] = [];
  for (const path of Object.keys(files).sort()) {
    if (format === "docx" && DOCX_VISIBLE_XML.test(path)) {
      let index = 0;
      for (const match of strFromU8(files[path]).matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)) {
        const text = textFromNodes(match[0], "w:t");
        if (text.trim()) units.push({ id: `${path}#p#${index}`, text });
        index++;
      }
    }
    if (format === "pptx" && PPTX_VISIBLE_XML.test(path)) {
      let index = 0;
      for (const match of strFromU8(files[path]).matchAll(/<a:p\b[\s\S]*?<\/a:p>/g)) {
        const text = textFromNodes(match[0], "a:t");
        if (text.trim()) units.push({ id: `${path}#p#${index}`, text });
        index++;
      }
    }
    if (format === "pptx" && PPTX_CHART_XML.test(path)) {
      let index = 0;
      for (const match of strFromU8(files[path]).matchAll(/<c:v>([\s\S]*?)<\/c:v>/g)) {
        const text = decodeXml(match[1]);
        if (text.trim() && !/^[+\-]?[\d.,%\s]+$/.test(text)) units.push({ id: `${path}#v#${index}`, text });
        index++;
      }
    }
    if (format === "docx" && DOCX_CHART_XML.test(path)) {
      let index = 0;
      for (const match of strFromU8(files[path]).matchAll(/<c:v>([\s\S]*?)<\/c:v>/g)) {
        const text = decodeXml(match[1]);
        if (text.trim() && !/^[+\-]?[\d.,%\s]+$/.test(text)) units.push({ id: `${path}#v#${index}`, text });
        index++;
      }
    }
    if (format === "xlsx" && XLSX_SHARED_XML.test(path)) {
      let index = 0;
      for (const match of strFromU8(files[path]).matchAll(/<si\b[\s\S]*?<\/si>/g)) {
        const text = textFromBareNodes(match[0]);
        if (text.trim()) units.push({ id: `${path}#si#${index}`, text });
        index++;
      }
    }
    if (format === "xlsx" && XLSX_WORKSHEET_XML.test(path)) {
      let index = 0;
      for (const match of strFromU8(files[path]).matchAll(/<is\b[\s\S]*?<\/is>/g)) {
        const text = textFromBareNodes(match[0]);
        if (text.trim()) units.push({ id: `${path}#is#${index}`, text });
        index++;
      }
    }
    if (format === "xlsx" && XLSX_DRAWING_XML.test(path)) {
      let index = 0;
      for (const match of strFromU8(files[path]).matchAll(/<a:p\b[\s\S]*?<\/a:p>/g)) {
        const text = textFromNodes(match[0], "a:t");
        if (text.trim()) units.push({ id: `${path}#p#${index}`, text });
        index++;
      }
    }
    if (format === "xlsx" && XLSX_CHART_XML.test(path)) {
      let index = 0;
      for (const match of strFromU8(files[path]).matchAll(/<c:v>([\s\S]*?)<\/c:v>/g)) {
        const text = decodeXml(match[1]);
        if (text.trim() && !/^[+\-]?[\d.,%\s]+$/.test(text)) units.push({ id: `${path}#v#${index}`, text });
        index++;
      }
    }
  }
  return units;
}

export const visibleTextTranslationOutput = {
  name: "artifact_visible_text_translation",
  description: "Complete translations of all supplied visible text units.",
  schema: {
    type: "object",
    additionalProperties: false,
    properties: {
      translations: {
        type: "array",
        items: {
          type: "object",
          additionalProperties: false,
          properties: { id: { type: "string" }, text: { type: "string" } },
          required: ["id", "text"],
        },
      },
    },
    required: ["translations"],
  } as Record<string, unknown>,
};

export function buildVisibleTextTranslationPrompt(units: VisibleTextUnit[], language: SupportedLanguage) {
  return `Translate every supplied user-visible text unit into ${language.name} (${language.code}). Return exactly one translation for every id, in the same order. Preserve all facts, numbers, dates, names, qualifiers, and meaning. Do not summarize, add, omit, combine, or split units. Keep source identifiers and trademarks unchanged. Translate UI labels, headings, headers, footers, table labels, captions, chart labels, and notes as normal text.\n\nUnits: ${JSON.stringify(units)}`;
}

export function parseVisibleTextTranslations(raw: string, source: VisibleTextUnit[]) {
  const parsed = JSON.parse(raw) as { translations?: VisibleTextUnit[] };
  invariant(Array.isArray(parsed.translations), "visible text translations are missing");
  invariant(parsed.translations.length === source.length, "visible text unit count changed");
  parsed.translations.forEach((unit, index) => {
    invariant(unit.id === source[index].id, `visible text id changed at position ${index + 1}`);
    invariant(typeof unit.text === "string" && (!source[index].text.trim() || unit.text.trim()), `visible text is empty at position ${index + 1}`);
  });
  return parsed.translations;
}

function replaceNodeText(block: string, tag: "w:t" | "a:t", replacement: string) {
  const pattern = tag === "w:t" ? /<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g : /<a:t(\s[^>]*)?>([\s\S]*?)<\/a:t>/g;
  let first = true;
  return block.replace(pattern, (_match, attributes = "") => {
    const value = first ? encodeXml(replacement) : "";
    first = false;
    return `<${tag}${attributes}>${value}</${tag}>`;
  });
}

function replaceBareNodeText(block: string, replacement: string) {
  let first = true;
  return block.replace(/<t(\s[^>]*)?>([\s\S]*?)<\/t>/g, (_match, attributes = "") => {
    const value = first ? encodeXml(replacement) : "";
    first = false;
    return `<t${attributes}>${value}</t>`;
  });
}

function officeFontSupport(xml: string, format: OfficeArtifactFormat, language: SupportedLanguage) {
  if (format === "docx") {
    const font = encodeXml(language.officeFont);
    const fontNode = `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/>`;
    let result = xml.replace(/<w:rPr(\s[^>]*)?>([\s\S]*?)<\/w:rPr>/g, (_match, attributes = "", body: string) => {
      const withoutFont = body.replace(/<w:rFonts\b[^>]*\/>/g, "");
      const rtlNode = language.direction === "rtl" && !/<w:rtl\b/.test(withoutFont) ? "<w:rtl/>" : "";
      return `<w:rPr${attributes}>${fontNode}${rtlNode}${withoutFont}</w:rPr>`;
    });
    result = result.replace(/<w:r(\s[^>]*)?>((?!<w:rPr)[\s\S]*?)<\/w:r>/g, (_match, attributes = "", body) => `<w:r${attributes}><w:rPr>${fontNode}${language.direction === "rtl" ? "<w:rtl/>" : ""}</w:rPr>${body}</w:r>`);
    if (language.direction === "rtl") {
      result = result.replace(/<w:pPr(\s[^>]*)?>/g, (match) => `${match}<w:bidi/>`);
      result = result.replace(/<w:p(\s[^>]*)?>((?!<w:pPr)[\s\S]*?)<\/w:p>/g, (_match, attributes = "", body) => `<w:p${attributes}><w:pPr><w:bidi/></w:pPr>${body}</w:p>`);
    }
    return result;
  }
  const font = encodeXml(language.officeFont);
  const fontNodes = `<a:latin typeface="${font}"/><a:ea typeface="${font}"/><a:cs typeface="${font}"/>`;
  let result = xml.replace(/<a:(rPr|defRPr)(\s[^>]*)?>([\s\S]*?)<\/a:\1>/g, (_match, tag: string, attributes = "", body: string) => {
    const withoutFonts = body.replace(/<a:(?:latin|ea|cs)\b[^>]*\/>/g, "");
    const nextAttributes = `${attributes}${language.direction === "rtl" && !/\brtl=/.test(attributes) ? ' rtl="1"' : ""}${!/\blang=/.test(attributes) ? ` lang="${language.locale}"` : ""}`;
    return `<a:${tag}${nextAttributes}>${fontNodes}${withoutFonts}</a:${tag}>`;
  });
  result = result.replace(/<a:r(\s[^>]*)?>((?!<a:rPr)[\s\S]*?)<\/a:r>/g, (_match, attributes = "", body) => `<a:r${attributes}><a:rPr lang="${language.locale}"${language.direction === "rtl" ? ' rtl="1"' : ""}>${fontNodes}</a:rPr>${body}</a:r>`);
  if (language.direction === "rtl") {
    result = result.replace(/<a:pPr(\s[^>]*)?>/g, (match, attributes = "") => /\brtl=/.test(attributes) ? match : match.replace(/>$/, ' rtl="1">'));
    result = result.replace(/<a:p(\s[^>]*)?>((?!<a:pPr)[\s\S]*?)<\/a:p>/g, (_match, attributes = "", body) => `<a:p${attributes}><a:pPr rtl="1"/>${body}</a:p>`);
  }
  return result;
}

export function translateOfficeArtifact(input: {
  bytes: Uint8Array;
  format: OfficeArtifactFormat;
  language: SupportedLanguage;
  sourceUnits: VisibleTextUnit[];
  translatedUnits: VisibleTextUnit[];
}) {
  const files = unzipSync(input.bytes);
  const translations = new Map(input.translatedUnits.map((unit) => [unit.id, unit.text]));
  let replaced = 0;
  for (const path of Object.keys(files)) {
    let xml = strFromU8(files[path]);
    if (input.format === "docx" && DOCX_VISIBLE_XML.test(path)) {
      let index = 0;
      xml = xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraph) => {
        const translation = translations.get(`${path}#p#${index++}`);
        if (translation === undefined) return paragraph;
        replaced++;
        return replaceNodeText(paragraph, "w:t", translation);
      });
      xml = officeFontSupport(xml, input.format, input.language);
      files[path] = strToU8(xml);
    }
    if (input.format === "pptx" && PPTX_VISIBLE_XML.test(path)) {
      let index = 0;
      xml = xml.replace(/<a:p\b[\s\S]*?<\/a:p>/g, (paragraph) => {
        const translation = translations.get(`${path}#p#${index++}`);
        if (translation === undefined) return paragraph;
        replaced++;
        return replaceNodeText(paragraph, "a:t", translation);
      });
      xml = officeFontSupport(xml, input.format, input.language);
      files[path] = strToU8(xml);
    }
    if (input.format === "pptx" && PPTX_CHART_XML.test(path)) {
      let index = 0;
      xml = xml.replace(/<c:v>([\s\S]*?)<\/c:v>/g, (node) => {
        const translation = translations.get(`${path}#v#${index++}`);
        if (translation === undefined) return node;
        replaced++;
        return `<c:v>${encodeXml(translation)}</c:v>`;
      });
      files[path] = strToU8(xml);
    }
    if (input.format === "docx" && DOCX_CHART_XML.test(path)) {
      let index = 0;
      xml = xml.replace(/<c:v>([\s\S]*?)<\/c:v>/g, (node) => {
        const translation = translations.get(`${path}#v#${index++}`);
        if (translation === undefined) return node;
        replaced++;
        return `<c:v>${encodeXml(translation)}</c:v>`;
      });
      files[path] = strToU8(xml);
    }
    if (input.format === "xlsx" && XLSX_SHARED_XML.test(path)) {
      let index = 0;
      xml = xml.replace(/<si\b[\s\S]*?<\/si>/g, (item) => {
        const translation = translations.get(`${path}#si#${index++}`);
        if (translation === undefined) return item;
        replaced++;
        return replaceBareNodeText(item, translation);
      });
      files[path] = strToU8(xml);
    }
    if (input.format === "xlsx" && XLSX_WORKSHEET_XML.test(path)) {
      let index = 0;
      xml = xml.replace(/<is\b[\s\S]*?<\/is>/g, (item) => {
        const translation = translations.get(`${path}#is#${index++}`);
        if (translation === undefined) return item;
        replaced++;
        return replaceBareNodeText(item, translation);
      });
      files[path] = strToU8(xml);
    }
    if (input.format === "xlsx" && XLSX_DRAWING_XML.test(path)) {
      let index = 0;
      xml = xml.replace(/<a:p\b[\s\S]*?<\/a:p>/g, (paragraph) => {
        const translation = translations.get(`${path}#p#${index++}`);
        if (translation === undefined) return paragraph;
        replaced++;
        return replaceNodeText(paragraph, "a:t", translation);
      });
      xml = officeFontSupport(xml, input.format, input.language);
      files[path] = strToU8(xml);
    }
    if (input.format === "xlsx" && XLSX_CHART_XML.test(path)) {
      let index = 0;
      xml = xml.replace(/<c:v>([\s\S]*?)<\/c:v>/g, (node) => {
        const translation = translations.get(`${path}#v#${index++}`);
        if (translation === undefined) return node;
        replaced++;
        return `<c:v>${encodeXml(translation)}</c:v>`;
      });
      files[path] = strToU8(xml);
    }
    if (input.format === "xlsx" && path === "xl/styles.xml") {
      const font = encodeXml(input.language.officeFont);
      xml = xml.replace(/<name\s+val="[^"]*"\s*\/>/g, `<name val="${font}"/>`);
      files[path] = strToU8(xml);
    }
  }
  invariant(replaced === input.sourceUnits.length, `only ${replaced} of ${input.sourceUnits.length} visible text units were replaced`);
  return zipSync(files, { level: 6 });
}

export function translatedArtifactFileName(filename: string, languageCode: string, version: number) {
  const dot = filename.lastIndexOf(".");
  const extension = dot >= 0 ? filename.slice(dot) : "";
  const base = (dot >= 0 ? filename.slice(0, dot) : filename).replace(/_v\d+$/i, "");
  return `${base}_${languageCode.replace(/[^a-z0-9]+/gi, "-")}_v${version}${extension}`;
}

export function translatedArtifactLineage(source: { artifactId: string; templateSourceId?: string | null }, contentLanguage: string) {
  return {
    parentArtifactId: source.artifactId,
    templateSourceId: source.templateSourceId ?? null,
    operation: "translate" as const,
    contentLanguage,
  };
}
