export type CommunicationLanguage = "en" | "de";

export type SupportedLanguage = {
  code: string;
  name: string;
  locale: string;
  direction: "ltr" | "rtl";
  officeFont: string;
  aliases: string[];
};

const LANGUAGE_DEFINITIONS: SupportedLanguage[] = [
  { code: "en", name: "English", locale: "en-US", direction: "ltr", officeFont: "Aptos", aliases: ["english", "englisch", "englische", "englischen", "englischer", "englisches"] },
  { code: "de", name: "German", locale: "de-DE", direction: "ltr", officeFont: "Aptos", aliases: ["german", "deutsch", "deutsche", "deutschen", "deutscher", "deutsches"] },
  { code: "fr", name: "French", locale: "fr-FR", direction: "ltr", officeFont: "Aptos", aliases: ["french", "français", "francais", "französisch", "franzoesisch"] },
  { code: "es", name: "Spanish", locale: "es-ES", direction: "ltr", officeFont: "Aptos", aliases: ["spanish", "español", "espanol", "spanisch"] },
  { code: "pt", name: "Portuguese", locale: "pt-PT", direction: "ltr", officeFont: "Aptos", aliases: ["portuguese", "português", "portugues", "portugiesisch"] },
  { code: "it", name: "Italian", locale: "it-IT", direction: "ltr", officeFont: "Aptos", aliases: ["italian", "italiano", "italienisch"] },
  { code: "zh-CN", name: "Simplified Chinese", locale: "zh-CN", direction: "ltr", officeFont: "Microsoft YaHei", aliases: ["simplified chinese", "chinese", "中文", "简体中文", "chinesisch", "vereinfachte chinesisch", "vereinfachtes chinesisch", "vereinfachten chinesisch", "vereinfachtem chinesisch"] },
  { code: "ja", name: "Japanese", locale: "ja-JP", direction: "ltr", officeFont: "Yu Gothic", aliases: ["japanese", "日本語", "japanisch"] },
  { code: "ko", name: "Korean", locale: "ko-KR", direction: "ltr", officeFont: "Malgun Gothic", aliases: ["korean", "한국어", "koreanisch"] },
  { code: "nl", name: "Dutch", locale: "nl-NL", direction: "ltr", officeFont: "Aptos", aliases: ["dutch", "nederlands", "niederländisch", "niederlaendisch"] },
  { code: "pl", name: "Polish", locale: "pl-PL", direction: "ltr", officeFont: "Aptos", aliases: ["polish", "polski", "polnisch"] },
  { code: "tr", name: "Turkish", locale: "tr-TR", direction: "ltr", officeFont: "Aptos", aliases: ["turkish", "türkçe", "turkce", "türkisch", "tuerkisch"] },
  { code: "ar", name: "Arabic", locale: "ar-SA", direction: "rtl", officeFont: "Arial", aliases: ["arabic", "العربية", "arabisch"] },
  { code: "ru", name: "Russian", locale: "ru-RU", direction: "ltr", officeFont: "Aptos", aliases: ["russian", "русский", "russisch"] },
  { code: "cs", name: "Czech", locale: "cs-CZ", direction: "ltr", officeFont: "Aptos", aliases: ["czech", "čeština", "cestina", "tschechisch"] },
  { code: "ro", name: "Romanian", locale: "ro-RO", direction: "ltr", officeFont: "Aptos", aliases: ["romanian", "română", "romana", "rumänisch", "rumaenisch"] },
  { code: "hu", name: "Hungarian", locale: "hu-HU", direction: "ltr", officeFont: "Aptos", aliases: ["hungarian", "magyar", "ungarisch"] },
  { code: "da", name: "Danish", locale: "da-DK", direction: "ltr", officeFont: "Aptos", aliases: ["danish", "dansk", "dänisch", "daenisch"] },
  { code: "sv", name: "Swedish", locale: "sv-SE", direction: "ltr", officeFont: "Aptos", aliases: ["swedish", "svenska", "schwedisch"] },
  { code: "no", name: "Norwegian", locale: "nb-NO", direction: "ltr", officeFont: "Aptos", aliases: ["norwegian", "norsk", "norwegisch"] },
  { code: "fi", name: "Finnish", locale: "fi-FI", direction: "ltr", officeFont: "Aptos", aliases: ["finnish", "suomi", "finnisch"] },
  { code: "el", name: "Greek", locale: "el-GR", direction: "ltr", officeFont: "Aptos", aliases: ["greek", "ελληνικά", "griechisch"] },
  { code: "uk", name: "Ukrainian", locale: "uk-UA", direction: "ltr", officeFont: "Aptos", aliases: ["ukrainian", "українська", "ukrainisch"] },
];

export const SUPPORTED_TRANSLATION_LANGUAGES = Object.freeze(LANGUAGE_DEFINITIONS);

function escaped(value: string) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function normalized(value: string) {
  return value.normalize("NFKC").toLocaleLowerCase();
}

function languageMentions(message: string) {
  const value = normalized(message);
  const mentions: Array<{ language: SupportedLanguage; index: number; length: number }> = [];
  for (const language of LANGUAGE_DEFINITIONS) {
    for (const alias of [...language.aliases].sort((a, b) => b.length - a.length)) {
      const pattern = new RegExp(`(?:^|[^\\p{L}])(${escaped(alias)})(?=$|[^\\p{L}])`, "giu");
      for (const match of value.matchAll(pattern)) {
        const offset = match[0].length - match[1].length;
        mentions.push({ language, index: (match.index ?? 0) + offset, length: match[1].length });
      }
    }
  }
  return mentions.sort((a, b) => a.index - b.index || a.length - b.length);
}

export function supportedLanguage(code: string | null | undefined) {
  if (!code) return null;
  return LANGUAGE_DEFINITIONS.find((language) => language.code.toLocaleLowerCase() === code.toLocaleLowerCase()) ?? null;
}

export function detectRequestedOutputLanguage(message: string): SupportedLanguage | null {
  const value = normalized(message);
  const mentions = languageMentions(value);
  if (!mentions.length) return null;
  const translationAction = /\b(?:translate|translated|translation|übersetz\w*|uebersetz\w*)\b/iu.test(value);
  const outputNoun = /^(?:\s+|[-–—,:])*?(?:version|language|report|document|presentation|deck|slides?|file|workbook|dashboard|fassung|sprache|bericht|dokument|präsentation|praesentation|folien?|datei|arbeitsmappe)/iu;
  const outputCue = /(?:\b(?:in|into|to|as|auf|ins|nach|als)\s+(?:der|die|das|einer|einem|einen)?\s*)$/iu;
  for (const mention of [...mentions].reverse()) {
    const before = value.slice(Math.max(0, mention.index - 45), mention.index);
    const after = value.slice(mention.index + mention.length, mention.index + mention.length + 35);
    if (translationAction || outputCue.test(before) || outputNoun.test(after)) return mention.language;
  }
  return null;
}

const GERMAN_SIGNALS = /\b(?:ich|mir|mich|bitte|erstelle|erstellen|generiere|generieren|nutze|verwende|speichere|übersetze|uebersetze|bericht|datei|vorlage|gesamten|gleich(?:e|en|er|es)?|letzte[nsr]?|aktuellen|fortschritt|kannst|möchte|moechte|auf|als|und|den|die|das|einen|eine|zum|zur)\b/giu;
const ENGLISH_SIGNALS = /\b(?:i|me|please|create|generate|use|save|translate|report|file|template|entire|same|last|latest|current|progress|can|want|in|as|and|the|a|to|for)\b/giu;

export function detectCommunicationLanguage(message: string): CommunicationLanguage {
  const value = normalized(message);
  const explicit = detectRequestedCommunicationLanguage(value);
  if (explicit) return explicit;
  const germanScore = (value.match(GERMAN_SIGNALS) ?? []).length * 2 + (value.match(/[äöüß]/g) ?? []).length;
  const englishScore = (value.match(ENGLISH_SIGNALS) ?? []).length * 2;
  return germanScore > englishScore ? "de" : "en";
}

export function detectRequestedCommunicationLanguage(message: string): CommunicationLanguage | null {
  const value = normalized(message);
  const asksForReply = /\b(?:respond|reply|answer|speak|communicate|antworte|antworten|sprich|sprechen|kommuniziere|kommunizieren)\b/iu.test(value);
  if (!asksForReply) return null;
  const requested = detectRequestedOutputLanguage(value);
  return requested?.code === "de" ? "de" : requested?.code === "en" ? "en" : null;
}

export function referencesPriorArtifact(message: string) {
  const value = normalized(message);
  return /\b(?:same\s+(?:report|file|document|presentation|deck|workbook)|the\s+(?:entire|whole)\s+(?:report|file|document|presentation|deck)|(?:last|latest|previous)(?:\s+generated)?\s+(?:report|file|document|presentation|deck|artifact)|this\s+(?:report|file|document|presentation|deck|artifact)|(?:save|export|convert|regenerate|use|update|revise|change)\s+(?:it|this|that)|gleich(?:e|en|er|es)?\s+(?:bericht|datei|dokument|präsentation|praesentation|folien|arbeitsmappe)|den\s+(?:gesamten|ganzen)\s+(?:bericht|report)|die\s+(?:gesamte|ganze)\s+(?:präsentation|praesentation|datei)|letzt(?:e|en|er|es)\s+(?:generiert(?:e|en|er|es)\s+)?(?:bericht|datei|dokument|präsentation|praesentation|artefakt)|dies(?:e|en|er|es)\s+(?:bericht|datei|dokument|präsentation|praesentation|artefakt)|(?:speichere|exportiere|konvertiere|generiere|nutze|verwende|aktualisiere|überarbeite|ueberarbeite|ändere|aendere)\s+(?:es|das|dies|diesen|diese|den\s+bericht|die\s+präsentation|die\s+praesentation))\b/iu.test(value);
}

export function inferArtifactLanguage(model: unknown): SupportedLanguage {
  const strings: string[] = [];
  const visit = (value: unknown, key = "") => {
    if (typeof value === "string" && !["sourceRefs", "status", "evidenceType", "layout", "type", "templateSlide"].includes(key)) strings.push(value);
    else if (Array.isArray(value)) value.forEach((entry) => visit(entry, key));
    else if (value && typeof value === "object") Object.entries(value).forEach(([childKey, child]) => visit(child, childKey));
  };
  visit(model);
  return supportedLanguage(detectCommunicationLanguage(strings.join(" "))) ?? LANGUAGE_DEFINITIONS[0];
}

export type ArtifactLanguageContext = {
  communicationLanguage: CommunicationLanguage;
  requestedOutputLanguage: SupportedLanguage | null;
  sourceArtifactLanguage: SupportedLanguage | null;
  outputLanguage: SupportedLanguage;
};

export function resolveArtifactLanguageContext(input: {
  message: string;
  sourceArtifactLanguageCode?: string | null;
  sourceArtifactModel?: unknown;
  modifiesExistingArtifact: boolean;
}): ArtifactLanguageContext {
  const communicationLanguage = detectCommunicationLanguage(input.message);
  const requestedOutputLanguage = detectRequestedOutputLanguage(input.message);
  const sourceArtifactLanguage = supportedLanguage(input.sourceArtifactLanguageCode)
    ?? (input.sourceArtifactModel ? inferArtifactLanguage(input.sourceArtifactModel) : null);
  const outputLanguage = requestedOutputLanguage
    ?? (input.modifiesExistingArtifact ? sourceArtifactLanguage : null)
    ?? supportedLanguage(communicationLanguage)
    ?? LANGUAGE_DEFINITIONS[0];
  return { communicationLanguage, requestedOutputLanguage, sourceArtifactLanguage, outputLanguage };
}

export function languageName(language: SupportedLanguage, communicationLanguage: CommunicationLanguage) {
  if (communicationLanguage === "de") {
    const names: Record<string, string> = {
      en: "Englisch", de: "Deutsch", fr: "Französisch", es: "Spanisch", pt: "Portugiesisch", it: "Italienisch",
      "zh-CN": "vereinfachtes Chinesisch", ja: "Japanisch", ko: "Koreanisch", nl: "Niederländisch", pl: "Polnisch",
      tr: "Türkisch", ar: "Arabisch", ru: "Russisch", cs: "Tschechisch", ro: "Rumänisch", hu: "Ungarisch",
      da: "Dänisch", sv: "Schwedisch", no: "Norwegisch", fi: "Finnisch", el: "Griechisch", uk: "Ukrainisch",
    };
    return names[language.code] ?? language.name;
  }
  return language.name;
}
