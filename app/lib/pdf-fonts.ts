import fontkit from "@pdf-lib/fontkit";
import { PDFDocument, PDFFont, StandardFonts } from "pdf-lib";

type FontModule = { default: string };

function decodeDataUrl(value: string) {
  const payload = value.slice(value.indexOf(",") + 1);
  if (value.includes(";base64,")) return Uint8Array.from(atob(payload), (character) => character.charCodeAt(0));
  return new TextEncoder().encode(decodeURIComponent(payload));
}

async function nodeFont(relativePath: string) {
  const { readFile } = await import("node:fs/promises");
  return new Uint8Array(await readFile(new URL(`../../node_modules/${relativePath}`, import.meta.url)));
}

async function fontBytes(languageCode: string) {
  if (languageCode === "zh-CN") {
    try { return decodeDataUrl((await import("@fontsource/noto-sans-sc/files/noto-sans-sc-chinese-simplified-400-normal.woff?inline") as FontModule).default); }
    catch { return nodeFont("@fontsource/noto-sans-sc/files/noto-sans-sc-chinese-simplified-400-normal.woff"); }
  }
  if (languageCode === "ja") {
    try { return decodeDataUrl((await import("@fontsource/noto-sans-jp/files/noto-sans-jp-japanese-400-normal.woff?inline") as FontModule).default); }
    catch { return nodeFont("@fontsource/noto-sans-jp/files/noto-sans-jp-japanese-400-normal.woff"); }
  }
  if (languageCode === "ko") {
    try { return decodeDataUrl((await import("@fontsource/noto-sans-kr/files/noto-sans-kr-korean-400-normal.woff?inline") as FontModule).default); }
    catch { return nodeFont("@fontsource/noto-sans-kr/files/noto-sans-kr-korean-400-normal.woff"); }
  }
  if (languageCode === "ar") {
    try { return decodeDataUrl((await import("@fontsource/noto-sans-arabic/files/noto-sans-arabic-arabic-400-normal.woff?inline") as FontModule).default); }
    catch { return nodeFont("@fontsource/noto-sans-arabic/files/noto-sans-arabic-arabic-400-normal.woff"); }
  }
  if (languageCode === "ru" || languageCode === "uk") {
    try { return decodeDataUrl((await import("@fontsource/noto-sans/files/noto-sans-cyrillic-400-normal.woff?inline") as FontModule).default); }
    catch { return nodeFont("@fontsource/noto-sans/files/noto-sans-cyrillic-400-normal.woff"); }
  }
  if (languageCode === "el") {
    try { return decodeDataUrl((await import("@fontsource/noto-sans/files/noto-sans-greek-400-normal.woff?inline") as FontModule).default); }
    catch { return nodeFont("@fontsource/noto-sans/files/noto-sans-greek-400-normal.woff"); }
  }
  try { return decodeDataUrl((await import("@fontsource/noto-sans/files/noto-sans-latin-ext-400-normal.woff?inline") as FontModule).default); }
  catch { return nodeFont("@fontsource/noto-sans/files/noto-sans-latin-ext-400-normal.woff"); }
}

export async function embedPdfFonts(document: PDFDocument, languageCode?: string): Promise<{ regular: PDFFont; bold: PDFFont }> {
  if (!languageCode || languageCode === "en") {
    return {
      regular: await document.embedFont(StandardFonts.Helvetica),
      bold: await document.embedFont(StandardFonts.HelveticaBold),
    };
  }
  document.registerFontkit(fontkit);
  const bytes = await fontBytes(languageCode);
  const regular = await document.embedFont(bytes, { subset: true });
  return { regular, bold: regular };
}
