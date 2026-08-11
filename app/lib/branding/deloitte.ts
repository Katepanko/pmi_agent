import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { DELOITTE_LOGO_PNG_BASE64, DELOITTE_MASTER_PPTX_BASE64 } from "./deloitte-assets.ts";

export const DeloitteBrand = {
  name: "Deloitte",
  colors: {
    black: "000000",
    white: "FFFFFF",
    brightGreen: "86BC25",
    green: "26890D",
    deepGreen: "046A38",
    darkGreen: "1C3D26",
    paleGreen: "F1F6E4",
    lightGray: "E6E6E6",
    coolGray: "75787B",
    red: "DA291C",
    amber: "ED8B00",
  },
  typography: {
    heading: "Aptos",
    body: "Aptos",
  },
  powerPoint: {
    templateName: "Deloitte Master.pptx",
    // The attached deck's first slide is linked to this exact title layout.
    titleLayout: "slideLayout4.xml",
    contentLayout: "slideLayout28.xml",
  },
  footer: {
    confidentiality: "Confidential",
    copyright: (year = new Date().getFullYear()) => `Deloitte ${year}`,
  },
  document: { logoLocation: "header" },
  excel: { logoLocation: "first-sheet and worksheet header" },
  html: { logoLocation: "persistent report header" },
} as const;

export const DELOITTE_LOGO_DATA_URI = `data:image/png;base64,${DELOITTE_LOGO_PNG_BASE64}`;

function decodeBase64(value: string) {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

export function deloitteLogoBytes() {
  return decodeBase64(DELOITTE_LOGO_PNG_BASE64);
}

function masterTemplateBytes() {
  return decodeBase64(DELOITTE_MASTER_PPTX_BASE64);
}

function xmlEscape(value: string) {
  return value.replace(/[&<>'"]/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&apos;", '"': "&quot;",
  })[character] ?? character);
}

function decoded(value: Uint8Array | undefined, path: string) {
  if (!value) throw new Error(`The Deloitte master package is missing ${path}.`);
  return strFromU8(value);
}

function replaceShapeText(xml: string, marker: string, paragraphXml: string) {
  let found = false;
  const result = xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (shape) => {
    if (!shape.includes(marker)) return shape;
    found = true;
    return shape.replace(/(<p:txBody>[\s\S]*?<a:lstStyle\/?>(?:<\/a:lstStyle>)?)[\s\S]*?(<\/p:txBody>)/, `$1${paragraphXml}$2`);
  });
  if (!found) throw new Error(`The Deloitte title template no longer contains ${marker}.`);
  return result;
}

function paragraph(value: string, size: number, options: { bold?: boolean; color?: string; align?: "l" | "r" } = {}) {
  const lines = value.split("\n");
  const runs = lines.map((line, index) => `${index ? "<a:br/>" : ""}<a:r><a:rPr lang="en-US" sz="${size}"${options.bold ? ' b="1"' : ""}>${options.color ? `<a:solidFill><a:srgbClr val="${options.color}"/></a:solidFill>` : ""}<a:latin typeface="Aptos"/></a:rPr><a:t>${xmlEscape(line)}</a:t></a:r>`).join("");
  return `<a:p><a:pPr${options.align ? ` algn="${options.align}"` : ""}/>${runs}<a:endParaRPr lang="en-US" sz="${size}"/></a:p>`;
}

function coverTitle(value: string) {
  const words = value.trim().split(/\s+/);
  if (value.length <= 38) return value;
  let first = "";
  let second = "";
  for (const word of words) {
    if (!second && `${first} ${word}`.trim().length <= Math.ceil(value.length / 2) + 4) first = `${first} ${word}`.trim();
    else second = `${second} ${word}`.trim();
  }
  return second ? `${first}\n${second}` : value;
}

function coverTitleSize(value: string) {
  return value.length > 82 ? 2300 : value.length > 55 ? 2700 : 3200;
}

function insertBeforeClosing(xml: string, closingTag: string, addition: string) {
  const index = xml.lastIndexOf(closingTag);
  if (index < 0) throw new Error(`Cannot update malformed Office XML: ${closingTag} not found.`);
  return `${xml.slice(0, index)}${addition}${xml.slice(index)}`;
}

function addContentType(xml: string, partName: string, contentType: string) {
  if (xml.includes(`PartName="${partName}"`)) return xml;
  return insertBeforeClosing(xml, "</Types>", `<Override PartName="${partName}" ContentType="${contentType}"/>`);
}

const SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const NOTES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
const NOTES_MASTER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";

export function applyDeloittePowerPointTemplate(input: {
  generated: Uint8Array;
  title: string;
  subtitle?: string;
  projectName?: string;
  location?: string;
  date?: string;
  audience: string;
  slideCount: number;
}) {
  const template = unzipSync(masterTemplateBytes());
  const generated = unzipSync(input.generated);
  const output: Record<string, Uint8Array> = { ...template };
  const logoPath = "ppt/media/deloitte-logo.png";
  output[logoPath] = deloitteLogoBytes();

  let coverXml = decoded(template["ppt/slides/slide1.xml"], "ppt/slides/slide1.xml");
  const fittedTitle = coverTitle(input.title);
  coverXml = replaceShapeText(coverXml, 'type="ctrTitle"', paragraph(fittedTitle, coverTitleSize(input.title), { color: DeloitteBrand.colors.deepGreen }));
  const metadata = [input.projectName, input.subtitle, input.audience, input.location, input.date].filter(Boolean).join("  |  ");
  const metadataSize = metadata.length > 95 ? 800 : metadata.length > 65 ? 1000 : metadata.length > 45 ? 1200 : 1400;
  coverXml = replaceShapeText(coverXml, 'type="body" sz="quarter" idx="10"', paragraph(metadata || input.audience, metadataSize, { bold: true, color: DeloitteBrand.colors.black }));
  coverXml = coverXml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (shape) => shape.includes('<p:ph type="pic"') ? "" : shape);
  output["ppt/slides/slide1.xml"] = strToU8(coverXml);

  let masterXml = decoded(template["ppt/slideMasters/slideMaster1.xml"], "ppt/slideMasters/slideMaster1.xml");
  masterXml = replaceShapeText(masterXml, 'name="CaseCode"', paragraph(`${input.title}\n${DeloitteBrand.footer.confidentiality}`, 800, { color: DeloitteBrand.colors.coolGray, align: "r" }));
  masterXml = replaceShapeText(masterXml, 'name="Copyright"', paragraph(DeloitteBrand.footer.copyright(), 800, { color: DeloitteBrand.colors.coolGray }));
  const logoPicture = `<p:pic><p:nvPicPr><p:cNvPr id="99" name="Deloitte logo" descr="Approved Deloitte logo"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdBrandLogo"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="10340000" y="115000"/><a:ext cx="1280000" cy="240000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  masterXml = insertBeforeClosing(masterXml, "</p:spTree>", logoPicture);
  output["ppt/slideMasters/slideMaster1.xml"] = strToU8(masterXml);

  let masterRels = decoded(template["ppt/slideMasters/_rels/slideMaster1.xml.rels"], "ppt/slideMasters/_rels/slideMaster1.xml.rels");
  masterRels = insertBeforeClosing(masterRels, "</Relationships>", `<Relationship Id="rIdBrandLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/deloitte-logo.png"/>`);
  output["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = strToU8(masterRels);

  const firstGeneratedRels = decoded(generated["ppt/slides/_rels/slide1.xml.rels"], "generated slide1 relationships");
  const notesRelationship = firstGeneratedRels.match(/<Relationship[^>]+Type="[^"]+\/notesSlide"[^>]*\/>/)?.[0];
  let coverRels = decoded(template["ppt/slides/_rels/slide1.xml.rels"], "ppt/slides/_rels/slide1.xml.rels");
  if (notesRelationship) coverRels = insertBeforeClosing(coverRels, "</Relationships>", notesRelationship);
  output["ppt/slides/_rels/slide1.xml.rels"] = strToU8(coverRels);

  for (let index = 2; index <= input.slideCount; index++) {
    const slidePath = `ppt/slides/slide${index}.xml`;
    const relsPath = `ppt/slides/_rels/slide${index}.xml.rels`;
    output[slidePath] = generated[slidePath];
    let rels = decoded(generated[relsPath], relsPath);
    rels = rels.replace(/Target="\.\.\/slideLayouts\/slideLayout\d+\.xml"/, `Target="../slideLayouts/${DeloitteBrand.powerPoint.contentLayout}"`);
    output[relsPath] = strToU8(rels);
  }

  for (const [path, bytes] of Object.entries(generated)) {
    if (path.startsWith("ppt/notesSlides/") || path.startsWith("ppt/notesMasters/")) output[path] = bytes;
  }

  let presentationXml = decoded(template["ppt/presentation.xml"], "ppt/presentation.xml");
  const slideIds = [`<p:sldId id="259" r:id="rId2"/>`, ...Array.from({ length: Math.max(0, input.slideCount - 1) }, (_, offset) => `<p:sldId id="${260 + offset}" r:id="rId${101 + offset}"/>`)].join("");
  presentationXml = presentationXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${slideIds}</p:sldIdLst>`);
  presentationXml = presentationXml.replace(/<p:notesMasterIdLst>[\s\S]*?<\/p:notesMasterIdLst>/, "");
  presentationXml = presentationXml.replace("<p:sldSz", `<p:notesMasterIdLst><p:notesMasterId r:id="rId100"/></p:notesMasterIdLst><p:sldSz`);
  output["ppt/presentation.xml"] = strToU8(presentationXml);

  let presentationRels = decoded(template["ppt/_rels/presentation.xml.rels"], "ppt/_rels/presentation.xml.rels");
  const newRelationships = [`<Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>`, ...Array.from({ length: Math.max(0, input.slideCount - 1) }, (_, offset) => `<Relationship Id="rId${101 + offset}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${offset + 2}.xml"/>`)].join("");
  presentationRels = insertBeforeClosing(presentationRels, "</Relationships>", newRelationships);
  output["ppt/_rels/presentation.xml.rels"] = strToU8(presentationRels);

  let contentTypes = decoded(template["[Content_Types].xml"], "[Content_Types].xml");
  contentTypes = addContentType(contentTypes, "/ppt/notesMasters/notesMaster1.xml", NOTES_MASTER_CONTENT_TYPE);
  for (let index = 1; index <= input.slideCount; index++) {
    if (index > 1) contentTypes = addContentType(contentTypes, `/ppt/slides/slide${index}.xml`, SLIDE_CONTENT_TYPE);
    contentTypes = addContentType(contentTypes, `/ppt/notesSlides/notesSlide${index}.xml`, NOTES_CONTENT_TYPE);
  }
  output["[Content_Types].xml"] = strToU8(contentTypes);

  return zipSync(output, { level: 6 });
}

export function validateDeloittePowerPoint(bytes: Uint8Array, input: { title: string; slideCount: number; metadata?: string[] }) {
  const files = unzipSync(bytes);
  const slidePaths = Object.keys(files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path));
  if (slidePaths.length !== input.slideCount) throw new Error("Deloitte PowerPoint validation failed: slide count mismatch.");
  const cover = decoded(files["ppt/slides/slide1.xml"], "ppt/slides/slide1.xml");
  const coverText = cover.replace(/<[^>]+>/g, " ").replace(/&amp;/g, "&").replace(/\s+/g, " ");
  if (!coverText.includes(input.title.replace(/\s+/g, " "))) throw new Error("Deloitte PowerPoint validation failed: requested title is missing from slide 1.");
  for (const value of input.metadata?.filter(Boolean) ?? []) {
    if (!coverText.includes(value.replace(/\s+/g, " "))) throw new Error(`Deloitte PowerPoint validation failed: cover metadata is missing (${value}).`);
  }
  if (/Click to add|Click icon|Presentation title|To edit/i.test(cover)) throw new Error("Deloitte PowerPoint validation failed: placeholder text remains on slide 1.");
  const coverRels = decoded(files["ppt/slides/_rels/slide1.xml.rels"], "slide 1 relationships");
  if (!coverRels.includes(`Target="../slideLayouts/${DeloitteBrand.powerPoint.titleLayout}"`)) throw new Error("Deloitte PowerPoint validation failed: slide 1 does not use the approved title layout.");
  if ((coverRels.match(/deloitte-logo\.png/g) ?? []).length) throw new Error("Deloitte PowerPoint validation failed: the cover logo is duplicated.");
  const master = decoded(files["ppt/slideMasters/slideMaster1.xml"], "Deloitte slide master");
  const masterRels = decoded(files["ppt/slideMasters/_rels/slideMaster1.xml.rels"], "Deloitte slide master relationships");
  if (!master.includes("Deloitte logo") || !masterRels.includes("deloitte-logo.png") || !files["ppt/media/deloitte-logo.png"]) throw new Error("Deloitte PowerPoint validation failed: the recurring logo is missing.");
  for (let index = 2; index <= input.slideCount; index++) {
    const rels = decoded(files[`ppt/slides/_rels/slide${index}.xml.rels`], `slide ${index} relationships`);
    if (!rels.includes(`Target="../slideLayouts/${DeloitteBrand.powerPoint.contentLayout}"`)) throw new Error(`Deloitte PowerPoint validation failed: slide ${index} is not linked to the Deloitte content layout.`);
  }
  return true;
}
