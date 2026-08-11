import assert from "node:assert/strict";
import test from "node:test";
import { strFromU8, unzipSync } from "fflate";
import {
  isPresentationRequest,
  parsePresentationModel,
  renderPresentation,
} from "../app/lib/presentation.ts";
import { extractOpenAIResponseText } from "../app/lib/llm/openai-response.ts";

function shapeContaining(slideXml, text) {
  return [...slideXml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)]
    .map((match) => match[0])
    .find((shape) => shape.includes(text));
}

test("extracts text from the raw Responses API output array", () => {
  const text = extractOpenAIResponseText({
    id: "resp_test",
    object: "response",
    status: "completed",
    output: [{
      type: "message",
      content: [
        { type: "output_text", text: "{\"title\":\"SteerCo report\"," },
        { type: "output_text", text: "\"slides\":[]}" },
      ],
    }],
  });
  assert.equal(text, "{\"title\":\"SteerCo report\",\"slides\":[]}");
});

test("recognizes explicit presentation requests and stored-deck revisions", () => {
  for (const prompt of [
    "Create a PowerPoint",
    "Create a presentation",
    "Create a SteerCo deck",
    "Generate PPT",
    "Generate PPTX",
    "Prepare 5 slides for the CFO",
    "Turn this into a management presentation",
  ]) assert.equal(isPresentationRequest(prompt), true, prompt);

  assert.equal(isPresentationRequest("Summarize the current integration status."), false);
  assert.equal(isPresentationRequest(
    "Change the decisions slide so it only contains decisions that require Steering Committee approval.",
    true,
  ), true);
});

test("parses a structured model and renders an openable, editable PPTX package", async () => {
  const model = parsePresentationModel(JSON.stringify({
    title: "SteerCo Integration Report",
    audience: "Steering Committee",
    executiveSummary: "Integration momentum remains intact, but one dependency requires intervention.",
    slides: [
      {
        title: "Integration momentum remains intact, but one dependency requires intervention",
        layout: "cover",
        items: [],
      },
      {
        title: "IT cutover now defines the integrated critical path",
        kicker: "Integration trajectory",
        keyMessage: "Finance testing depends on an agreed migration sequence.",
        layout: "trajectory",
        items: [{
          label: "IT cutover",
          detail: "Migration sequence remains open.",
          implication: "Finance testing may slip.",
          status: "red",
          evidenceType: "fact",
          sourceRefs: ["risk-register"],
        }],
        sourceNotes: ["Risk Register, current extract"],
      },
      {
        title: "Task-level evidence shows 50.7% mean progress, well below both reported roll-ups and the delivery trajectory assumed in the current plan",
        kicker: "Progress reconciliation",
        keyMessage: "Report task and workstream detail while the roll-up basis is unresolved.",
        layout: "cards",
        items: [{
          label: "Roll-up basis",
          detail: "The calculation basis requires reconciliation.",
          status: "amber",
          evidenceType: "inference",
        }],
      },
    ],
  }), "Steering Committee");

  const bytes = await renderPresentation(model);
  assert.ok(bytes.byteLength > 50_000);
  const files = unzipSync(bytes);
  assert.ok(files["[Content_Types].xml"]);
  assert.ok(files["ppt/presentation.xml"]);
  assert.ok(files["ppt/slideMasters/slideMaster1.xml"]);
  assert.ok(files["ppt/media/deloitte-logo.png"]);
  assert.ok(files["ppt/slides/slide1.xml"]);
  assert.ok(files["ppt/slides/slide2.xml"]);
  assert.ok(files["ppt/slides/slide3.xml"]);
  assert.equal(Object.keys(files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length, 3);
  const slide2Xml = strFromU8(files["ppt/slides/slide2.xml"]);
  const slide3Xml = strFromU8(files["ppt/slides/slide3.xml"]);
  const masterXml = strFromU8(files["ppt/slideMasters/slideMaster1.xml"]);
  assert.match(slide2Xml, /IT cutover now defines/);
  assert.match(strFromU8(files["ppt/slides/slide1.xml"]), /SteerCo Integration Report/);
  assert.match(strFromU8(files["ppt/slides\/_rels/slide1.xml.rels"]), /slideLayout4\.xml/);
  assert.match(strFromU8(files["ppt/slides\/_rels/slide2.xml.rels"]), /slideLayout28\.xml/);
  assert.match(masterXml, /Deloitte logo/);
  assert.equal((masterXml.match(/type="slidenum"/gi) ?? []).length, 1);
  assert.doesNotMatch(slide2Xml, /type="(?:sldNum|slidenum)"/i);
  assert.doesNotMatch(slide2Xml, /<a:ext cx="(?:146304|164592)" cy="6858000"\/>/);

  const titleShape = shapeContaining(slide2Xml, "IT cutover now defines");
  const subtitleShape = shapeContaining(slide2Xml, "Finance testing depends");
  assert.ok(titleShape);
  assert.ok(subtitleShape);
  assert.match(titleShape, /<a:off x="502920" y="530352"\/>/);
  assert.match(titleShape, /sz="3000"/);
  assert.match(subtitleShape, /<a:off x="502920" y="1115568"\/>/);

  const longTitleShape = shapeContaining(slide3Xml, "Task-level evidence shows");
  const longSubtitleShape = shapeContaining(slide3Xml, "Report task and workstream detail");
  assert.ok(longTitleShape);
  assert.ok(longSubtitleShape);
  assert.match(longTitleShape, /<a:off x="502920" y="530352"\/>/);
  assert.match(longTitleShape, /<a:ext cx="11018520" cy="804672"\/>/);
  assert.match(longTitleShape, /sz="2400"/);
  assert.match(longSubtitleShape, /<a:off x="502920" y="1389888"\/>/);
  assert.match(Object.entries(files).filter(([name]) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name)).map(([, value]) => strFromU8(value)).join("\n"), /risk-register/);
});
