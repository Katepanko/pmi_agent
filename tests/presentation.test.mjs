import assert from "node:assert/strict";
import test from "node:test";
import { strFromU8, unzipSync } from "fflate";
import {
  isPresentationRequest,
  parsePresentationModel,
  renderPresentation,
} from "../app/lib/presentation.ts";
import { extractOpenAIResponseText } from "../app/lib/llm/openai-response.ts";

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
    ],
  }), "Steering Committee");

  const bytes = await renderPresentation(model);
  assert.ok(bytes.byteLength > 50_000);
  const files = unzipSync(bytes);
  assert.ok(files["[Content_Types].xml"]);
  assert.ok(files["ppt/presentation.xml"]);
  assert.ok(files["ppt/slideMasters/slideMaster1.xml"]);
  assert.ok(files["ppt/slides/slide1.xml"]);
  assert.ok(files["ppt/slides/slide2.xml"]);
  assert.equal(Object.keys(files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length, 2);
  assert.match(strFromU8(files["ppt/slides/slide2.xml"]), /IT cutover now defines/);
  assert.match(Object.entries(files).filter(([name]) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name)).map(([, value]) => strFromU8(value)).join("\n"), /risk-register/);
});
