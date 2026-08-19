import assert from "node:assert/strict";
import test from "node:test";
import { generateStructuredModel } from "../app/lib/structured-generation.ts";
import { artifactStructuredOutput } from "../app/lib/artifact-schema.ts";
import { openAIStructuredOutputRequest } from "../app/lib/llm/openai-structured-output.ts";
import { anthropicStructuredOutputRequest, extractAnthropicResponse } from "../app/lib/llm/anthropic-structured-output.ts";

const model = {
  key: "openai-gpt56",
  displayName: "GPT-5.6",
  provider: "openai",
  modelId: "gpt-5.6",
  available: true,
};

const validPresentation = JSON.stringify({
  title: "PMI SteerCo Update",
  subtitle: null,
  projectName: null,
  location: null,
  date: null,
  audience: "Steering Committee",
  executiveSummary: "One dependency requires Steering Committee attention.",
  slides: [{
    title: "One dependency requires Steering Committee attention",
    kicker: "Executive summary",
    keyMessage: null,
    layout: "summary",
    items: [],
    sourceNotes: [],
  }],
});

test("artifact planning requests the strict presentation schema", async () => {
  const calls = [];
  const provider = {
    async generate(request) {
      calls.push(request);
      return validPresentation;
    },
  };

  const result = await generateStructuredModel({
    provider,
    model,
    system: "Plan a presentation.",
    userMessage: "Create a PowerPoint for SteerCo.",
    structuredOutput: artifactStructuredOutput("pptx"),
    parse: JSON.parse,
    outputLabel: "PPTX content model",
  });

  assert.equal(result.title, "PMI SteerCo Update");
  assert.equal(calls.length, 1);
  assert.equal(calls[0].structuredOutput.name, "pmi_presentation");
  assert.equal(calls[0].structuredOutput.schema.additionalProperties, false);
  assert.ok(calls[0].structuredOutput.schema.required.includes("slides"));
});

test("artifact planning repairs one malformed response before failing the request", async () => {
  const calls = [];
  const provider = {
    async generate(request) {
      calls.push(request);
      return calls.length === 1 ? "I will prepare the presentation." : validPresentation;
    },
  };

  const result = await generateStructuredModel({
    provider,
    model,
    system: "Plan a presentation.",
    userMessage: "Create a PowerPoint for SteerCo.",
    structuredOutput: artifactStructuredOutput("pptx"),
    parse: JSON.parse,
    outputLabel: "PPTX content model",
  });

  assert.equal(result.slides.length, 1);
  assert.equal(calls.length, 2);
  assert.match(calls[1].system, /previous response could not be parsed/i);
  assert.match(calls[1].messages.at(-1).content, /complete valid JSON/i);
});

test("OpenAI artifact requests send the Responses API text.format schema", () => {
  const requestBody = openAIStructuredOutputRequest(artifactStructuredOutput("pptx"));

  assert.equal(requestBody.text.format.type, "json_schema");
  assert.equal(requestBody.text.format.name, "pmi_presentation");
  assert.equal(requestBody.text.format.strict, true);
  assert.equal(requestBody.max_output_tokens, 12_000);
});

test("Anthropic artifact requests force a schema-backed tool result", () => {
  const structuredOutput = artifactStructuredOutput("pptx");
  const requestBody = anthropicStructuredOutputRequest(structuredOutput);

  assert.equal(requestBody.tools[0].name, "pmi_presentation");
  assert.equal(requestBody.tools[0].input_schema.additionalProperties, false);
  assert.deepEqual(requestBody.tool_choice, { type: "tool", name: "pmi_presentation" });
});

test("Anthropic artifact generation serializes tool input instead of parsing model-authored JSON text", () => {
  const raw = extractAnthropicResponse([{
    type: "tool_use",
    name: "pmi_presentation",
    input: JSON.parse(validPresentation),
  }], "pmi_presentation");

  assert.deepEqual(JSON.parse(raw), JSON.parse(validPresentation));
});
