export type AnthropicContentPart = {
  type?: string;
  text?: string;
  name?: string;
  input?: unknown;
};

const CLAUDE_55_MODEL = /^claude-(?:opus|sonnet)-5-5(?:-|$)/;
const UNSUPPORTED_SCHEMA_CONSTRAINTS = new Set(["minimum", "maximum", "exclusiveMinimum", "exclusiveMaximum", "minLength", "maxLength"]);

export function usesAnthropicJsonOutput(modelId: string) {
  return CLAUDE_55_MODEL.test(modelId);
}

function supportedAnthropicSchema(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(supportedAnthropicSchema);
  if (!value || typeof value !== "object") return value;
  return Object.fromEntries(Object.entries(value as Record<string, unknown>)
    .filter(([key]) => !UNSUPPORTED_SCHEMA_CONSTRAINTS.has(key))
    .map(([key, entry]) => [key, supportedAnthropicSchema(entry)]));
}

export function anthropicStructuredOutputRequest(output?: {
  name: string;
  schema: Record<string, unknown>;
}, modelId = "") {
  if (!output) return {};
  if (usesAnthropicJsonOutput(modelId)) {
    return {
      output_config: {
        format: {
          type: "json_schema",
          schema: supportedAnthropicSchema(output.schema),
        },
      },
    };
  }
  return {
    tools: [{
      name: output.name,
      description: "Return the complete requested artifact content model. This tool records the final answer; do not return the model as prose.",
      input_schema: output.schema,
    }],
    tool_choice: { type: "tool", name: output.name },
  };
}

export function extractAnthropicResponse(parts: AnthropicContentPart[] | undefined, structuredOutputName?: string) {
  if (structuredOutputName) {
    const toolUse = parts?.find((part) => part.type === "tool_use" && part.name === structuredOutputName);
    if (toolUse?.input && typeof toolUse.input === "object") return JSON.stringify(toolUse.input);
  }
  return parts?.filter((part) => part.type === "text").map((part) => part.text ?? "").join("") ?? "";
}
