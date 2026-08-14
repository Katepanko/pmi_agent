export type AnthropicContentPart = {
  type?: string;
  text?: string;
  name?: string;
  input?: unknown;
};

export function anthropicStructuredOutputRequest(output?: {
  name: string;
  schema: Record<string, unknown>;
}) {
  if (!output) return {};
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
    return "";
  }
  return parts?.filter((part) => part.type === "text").map((part) => part.text ?? "").join("") ?? "";
}
