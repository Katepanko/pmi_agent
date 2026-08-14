export function openAIStructuredOutputRequest(output?: {
  name: string;
  schema: Record<string, unknown>;
}) {
  if (!output) return {};
  return {
    text: {
      format: {
        type: "json_schema",
        name: output.name,
        strict: true,
        schema: output.schema,
      },
    },
    max_output_tokens: 12_000,
  };
}
