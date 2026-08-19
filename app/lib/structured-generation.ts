import type { LLMProvider, StructuredOutput } from "./llm/provider.ts";
import type { ModelRegistryEntry } from "./models.ts";

export async function generateStructuredModel<T>(input: {
  provider: LLMProvider;
  model: ModelRegistryEntry;
  system: string;
  userMessage: string;
  structuredOutput: StructuredOutput;
  parse: (raw: string) => T;
  outputLabel: string;
  signal?: AbortSignal;
}): Promise<T> {
  const firstOutput = await input.provider.generate({
    model: input.model,
    system: input.system,
    messages: [{ role: "user", content: input.userMessage }],
    structuredOutput: input.structuredOutput,
    signal: input.signal,
  });

  try {
    return input.parse(firstOutput);
  } catch (firstParseError) {
    const firstReason = firstParseError instanceof Error ? firstParseError.message : "invalid structured output";
    const repairedOutput = await input.provider.generate({
      model: input.model,
      system: `${input.system}\n\nThe previous response could not be parsed (${firstReason}). Return the complete model again. Output only one JSON object that conforms exactly to the required schema.`,
      messages: [
        { role: "user", content: input.userMessage },
        { role: "assistant", content: firstOutput.slice(0, 24_000) },
        { role: "user", content: "Repair the response as complete valid JSON. Do not explain the repair." },
      ],
      structuredOutput: input.structuredOutput,
      signal: input.signal,
    });

    try {
      return input.parse(repairedOutput);
    } catch (secondParseError) {
      const secondReason = secondParseError instanceof Error ? secondParseError.message : "invalid structured output";
      throw new Error(`The model failed to return a valid ${input.outputLabel} after two attempts: ${secondReason}`);
    }
  }
}
