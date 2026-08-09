import type { ModelRegistryEntry } from "../models";

export type LLMMessage = { role: "user" | "assistant"; content: string };

export type GenerationRequest = {
  model: ModelRegistryEntry;
  system: string;
  messages: LLMMessage[];
  signal?: AbortSignal;
};

export interface LLMProvider {
  generate(request: GenerationRequest): Promise<string>;
  stream(request: GenerationRequest): Promise<ReadableStream<Uint8Array>>;
}

export function textStream(text: string) {
  const encoder = new TextEncoder();
  return new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(encoder.encode(text));
      controller.close();
    },
  });
}

export function mapServerSentEvents(
  body: ReadableStream<Uint8Array>,
  selectText: (payload: Record<string, unknown>) => string | null,
) {
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  let buffer = "";

  return body.pipeThrough(
    new TransformStream<Uint8Array, Uint8Array>({
      transform(chunk, controller) {
        buffer += decoder.decode(chunk, { stream: true });
        const lines = buffer.split("\n");
        buffer = lines.pop() ?? "";
        for (const line of lines) {
          if (!line.startsWith("data:")) continue;
          const data = line.slice(5).trim();
          if (!data || data === "[DONE]") continue;
          try {
            const selected = selectText(JSON.parse(data) as Record<string, unknown>);
            if (selected) controller.enqueue(encoder.encode(selected));
          } catch {
            // Provider heartbeats and incomplete events intentionally emit no text.
          }
        }
      },
    }),
  );
}
