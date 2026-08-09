import type { GenerationRequest, LLMProvider } from "./provider";
import { mapServerSentEvents } from "./provider";

export class AnthropicProvider implements LLMProvider {
  constructor(private readonly apiKey: string) {}

  async generate(request: GenerationRequest) {
    const response = await this.request(request, false);
    const payload = (await response.json()) as {
      content?: Array<{ type: string; text?: string }>;
      error?: { message?: string };
    };
    if (!response.ok) throw new Error(payload.error?.message ?? "Anthropic request failed.");
    return payload.content?.filter((part) => part.type === "text").map((part) => part.text ?? "").join("") ?? "";
  }

  async stream(request: GenerationRequest) {
    const response = await this.request(request, true);
    if (!response.ok || !response.body) throw new Error(await response.text());
    return mapServerSentEvents(response.body, (event) => {
      const delta = event.delta as { type?: string; text?: string } | undefined;
      return event.type === "content_block_delta" && delta?.type === "text_delta" ? delta.text ?? null : null;
    });
  }

  private request(request: GenerationRequest, stream: boolean) {
    return fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "x-api-key": this.apiKey,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
      },
      body: JSON.stringify({
        model: request.model.modelId,
        system: request.system,
        messages: request.messages,
        max_tokens: 8_000,
        stream,
      }),
      signal: request.signal,
    });
  }
}
