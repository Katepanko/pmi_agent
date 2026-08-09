import type { GenerationRequest, LLMProvider } from "./provider";
import { mapServerSentEvents } from "./provider";

export class OpenAIProvider implements LLMProvider {
  constructor(private readonly apiKey: string) {}

  async generate(request: GenerationRequest) {
    const response = await this.request(request, false);
    const payload = (await response.json()) as { output_text?: string; error?: { message?: string } };
    if (!response.ok) throw new Error(payload.error?.message ?? "OpenAI request failed.");
    return payload.output_text ?? "";
  }

  async stream(request: GenerationRequest) {
    const response = await this.request(request, true);
    if (!response.ok || !response.body) throw new Error(await response.text());
    return mapServerSentEvents(response.body, (event) =>
      event.type === "response.output_text.delta" && typeof event.delta === "string"
        ? event.delta
        : null,
    );
  }

  private request(request: GenerationRequest, stream: boolean) {
    return fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        authorization: `Bearer ${this.apiKey}`,
        "content-type": "application/json",
      },
      body: JSON.stringify({
        model: request.model.modelId,
        instructions: request.system,
        input: request.messages,
        stream,
      }),
      signal: request.signal,
    });
  }
}
