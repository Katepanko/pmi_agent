import type { GenerationRequest, LLMProvider } from "./provider";
import { mapServerSentEvents } from "./provider";
import { extractOpenAIResponseText, type OpenAIResponsePayload } from "./openai-response";

export class OpenAIProvider implements LLMProvider {
  private readonly apiKey: string;

  constructor(apiKey: string) {
    this.apiKey = apiKey;
  }

  async generate(request: GenerationRequest) {
    const response = await this.request(request, false);
    const payload = (await response.json()) as OpenAIResponsePayload;
    if (!response.ok) throw new Error(payload.error?.message ?? "OpenAI request failed.");
    const output = extractOpenAIResponseText(payload);
    if (!output.trim()) {
      const reason = payload.incomplete_details?.reason
        ? ` The response was incomplete: ${payload.incomplete_details.reason}.`
        : payload.status ? ` Response status: ${payload.status}.` : "";
      throw new Error(`OpenAI returned no text output.${reason}`);
    }
    return output;
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
