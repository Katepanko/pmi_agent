import { AnthropicProvider } from "./anthropic";
import { OpenAIProvider } from "./openai";
import type { LLMProvider } from "./provider";
import type { ProviderName } from "../models";

export function getProvider(name: ProviderName): LLMProvider {
  if (name === "openai") return new OpenAIProvider(process.env.OPENAI_API_KEY ?? "");
  if (name === "anthropic") return new AnthropicProvider(process.env.ANTHROPIC_API_KEY ?? "");
  throw new Error(`Unsupported provider: ${String(name)}`);
}
