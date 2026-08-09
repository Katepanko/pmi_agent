export type ProviderName = "openai" | "anthropic";

export type ModelRegistryEntry = {
  key: string;
  displayName: string;
  provider: ProviderName;
  modelId: string;
  available: boolean;
  unavailableReason?: string;
};

type RuntimeConfig = Record<string, string | undefined>;

const definitions = [
  ["anthropic-sonnet", "Anthropic Sonnet 5", "anthropic", "ANTHROPIC_SONNET_MODEL"],
  ["anthropic-opus", "Anthropic Opus 5", "anthropic", "ANTHROPIC_OPUS_MODEL"],
  ["openai-gpt55", "GPT-5.5", "openai", "OPENAI_GPT55_MODEL"],
  ["openai-gpt56", "GPT-5.6", "openai", "OPENAI_GPT56_MODEL"],
  ["openai-mini", "GPT-4 mini", "openai", "OPENAI_MINI_MODEL"],
] as const;

export function getModelRegistry(config: RuntimeConfig = process.env): ModelRegistryEntry[] {
  return definitions.map(([key, displayName, provider, modelEnv]) => {
    const modelId = config[modelEnv]?.trim() ?? "";
    const keyEnv = provider === "openai" ? "OPENAI_API_KEY" : "ANTHROPIC_API_KEY";
    const hasKey = Boolean(config[keyEnv]?.trim());
    const available = hasKey && Boolean(modelId);
    const missing = [!hasKey ? keyEnv : null, !modelId ? modelEnv : null].filter(Boolean);

    return {
      key,
      displayName,
      provider,
      modelId,
      available,
      unavailableReason: available ? undefined : `Configure ${missing.join(" and ")}`,
    };
  });
}

export function requireModel(modelKey: string, config: RuntimeConfig = process.env) {
  const model = getModelRegistry(config).find((entry) => entry.key === modelKey);
  if (!model) throw new Error(`Unknown model selection: ${modelKey}`);
  if (!model.available) throw new Error(model.unavailableReason ?? "Selected model is unavailable.");
  return model;
}
