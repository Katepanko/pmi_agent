import { AsyncLocalStorage } from "node:async_hooks";

export type RuntimeBindings = {
  DB?: D1Database;
  FILES?: R2Bucket;
};

type RuntimeGlobal = typeof globalThis & {
  __PMI_AGENT_RUNTIME_BINDINGS__?: AsyncLocalStorage<RuntimeBindings>;
  __PMI_AGENT_RUNTIME_BINDINGS_FALLBACK__?: RuntimeBindings;
};

const runtimeGlobal = globalThis as RuntimeGlobal;
const bindingStorage = runtimeGlobal.__PMI_AGENT_RUNTIME_BINDINGS__ ??= new AsyncLocalStorage<RuntimeBindings>();

export function withRuntimeBindings<T>(bindings: RuntimeBindings, work: () => T): T {
  // Cloudflare bindings are deployment-scoped and identical for every request in
  // an isolate. Keep a stable fallback in addition to the request-local store so
  // long model/rendering chains cannot lose access when an async boundary drops
  // AsyncLocalStorage context.
  runtimeGlobal.__PMI_AGENT_RUNTIME_BINDINGS_FALLBACK__ = bindings;
  return bindingStorage.run(bindings, work);
}

export function getRuntimeBindings(): RuntimeBindings {
  return bindingStorage.getStore() ?? runtimeGlobal.__PMI_AGENT_RUNTIME_BINDINGS_FALLBACK__ ?? {};
}
