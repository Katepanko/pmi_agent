import { AsyncLocalStorage } from "node:async_hooks";

export type RuntimeBindings = {
  DB?: D1Database;
  FILES?: R2Bucket;
};

type RuntimeGlobal = typeof globalThis & {
  __PMI_AGENT_RUNTIME_BINDINGS__?: AsyncLocalStorage<RuntimeBindings>;
};

const runtimeGlobal = globalThis as RuntimeGlobal;
const bindingStorage = runtimeGlobal.__PMI_AGENT_RUNTIME_BINDINGS__ ??= new AsyncLocalStorage<RuntimeBindings>();

export function withRuntimeBindings<T>(bindings: RuntimeBindings, work: () => T): T {
  return bindingStorage.run(bindings, work);
}

export function getRuntimeBindings(): RuntimeBindings {
  return bindingStorage.getStore() ?? {};
}
