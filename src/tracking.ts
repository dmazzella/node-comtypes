import { rawRelease } from "./low-level/vtable.ts";

/** One owned interface reference. */
export interface Reference {
  readonly address: bigint;
  /** Apartment generation the reference belongs to. */
  readonly generation: number;
  released: boolean;
}

// Per-thread state: each worker has its own module instance.
const state = {
  /** Successful CoInitializeEx calls not yet balanced by CoUninitialize. */
  initCount: 0,
  /** Incremented when the apartment is torn down; stale references are never released. */
  generation: 0,
  live: new Set<Reference>(),
};

// Safety net only: deterministic release (`Release()`, `using`) is the rule.
const registry = new FinalizationRegistry<Reference>((reference) => {
  if (reference.released || reference.generation !== state.generation) return;
  reference.released = true;
  state.live.delete(reference);
  try {
    rawRelease(reference.address);
  } catch {
    // Nothing can be reported from a finalizer.
  }
});

export function track(address: bigint, owner: object): Reference {
  const reference: Reference = { address, generation: state.generation, released: false };
  state.live.add(reference);
  registry.register(owner, reference, reference);
  return reference;
}

/** Forget a reference without releasing it (already released or detached). */
export function untrack(reference: Reference): void {
  reference.released = true;
  state.live.delete(reference);
  registry.unregister(reference);
}

/** Release all live references, newest first. Returns how many were released. */
export function releaseAll(): number {
  const references = [...state.live].reverse();
  for (const reference of references) {
    untrack(reference);
    try {
      rawRelease(reference.address);
    } catch {
      // Keep releasing the others.
    }
  }
  return references.length;
}

/** Apartment generation of this thread: changes when the apartment is torn down. */
export function apartmentGeneration(): number {
  return state.generation;
}

export function liveCount(): number {
  return state.live.size;
}

export function initCount(): number {
  return state.initCount;
}

export function initialized(): void {
  state.initCount++;
}

/** Account for one CoUninitialize; on the last one, release everything and start a new generation. */
export function uninitializing(): void {
  if (state.initCount === 0) return;
  if (state.initCount === 1) {
    releaseAll();
    state.generation++;
  }
  state.initCount--;
}
