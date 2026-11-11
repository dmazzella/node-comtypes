// Work done in steps: a generator that yields between units of work (one type
// of a type library). It runs to completion synchronously, or in time slices
// that leave the event loop free in between (timers, I/O, an Electron UI).

/**
 * Work that yields between its units; it may also yield a promise to wait for
 * (only when run with {@link runSliced}), and receives its value back.
 */
export type Steps<T> = Generator<Promise<unknown> | void, T, unknown>;

/** Run `steps` to completion on the current thread. */
export function runSync<T>(steps: Steps<T>): T {
  for (;;) {
    const step = steps.next();
    if (step.done) return step.value;
    if (step.value instanceof Promise) {
      steps.return(undefined as T);
      throw new Error("work that waits for a promise cannot run synchronously");
    }
  }
}

/** Milliseconds of work between two turns of the event loop. */
const SLICE_MS = 8;

const nextTurn = () => new Promise<void>((resolve) => setImmediate(resolve));

/**
 * Run `steps` in slices of a few milliseconds, yielding to the event loop
 * between them. Starts after the current synchronous code, like any promise.
 */
export async function runSliced<T>(steps: Steps<T>): Promise<T> {
  await null;
  let deadline = performance.now() + SLICE_MS;
  let step = steps.next();
  for (;;) {
    if (step.done) return step.value;
    if (step.value instanceof Promise) {
      let value: unknown;
      try {
        value = await step.value;
      } catch (error) {
        // Delivered as an exception at the yield point.
        step = steps.throw(error);
        deadline = performance.now() + SLICE_MS;
        continue;
      }
      step = steps.next(value);
      deadline = performance.now() + SLICE_MS;
      continue;
    }
    if (performance.now() >= deadline) {
      await nextTurn();
      deadline = performance.now() + SLICE_MS;
    }
    step = steps.next();
  }
}
