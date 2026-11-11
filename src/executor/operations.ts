// Asynchronous COM operations and their cancellation, as node-python's
// `run(promise, { signal })`: an operation records the calls it queues on the
// executor, and aborting it removes those not started yet; calls already
// running complete. No AbortController or listener per call: only `run()`
// with a signal pays for cancellation.

import type * as ctypes from "node-ctypes";
import { cancelJob } from "./executor.ts";

/** The executor calls of one asynchronous operation, cancelled by `run()`. */
export class Operation {
  #aborted = false;
  #reason: unknown;
  // The first queued call, then the others (an operation usually makes one).
  #executor: ctypes.Executor | null = null;
  #job = 0;
  #more: [ctypes.Executor, number][] | null = null;

  /** Whether `run()` aborted the operation: its next calls are rejected with {@link reason}. */
  get aborted(): boolean {
    return this.#aborted;
  }

  get reason(): unknown {
    return this.#reason;
  }

  /** Record call `job` just queued on `executor`. */
  queued(executor: ctypes.Executor, job: number): void {
    if (this.#executor === null) {
      this.#executor = executor;
      this.#job = job;
    } else {
      (this.#more ??= []).push([executor, job]);
    }
  }

  /** Remove the queued calls not started yet; their promises reject with `reason`. */
  abort(reason: unknown): void {
    if (this.#aborted) return;
    this.#aborted = true;
    this.#reason = reason;
    if (this.#executor) cancelJob(this.#executor, this.#job, reason);
    for (const [executor, job] of this.#more ?? []) cancelJob(executor, job, reason);
    this.#executor = null;
    this.#more = null;
  }
}

const operations = new WeakMap<Promise<unknown>, Operation>();

/** Start an operation whose executor calls are recorded in `operation`; `run()` can then cancel it. */
export function operation<T>(start: (operation: Operation) => Promise<T>): Promise<T> {
  const current = new Operation();
  const promise = start(current);
  operations.set(promise, current);
  return promise;
}

/** `derived` (e.g. a converted result of `source`) cancels the operation of `source`. */
export function linkOperation<T>(derived: Promise<T>, source: Promise<unknown>): Promise<T> {
  const current = operations.get(source);
  if (current) operations.set(derived, current);
  return derived;
}

export interface RunOptions {
  /** Aborting it rejects with `signal.reason` and skips the operation's queued executor calls. */
  readonly signal?: AbortSignal;
}

function isThenable(value: unknown): value is PromiseLike<unknown> {
  return value !== null && (typeof value === "object" || typeof value === "function") && typeof (value as { then?: unknown }).then === "function";
}

/**
 * Await a COM operation (the promise of a method call, `CreateObject`,
 * `GetObject`, ...) with an AbortSignal: `run(obj.Slow(), { signal:
 * AbortSignal.timeout(1000) })`. On abort the promise rejects with
 * `signal.reason`; executor calls still queued are skipped, a running call
 * completes (COM does not interrupt it). Python equivalent: none
 * (node-python's `python.run`).
 */
export function run<T>(promise: PromiseLike<T>, { signal }: RunOptions = {}): Promise<T> {
  if (!isThenable(promise)) return Promise.reject(new TypeError("run() needs the promise of a COM operation"));
  if (!signal) return Promise.resolve(promise);
  const current = operations.get(promise as Promise<unknown>);
  return new Promise<T>((resolve, reject) => {
    const abort = () => {
      current?.abort(signal.reason);
      reject(signal.reason);
    };
    Promise.resolve(promise).then(
      (value) => {
        signal.removeEventListener("abort", abort);
        resolve(value);
      },
      (error: unknown) => {
        signal.removeEventListener("abort", abort);
        reject(error);
      },
    );
    if (signal.aborted) abort();
    else signal.addEventListener("abort", abort, { once: true });
  });
}
