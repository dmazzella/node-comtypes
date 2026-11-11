// The COM executor: a node-ctypes Executor (a native thread with a FIFO queue)
// in the multithreaded apartment. Native calls prepared on the JavaScript
// thread run there one at a time, in arrival order, so that the JavaScript
// thread never waits for them. IErrorInfo is per thread: the job queued right
// after each HRESULT call takes it on the same thread.
import * as ctypes from "node-ctypes";
import { COINIT_MULTITHREADED } from "../constants.ts";
import type { COMErrorDetails } from "../errors.ts";
import { GUID } from "../guid.ts";
import type { GUIDLike } from "../guid.ts";
import { ole32, oleaut32, takeBSTR } from "../low-level/native.ts";
import { functionCache, vtableEntry } from "../low-level/vtable.ts";
import type { Operation } from "./operations.ts";

const { c_int32, c_uint32, c_void_p } = ctypes;

/** Counters of the COM executor (`stats()`). */
export interface ExecutorStats {
  /** Whether the executor thread is running. */
  readonly running: boolean;
  /** Operations queued or running. */
  readonly pending: number;
  /** Operations settled by the current executor. */
  readonly completed: number;
  /** Incremented each time a new executor starts: references owned by an earlier one are gone. */
  readonly generation: number;
}

export interface NativeCall {
  /** Function address (e.g. a vtable entry). */
  readonly fn: bigint;
  readonly restype: ctypes.FFIType | null;
  /** Native argument types; keep the same array for the same signature (bindings are cached by it). */
  readonly argtypes: readonly ctypes.FFIType[];
  readonly args: readonly unknown[];
  /** On an HRESULT failure, read IErrorInfo for `object` and interface `iid` (when it supports error info). */
  readonly errorInfo?: { readonly object: bigint; readonly iid: GUIDLike };
}

export interface NativeResult {
  readonly value: unknown;
  /** IErrorInfo details read on the executor thread, for a failed HRESULT. */
  readonly details: COMErrorDetails | null;
}

export interface CallOptions {
  /** Cancels the call while it is still queued: the promise rejects with `signal.reason`. */
  readonly signal?: AbortSignal;
  /** Records the call in this operation, which `run()` can cancel (used by node-comtypes' own calls). */
  readonly operation?: Operation;
}

const IID_ISupportErrorInfo = new GUID("{DF0B3D60-548F-101B-8E65-08002B2BD119}");
// IErrorInfo getters and ISupportErrorInfo::InterfaceSupportsErrorInfo share
// the HRESULT (this, pointer) signature.
const PointerArgAt = functionCache(() => ctypes.WINFUNCTYPE(c_int32, c_void_p, c_void_p));
const QueryInterfaceAt = functionCache(() => ctypes.WINFUNCTYPE(c_int32, c_void_p, c_void_p, c_void_p));
const ReleaseAt = functionCache(() => ctypes.WINFUNCTYPE(c_uint32, c_void_p));

let executor: ctypes.Executor | null = null;
let closing: Promise<void> | null = null;
let generation = 0;
let exitHandler: (() => void) | null = null;

/** The running executor, started on first use with its thread in the multithreaded apartment. */
function current(): ctypes.Executor {
  if (executor) return executor;
  const created = new ctypes.Executor({ name: "node-comtypes" });
  // The first job: every later one runs in the multithreaded apartment. A
  // failure surfaces as CO_E_NOTINITIALIZED from the calls themselves.
  ole32()
    .CoInitializeEx.callOn(created, null, COINIT_MULTITHREADED)
    .catch(() => undefined);
  executor = created;
  generation++;
  if (!exitHandler) {
    // As node-python: finalize on a normal exit, so that proxies to other
    // processes are released; a forced process.exit() abandons the thread.
    exitHandler = () => void finalize();
    process.on("beforeExit", exitHandler);
  }
  return created;
}

/** The running executor (started on demand), for code that queues its own calls. */
export function executorTarget(): ctypes.Executor {
  return current();
}

/** Counters of the COM executor. Python equivalent: node-python's `python.stats()`. */
export function stats(): ExecutorStats {
  return { running: executor?.running ?? false, pending: executor?.pending ?? 0, completed: executor?.completed ?? 0, generation };
}

/** Whether references owned by executor `generation` are still alive (its executor is running). */
export function executorAlive(generation: number): boolean {
  return executor !== null && generation === executorGeneration();
}

/** Generation of the current executor (0 before the first one starts). */
export function executorGeneration(): number {
  return generation;
}

/**
 * Drain the operations already accepted by the COM executor, leave its
 * apartment and stop its thread; resolves once their promises have settled. Called automatically on a normal
 * `beforeExit`. A later asynchronous operation starts a new executor.
 */
export function finalize(): Promise<void> {
  if (closing) return closing;
  const stopping = executor;
  if (!stopping) return Promise.resolve();
  executor = null;
  if (exitHandler) {
    process.off("beforeExit", exitHandler);
    exitHandler = null;
  }
  // FIFO: after every accepted call, before the thread stops.
  ole32()
    .CoUninitialize.callOn(stopping)
    .catch(() => undefined);
  // The promises of the drained operations settle a few microtasks after their
  // calls: one more turn of the event loop lets them settle first.
  closing = stopping
    .close()
    .then(() => new Promise<void>((resolve) => setImmediate(resolve)))
    .finally(() => {
      closing = null;
    });
  return closing;
}

type Bound = ((...args: unknown[]) => unknown) & { callOn(target: ctypes.ExecutorTarget, ...args: unknown[]): Promise<unknown> };

const prototypes = new WeakMap<readonly ctypes.FFIType[], Map<ctypes.FFIType | null, (address: bigint) => Bound>>();

/** The native function at `address`, bound once per signature and address. */
function bind(restype: ctypes.FFIType | null, argtypes: readonly ctypes.FFIType[], address: bigint): Bound {
  let byRestype = prototypes.get(argtypes);
  if (!byRestype) prototypes.set(argtypes, (byRestype = new Map()));
  let cache = byRestype.get(restype);
  if (!cache) byRestype.set(restype, (cache = functionCache(() => ctypes.WINFUNCTYPE(restype, ...argtypes)) as unknown as (address: bigint) => Bound));
  return cache(address);
}

// node-ctypes numbers the jobs of an Executor. A callOn target with `_job`
// receives the id of its own call; `_lastJob` is the id of the last call
// submitted and `_cancel(id, reason)` removes a call still queued (what its
// own callOn does with an AbortSignal). They are not part of the typed API,
// so they are reached here only.
type ExecutorJobs = { readonly _lastJob: number; _cancel(job: number, reason: unknown): boolean };

/** A callOn target that receives the id of its call in `_job` (0 when it was not queued). */
export interface CallTarget {
  readonly executor: ctypes.Executor;
  readonly signal?: AbortSignal;
  _job: number;
}

/** Id of the last call submitted to `target`: another call was submitted after a call whose id is lower. */
export function lastJob(target: ctypes.Executor): number {
  return (target as unknown as ExecutorJobs)._lastJob;
}

/** Remove call `job` from the queue of `target` if it has not started: its promise rejects with `reason`. */
export function cancelJob(target: ctypes.Executor, job: number, reason: unknown): boolean {
  return (target as unknown as ExecutorJobs)._cancel(job, reason);
}

/** Queue `GetErrorInfo` on `target`: the thread's error object (owned), or `null`. */
function takeErrorObject(target: ctypes.Executor): Promise<bigint | null> {
  const out = new c_void_p();
  return oleaut32()
    .GetErrorInfo.callOn(target, 0, ctypes.byref(out))
    .then((hr) => (hr === 0 && out.value ? out.value : null));
}

function release(target: ctypes.Executor, address: bigint): void {
  ReleaseAt(vtableEntry(address, 2))
    .callOn(target, address)
    .catch(() => undefined);
}

/** Whether `object` declares error information for interface `iid` (ISupportErrorInfo), on the executor. */
async function supportsErrorInfo(target: ctypes.Executor, object: bigint, iid: GUID): Promise<boolean> {
  const out = new c_void_p();
  const hr = await QueryInterfaceAt(vtableEntry(object, 0)).callOn(target, object, IID_ISupportErrorInfo._buffer, ctypes.byref(out));
  if (hr < 0 || !out.value) return false;
  const support = out.value;
  try {
    return (await PointerArgAt(vtableEntry(support, 3)).callOn(target, support, iid._buffer)) === 0;
  } finally {
    release(target, support);
  }
}

/** Read an error object on the executor, where it belongs (it may be a proxy). */
async function readErrorObject(target: ctypes.Executor, errorObject: bigint): Promise<COMErrorDetails> {
  const guid = new GUID();
  const context = new c_uint32();
  const strings = [4, 5, 6].map(() => new c_void_p());
  const [hasGuid, source, description, helpFile, hasContext] = await Promise.all([
    PointerArgAt(vtableEntry(errorObject, 3)).callOn(target, errorObject, guid._buffer),
    ...strings.map((out, index) => PointerArgAt(vtableEntry(errorObject, 4 + index)).callOn(target, errorObject, ctypes.byref(out))),
    PointerArgAt(vtableEntry(errorObject, 7)).callOn(target, errorObject, ctypes.byref(context)),
  ]);
  // BSTRs are allocated by the task allocator: freed from any thread.
  const text = (hr: number, out: InstanceType<typeof c_void_p>) => (hr >= 0 ? takeBSTR(out.value) : null);
  return {
    guid: hasGuid >= 0 && !guid.isNull() ? guid : null,
    source: text(source, strings[0]),
    description: text(description, strings[1]),
    helpFile: text(helpFile, strings[2]),
    helpContext: hasContext >= 0 ? context.value : 0,
  };
}

async function errorDetails(target: ctypes.Executor, errorObject: Promise<bigint | null>, errorInfo: NonNullable<NativeCall["errorInfo"]>): Promise<COMErrorDetails | null> {
  const taken = await errorObject;
  if (!taken) return null;
  try {
    // As Python comtypes: only for an object that declares support for the interface.
    if (!(await supportsErrorInfo(target, errorInfo.object, new GUID(errorInfo.iid)))) return null;
    return await readErrorObject(target, taken);
  } catch {
    return null;
  } finally {
    release(target, taken);
  }
}

/**
 * Run a raw native call (e.g. a vtable entry, with `this` as the first
 * argument) on the COM executor thread. Memory passed by address stays
 * referenced until the call completes; for HRESULT failures, `errorInfo`
 * reads `IErrorInfo` on the executor thread.
 *
 * As Python comtypes, `GetErrorInfo` runs only after a failure, and it must
 * be the next call on the executor thread: a later call could replace or clear
 * the error object. That holds when nothing was submitted after the failed
 * call (code that awaits each call); otherwise, e.g. in a burst of calls
 * already queued, the failure has no details.
 */
export function executorCall(call: NativeCall, options: CallOptions = {}): Promise<NativeResult> {
  // A single promise step after the call (no async function): this runs for every method call.
  let target: ctypes.Executor;
  let pending: Promise<unknown>;
  let job: number;
  try {
    target = current();
    const fn = bind(call.restype, call.argtypes, call.fn);
    if (options.operation?.aborted) throw options.operation.reason;
    // The id of this call, written by node-ctypes for this invocation only.
    const callTarget: CallTarget = { executor: target, signal: options.signal, _job: 0 };
    pending = fn.callOn(callTarget, ...call.args);
    job = callTarget._job;
    if (job !== 0) options.operation?.queued(target, job);
  } catch (error) {
    return Promise.reject(error);
  }
  const { errorInfo } = call;
  return pending.then((value) => {
    if (!(typeof value === "number" && value < 0) || !errorInfo || lastJob(target) !== job) return { value, details: null };
    return errorDetails(target, takeErrorObject(target), errorInfo).then((details) => ({ value, details }));
  });
}
