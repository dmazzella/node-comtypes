import * as ctypes from "node-ctypes";
import { COMError } from "./errors.ts";
import { errorInfoFor } from "./errorinfo.ts";
import { handOver, homeApartment, isAgile, isProxy, referenceForExecutor, releaseOnExecutor } from "./executor/apartment.ts";
import type { ResultLocation } from "./executor/apartment.ts";
import { executorAlive, executorCall, executorGeneration } from "./executor/executor.ts";
import { operation } from "./executor/operations.ts";
import type { Operation } from "./executor/operations.ts";
import { GUID } from "./guid.ts";
import { E_NOINTERFACE } from "./hresult.ts";
import type { InterfaceDef, InterfacePointer, MethodSpec, SyncPointer } from "./interface.ts";
import { functionCache, rawQueryInterface, rawRelease, vtableEntry } from "./low-level/vtable.ts";
import { track, untrack } from "./tracking.ts";
import type { Reference } from "./tracking.ts";

type Callable = (...args: any[]) => any;

/** How an asynchronous call reaches the executor. */
export interface AsyncContext {
  /** The home thread is single-threaded: interface results are handed back to it. */
  readonly handOver: boolean;
  /** Further result locations to hand over (e.g. the VARIANT of `IDispatch::Invoke`). */
  readonly results?: readonly ResultLocation[];
  /** The operation recording the call: `run()` cancels it while it is queued. */
  readonly operation?: Operation;
}

/** A vtable slot of an interface: binds the function found at a given address. */
export interface Slot {
  readonly spec: MethodSpec;
  /** `Interface::Method`, used in error messages. */
  readonly label: string;
  bind(address: bigint): Callable;
  /** Call the bound function: raw arguments, or converted by the spec's marshal (COMMETHOD). */
  call(fn: Callable, address: bigint, args: readonly unknown[]): unknown;
  /** Call the function `entry` of object `target` on the COM executor. */
  callAsync(entry: bigint, target: bigint, args: readonly unknown[], context: AsyncContext): Promise<unknown>;
}

/** Interface declaring a slot; `definition` is set once the interface is complete. */
export interface SlotOwner {
  readonly name: string;
  readonly iid: GUID;
  definition: InterfaceDef<any> | null;
}

const IID_IUnknown = new GUID("{00000000-0000-0000-C000-000000000046}");

/**
 * Create the slot for `spec`, declared by interface `owner`. HRESULT methods
 * are bound with a `c_int32` restype and an errcheck raising {@link COMError},
 * with `IErrorInfo` details when the object supports them for `owner`.
 */
export function createSlot(spec: MethodSpec, owner: SlotOwner): Slot {
  const label = `${owner.name}::${spec.name}`;
  const hresult = spec.restype === ctypes.HRESULT;
  // IUnknown methods never set error info (and QueryInterface failures are common).
  const errorInfo = !owner.iid.equals(IID_IUnknown);
  const nativeRestype = hresult ? ctypes.c_int32 : spec.restype;
  // One array per slot: the executor caches its bindings by it.
  const nativeArgtypes: readonly ctypes.FFIType[] = [ctypes.c_void_p, ...spec.argtypes];
  const cache = functionCache(() => ctypes.WINFUNCTYPE(nativeRestype, ctypes.c_void_p, ...spec.argtypes)) as (address: bigint) => Callable;
  const errcheck = (result: number, _fn: unknown, args: unknown[]) => {
    if (result < 0) throw new COMError(result, { method: label, details: errorInfo ? errorInfoFor(args[0] as bigint, owner.iid) : null });
    return result;
  };
  const { marshal, marshalAsync } = spec;
  return {
    spec,
    label,
    call: marshal ? (fn, address, args) => marshal(fn, address, args, owner.definition!) : (fn, address, args) => fn(address, ...args),
    bind(address) {
      const fn = cache(address) as Callable & { errcheck: ctypes.ErrcheckCallback | null };
      if (hresult && fn.errcheck !== errcheck) fn.errcheck = errcheck;
      return fn;
    },
    callAsync(entry, target, args, context) {
      // Promise steps instead of an async function: this runs for every method call.
      const run = (native: readonly unknown[], results: readonly ResultLocation[]): Promise<unknown> =>
        executorCall({ fn: entry, restype: nativeRestype, argtypes: nativeArgtypes, args: [target, ...native], ...(hresult && errorInfo ? { errorInfo: { object: target, iid: owner.iid } } : {}) }, { operation: context.operation }).then(
          ({ value, details }) => {
            if (hresult && (value as number) < 0) throw new COMError(value as number, { method: label, details });
            const locations = context.results?.length ? [...results, ...context.results] : results;
            return context.handOver && locations.length ? handOver(locations).then(() => value) : value;
          },
        );
      return marshalAsync ? marshalAsync(run, args, owner.definition!) : run(args, []);
    },
  };
}

export interface FromAddressOptions {
  /**
   * `true` when the caller keeps its own reference (borrowed pointer): the
   * new pointer takes one more with `AddRef`. Default `false`: the new pointer
   * takes ownership of an existing reference (e.g. an out parameter).
   */
  readonly addRef?: boolean;
}

/**
 * How asynchronous calls reach the object:
 * - `home`: an object bound to the home single-threaded apartment, called
 *   there (synchronously: the work would run on the home thread anyway);
 * - `shared`: the same address on the executor (home in the MTA, or an agile object);
 * - `proxy`: the executor's own reference, from the Global Interface Table.
 */
/** Vtable entries of the address asynchronous calls go to, by slot. */
type Entries = { readonly address: bigint; readonly slots: bigint[] };

type Route = { readonly kind: "home" } | { readonly kind: "shared"; readonly handOver: boolean } | { readonly kind: "proxy"; readonly address: bigint; readonly generation: number };

const HOME: Route = Object.freeze({ kind: "home" });
const SHARED_MTA: Route = Object.freeze({ kind: "shared", handOver: false });
const SHARED_AGILE: Route = Object.freeze({ kind: "shared", handOver: true });

/**
 * Decides whether the arguments of a call can travel to the executor: COM
 * objects among them must be usable there (installed by the automation
 * layer, which knows VARIANT, ByRef and named arguments).
 */
export type ArgumentInspector = (args: readonly unknown[], homeMultithreaded: boolean) => boolean;

let inspectArguments: ArgumentInspector = () => true;

export function setArgumentInspector(inspector: ArgumentInspector): void {
  inspectArguments = inspector;
}

/**
 * Threading model of an object implemented in JavaScript at `address`
 * (`COMObject`), `null` for other objects (installed by the server layer).
 */
export type LocalThreading = (address: bigint) => "apartment" | "free" | null;

let localThreading: LocalThreading = () => null;

export function setLocalThreading(resolver: LocalThreading): void {
  localThreading = resolver;
}

// The executor's references of `proxy` routes, released if their pointer is forgotten.
const executorReferences = new FinalizationRegistry<{ readonly address: bigint; readonly generation: number }>(({ address, generation }) => {
  try {
    releaseOnExecutor(address, generation);
  } catch {
    // Nothing can be reported from a finalizer.
  }
});

let invoke: (pointer: COMPointer, slot: number, args: unknown[]) => unknown;
let invokeAsync: (pointer: COMPointer, slot: number, args: unknown[], options: AsyncInvokeOptions) => Promise<unknown>;
let shared: (pointer: COMPointer) => boolean;
let createView: (pointer: COMPointer) => COMPointer;

/** Call vtable slot `slot` of `pointer` synchronously, on the home thread (used by accessors and internal code). */
export function invokeSlot(pointer: COMPointer, slot: number, args: unknown[]): unknown {
  return invoke(pointer, slot, args);
}

export interface AsyncInvokeOptions {
  /** Further locations of interface results to hand back to a single-threaded home (e.g. raw VARIANT outputs). */
  readonly results?: readonly ResultLocation[];
  /** JavaScript values behind already converted arguments, checked instead of `args` (e.g. the arguments of `IDispatch::Invoke`). */
  readonly inspect?: readonly unknown[];
}

/**
 * Call vtable slot `slot` of `pointer` on the COM executor (used by the
 * generated interface methods). When the object or the arguments cannot
 * travel to the executor, the call runs synchronously on the home thread.
 */
export function invokeSlotAsync(pointer: COMPointer, slot: number, args: unknown[], options: AsyncInvokeOptions = {}): Promise<unknown> {
  return invokeAsync(pointer, slot, args, options);
}

/** Whether a home pointer is usable as it is on the executor (agile, or both in the MTA). */
export function sharedWithExecutor(pointer: COMPointer): boolean {
  return shared(pointer);
}

/**
 * A view of `pointer` whose interface methods run synchronously on the home
 * thread, sharing its reference (releasing either releases both). For event
 * handlers, COM servers and code that needs the lowest latency.
 */
export function sync<I extends InterfaceDef<any>>(pointer: COMPointer<I> | InterfacePointer<I>): SyncPointer<I> {
  if (!(pointer instanceof COMPointer)) throw new TypeError("sync() expects a COM pointer");
  return createView(pointer) as SyncPointer<I>;
}

// Set while constructing a synchronous view: the new instance shares this pointer's state.
let viewOf: COMPointer | null = null;

/**
 * An owned reference to a COM interface. Python equivalent: `POINTER(IUnknown)`.
 *
 * Each instance owns exactly one reference, released by {@link COMPointer.Release}
 * or `using`. A `FinalizationRegistry` releases forgotten pointers as a safety
 * net, never after the apartment has been uninitialized.
 *
 * Interface methods are installed by {@link COMInterface} on per-interface
 * subclasses: they return promises and run on the COM executor; properties
 * stay synchronous. Obtain pointers from `CoCreateInstance`, `QueryInterface`
 * or {@link COMPointer.fromAddress}.
 */
export class COMPointer<I extends InterfaceDef = InterfaceDef> implements Disposable {
  readonly interface: I;
  #reference: Reference;
  #functions: (Callable | undefined)[] = [];
  #state: { route: Route | Promise<Route> | null; inflight: number; view: COMPointer | null; entries: Entries | null } = { route: null, inflight: 0, view: null, entries: null };
  // A view keeps its pointer (the owner tracked for release) alive.
  #owner: COMPointer | null = null;

  static {
    invoke = (pointer, slot, args) => pointer.#invoke(slot, args);
    invokeAsync = (pointer, slot, args, options) => pointer.#invokeAsync(slot, args, options);
    shared = (pointer) => pointer.#shared();
    createView = (pointer) => {
      if (pointer.#state.view) return pointer.#state.view;
      viewOf = pointer;
      try {
        const SyncClass = (pointer.interface as InterfaceDef<any>).SyncPointer;
        return (pointer.#state.view = new SyncClass(pointer.#reference.address, pointer.interface));
      } finally {
        viewOf = null;
      }
    };
  }

  constructor(address: bigint, iface: I) {
    if (typeof address !== "bigint" || address === 0n) throw new TypeError("COMPointer requires a non-NULL address");
    this.interface = iface;
    if (viewOf) {
      // A synchronous view: same reference, same state.
      this.#reference = viewOf.#reference;
      this.#functions = viewOf.#functions;
      this.#state = viewOf.#state;
      this.#owner = viewOf;
      return;
    }
    this.#reference = track(address, this);
  }

  /** Wrap a raw interface pointer. */
  static fromAddress<T extends InterfaceDef>(address: bigint, iface: T, options: FromAddressOptions = {}): InterfacePointer<T> {
    if (typeof address !== "bigint" || address === 0n) throw new TypeError("fromAddress requires a non-NULL address");
    const pointer = new iface.Pointer(address, iface) as InterfacePointer<T>;
    if (options.addRef) pointer.AddRef();
    return pointer;
  }

  #live(): Reference {
    if (this.#reference.released) throw new Error(`${this.interface.name} pointer has been released`);
    return this.#reference;
  }

  #slot(slot: number) {
    const binding = this.interface.slots[slot];
    if (!binding) throw new RangeError(`${this.interface.name} has no vtable slot ${slot}`);
    return binding;
  }

  #invoke(slot: number, args: unknown[]): unknown {
    const { address } = this.#live();
    const binding = this.#slot(slot);
    let fn = this.#functions[slot];
    if (fn === undefined) fn = this.#functions[slot] = binding.bind(vtableEntry(address, slot));
    return binding.call(fn, address, args);
  }

  /** The route for asynchronous calls, computed once (a `proxy` route again for a new executor). */
  #route(): Route | Promise<Route> {
    const cached = this.#state.route;
    if (cached !== null && (cached instanceof Promise || cached.kind !== "proxy" || executorAlive(cached.generation))) return cached;
    const address = this.#live().address;
    const local = localThreading(address);
    const home = homeApartment();
    let route: Route | Promise<Route>;
    // Objects implemented in JavaScript run their callbacks on the home thread
    // unless free-threaded; without COM on the home thread, calls stay there.
    if (local === "apartment" || home === "none") route = HOME;
    else if (home === "mta") route = SHARED_MTA;
    else if (local === "free") route = SHARED_AGILE;
    // A proxy answers IClientSecurity locally; QueryInterface(IAgileObject) would go to the server.
    else if (isProxy(address)) route = this.#proxyRoute(address);
    else if (isAgile(address)) route = SHARED_AGILE;
    else route = HOME;
    this.#state.entries = null;
    this.#state.route = route;
    return route;
  }

  #proxyRoute(address: bigint): Promise<Route> {
    return referenceForExecutor(address, this.interface.iid).then(
      (executorAddress) => {
        if (!executorAddress) return (this.#state.route = HOME);
        const route: Route = { kind: "proxy", address: executorAddress, generation: executorGeneration() };
        if (this.#reference.released) {
          releaseOnExecutor(executorAddress, route.generation);
          return (this.#state.route = HOME);
        }
        executorReferences.register(this, { address: executorAddress, generation: route.generation }, this.#state);
        return (this.#state.route = route);
      },
      () => (this.#state.route = HOME),
    );
  }

  #shared(): boolean {
    const route = this.#state.route;
    if (route !== null && !(route instanceof Promise)) return route.kind === "shared";
    if (route !== null) return false;
    const address = this.#live().address;
    const local = localThreading(address);
    const home = homeApartment();
    if (local === "apartment" || home === "none") return false;
    if (home === "mta") {
      this.#state.route = SHARED_MTA;
      return true;
    }
    if (local === "free" || (!isProxy(address) && isAgile(address))) {
      this.#state.route = SHARED_AGILE;
      return true;
    }
    return false;
  }

  #invokeAsync(slot: number, args: unknown[], options: AsyncInvokeOptions): Promise<unknown> {
    return operation((current) => {
      try {
        this.#live();
        const binding = this.#slot(slot);
        const route = this.#route();
        if (route instanceof Promise) return route.then((settled) => this.#dispatch(settled, slot, binding, args, options, current));
        return this.#dispatch(route, slot, binding, args, options, current);
      } catch (error) {
        return Promise.reject(error);
      }
    });
  }

  #dispatch(route: Route, slot: number, binding: Slot, args: unknown[], options: AsyncInvokeOptions, current: Operation): Promise<unknown> {
    const reference = this.#live();
    // A reference of an executor stopped meanwhile (finalize): route again on the current one.
    if (route.kind === "proxy" && !executorAlive(route.generation)) {
      if (this.#state.route === route) {
        this.#state.route = null;
        this.#state.entries = null;
      }
      const fresh = this.#route();
      return fresh instanceof Promise ? fresh.then((settled) => this.#dispatch(settled, slot, binding, args, options, current)) : this.#dispatch(fresh, slot, binding, args, options, current);
    }
    const { results, inspect } = options;
    if (route.kind === "home" || !inspectArguments(inspect ?? args, route.kind === "shared" && !route.handOver)) {
      // Called where the object lives: on the home thread, synchronously.
      try {
        return Promise.resolve(this.#invoke(slot, args));
      } catch (error) {
        return Promise.reject(error);
      }
    }
    const target = route.kind === "proxy" ? route.address : reference.address;
    const handOverResults = route.kind === "proxy" || route.handOver;
    // The vtable of an address does not change: read each entry once.
    let entries = this.#state.entries;
    if (entries?.address !== target) entries = this.#state.entries = { address: target, slots: [] };
    const entry = (entries.slots[slot] ??= vtableEntry(target, slot));
    this.#state.inflight++;
    return binding.callAsync(entry, target, args, { handOver: handOverResults, results, operation: current }).finally(() => {
      this.#state.inflight--;
    });
  }

  /** Interface pointer value; throws once released. */
  get address(): bigint {
    return this.#live().address;
  }

  /** Passed as the pointer value to node-ctypes functions (Python parity). */
  get _as_parameter_(): bigint {
    return this.address;
  }

  get released(): boolean {
    return this.#reference.released;
  }

  /** Call a vtable method by name, synchronously, bypassing the typed accessors (e.g. the raw `QueryInterface`). */
  callMethod(name: string, ...args: unknown[]): unknown {
    const slot = this.interface.methodIndex.get(name);
    if (slot === undefined) throw new TypeError(`${this.interface.name} has no method ${name}`);
    return this.#invoke(slot, args);
  }

  /** New owned pointer for `iface`; throws {@link COMError} (`E_NOINTERFACE`) when unsupported. */
  QueryInterface<T extends InterfaceDef>(iface: T): InterfacePointer<T> {
    const out = new ctypes.c_void_p();
    this.#invoke(0, [iface.iid._buffer, ctypes.byref(out)]);
    if (!out.value) throw new COMError(E_NOINTERFACE, { method: `${this.interface.name}::QueryInterface` });
    return new iface.Pointer(out.value, iface) as InterfacePointer<T>;
  }

  /** Like {@link QueryInterface}, but `null` when the interface is not supported. */
  tryQueryInterface<T extends InterfaceDef>(iface: T): InterfacePointer<T> | null {
    const [hr, address] = rawQueryInterface(this.#live().address, iface.iid);
    if (hr === E_NOINTERFACE) return null;
    if (hr < 0 || !address) throw new COMError(hr, { method: `${this.interface.name}::QueryInterface` });
    return new iface.Pointer(address, iface) as InterfacePointer<T>;
  }

  /**
   * Raw `AddRef`, returning the new count. The extra reference is not owned
   * by this pointer: balance it with `callMethod("Release")` or hand it over.
   */
  AddRef(): number {
    return this.#invoke(1, []) as number;
  }

  /** Forget the executor's reference of a `proxy` route. */
  #dropExecutorReference(): void {
    const route = this.#state.route;
    if (route !== null && !(route instanceof Promise) && route.kind === "proxy") {
      executorReferences.unregister(this.#state);
      releaseOnExecutor(route.address, route.generation);
    }
    this.#state.route = null;
    this.#state.entries = null;
  }

  /**
   * Release the owned reference and invalidate this pointer. Returns the new
   * count, or 0 while asynchronous calls on the same address are still
   * running: the release then follows them on the executor.
   */
  Release(): number {
    const reference = this.#live();
    const route = this.#state.route;
    untrack(reference);
    this.#dropExecutorReference();
    if (this.#state.inflight > 0 && route !== null && !(route instanceof Promise) && route.kind === "shared") {
      releaseOnExecutor(reference.address, executorGeneration());
      return 0;
    }
    return rawRelease(reference.address);
  }

  /** COM identity: both pointers reach the same `IUnknown`. */
  isSameObject(other: COMPointer): boolean {
    if (!(other instanceof COMPointer)) throw new TypeError("isSameObject expects a COMPointer");
    const identity = (pointer: COMPointer) => {
      const [hr, address] = rawQueryInterface(pointer.address, IID_IUnknown);
      if (hr < 0 || !address) throw new COMError(hr, { method: "IUnknown::QueryInterface" });
      rawRelease(address);
      return address;
    };
    return identity(this) === identity(other);
  }

  /** Give up ownership without releasing: the caller becomes responsible for the reference. */
  detach(): bigint {
    const reference = this.#live();
    untrack(reference);
    this.#dropExecutorReference();
    return reference.address;
  }

  [Symbol.dispose](): void {
    if (!this.#reference.released) this.Release();
  }

  toString(): string {
    return this.#reference.released ? `<POINTER(${this.interface.name}) released>` : `<POINTER(${this.interface.name}) ${formatAddress(this.#reference.address)}>`;
  }

  [Symbol.for("nodejs.util.inspect.custom")](): string {
    return this.toString();
  }
}

function formatAddress(address: bigint): string {
  return `0x${address.toString(16).toUpperCase().padStart(16, "0")}`;
}
