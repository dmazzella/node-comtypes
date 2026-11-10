import * as ctypes from "node-ctypes";
import { COMError } from "./errors.ts";
import { errorInfoFor } from "./errorinfo.ts";
import { GUID } from "./guid.ts";
import { E_NOINTERFACE } from "./hresult.ts";
import type { InterfaceDef, InterfacePointer, MethodSpec } from "./interface.ts";
import { functionCache, rawQueryInterface, rawRelease, vtableEntry } from "./low-level/vtable.ts";
import { track, untrack } from "./tracking.ts";
import type { Reference } from "./tracking.ts";

type Callable = (...args: any[]) => any;

/** A vtable slot of an interface: binds the function found at a given address. */
export interface Slot {
  readonly spec: MethodSpec;
  /** `Interface::Method`, used in error messages. */
  readonly label: string;
  bind(address: bigint): Callable;
  /** Call the bound function: raw arguments, or converted by the spec's marshal (COMMETHOD). */
  call(fn: Callable, address: bigint, args: readonly unknown[]): unknown;
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
  const cache = functionCache(() => ctypes.WINFUNCTYPE(hresult ? ctypes.c_int32 : spec.restype, ctypes.c_void_p, ...spec.argtypes)) as (address: bigint) => Callable;
  const errcheck = (result: number, _fn: unknown, args: unknown[]) => {
    if (result < 0) throw new COMError(result, { method: label, details: errorInfo ? errorInfoFor(args[0] as bigint, owner.iid) : null });
    return result;
  };
  const { marshal } = spec;
  return {
    spec,
    label,
    call: marshal ? (fn, address, args) => marshal(fn, address, args, owner.definition!) : (fn, address, args) => fn(address, ...args),
    bind(address) {
      const fn = cache(address) as Callable & { errcheck: ctypes.ErrcheckCallback | null };
      if (hresult && fn.errcheck !== errcheck) fn.errcheck = errcheck;
      return fn;
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

let invoke: (pointer: COMPointer, slot: number, args: unknown[]) => unknown;

/** Call vtable slot `slot` of `pointer` (used by the generated interface methods). */
export function invokeSlot(pointer: COMPointer, slot: number, args: unknown[]): unknown {
  return invoke(pointer, slot, args);
}

/**
 * An owned reference to a COM interface. Python equivalent: `POINTER(IUnknown)`.
 *
 * Each instance owns exactly one reference, released by {@link COMPointer.Release}
 * or `using`. A `FinalizationRegistry` releases forgotten pointers as a safety
 * net, never after the apartment has been uninitialized.
 *
 * Interface methods are installed by {@link COMInterface} on per-interface
 * subclasses; obtain pointers from `CoCreateInstance`, `QueryInterface` or
 * {@link COMPointer.fromAddress}.
 */
export class COMPointer<I extends InterfaceDef = InterfaceDef> implements Disposable {
  readonly interface: I;
  #reference: Reference;
  #functions: (Callable | undefined)[] = [];

  static {
    invoke = (pointer, slot, args) => pointer.#invoke(slot, args);
  }

  constructor(address: bigint, iface: I) {
    if (typeof address !== "bigint" || address === 0n) throw new TypeError("COMPointer requires a non-NULL address");
    this.interface = iface;
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

  #invoke(slot: number, args: unknown[]): unknown {
    const { address } = this.#live();
    const binding = this.interface.slots[slot];
    if (!binding) throw new RangeError(`${this.interface.name} has no vtable slot ${slot}`);
    let fn = this.#functions[slot];
    if (fn === undefined) fn = this.#functions[slot] = binding.bind(vtableEntry(address, slot));
    return binding.call(fn, address, args);
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

  /** Call a vtable method by name, bypassing the typed accessors (e.g. the raw `QueryInterface`). */
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

  /** Release the owned reference and invalidate this pointer. Returns the new count. */
  Release(): number {
    const reference = this.#live();
    untrack(reference);
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
