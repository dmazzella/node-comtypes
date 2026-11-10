import * as ctypes from "node-ctypes";
import { getIDsOfNames, invoke } from "../automation/dispatch.ts";
import { DISPATCH_METHOD } from "../automation/constants.ts";
import type { BorrowedReference } from "../automation/variant.ts";
import { VARIANT_SIZE, comReference, writeVariant } from "../automation/variant.ts";
import { classifyParamType } from "../commethod.ts";
import type { ParamInfo, ParamKind } from "../commethod.ts";
import { COMError } from "../errors.ts";
import { GUID, toGUID } from "../guid.ts";
import type { GUIDLike } from "../guid.ts";
import { DISP_E_BADINDEX, TYPE_E_ELEMENTNOTFOUND, CLASS_E_NOAGGREGATION, CONNECT_E_CANNOTCONNECT, CONNECT_E_NOCONNECTION, E_FAIL, E_NOINTERFACE, E_NOTIMPL, E_POINTER, E_UNEXPECTED, S_FALSE, S_OK } from "../hresult.ts";
import { derivesFrom } from "../interface.ts";
import type { InterfaceDef, InterfacePointer, MethodSpec } from "../interface.ts";
import { IClassFactory } from "../interfaces/classfactory.ts";
import { IConnectionPoint, IConnectionPointContainer, IProvideClassInfo2 } from "../interfaces/connectionpoint.ts";
import { IDispatch, IEnumVARIANT } from "../interfaces/dispatch.ts";
import { ISupportErrorInfo } from "../interfaces/errorinfo.ts";
import { IUnknown } from "../interfaces/unknown.ts";
import { POINTER_SIZE, readMemory, writeMemory } from "../low-level/native.ts";
import { COMPointer } from "../pointer.ts";
import { ReturnHRESULT, dispatchGetIDsOfNames, dispatchInvoke, defaultIdsOfNames, defaultInvoke } from "./automation.ts";
import type { InvokeArguments } from "./automation.ts";
import { setErrorInfo } from "./errorinfo.ts";
import { typeLibLocationOf } from "../typelib/coclass.ts";
import type { TypeLibLocation } from "../typelib/coclass.ts";
import { loadTypeLib } from "../typelib/load.ts";
import { ITypeInfo, ITypeLib } from "../interfaces/typeinfo.ts";
import { callbackArgtype, decodeIndirect, decodeInput, encodeOutput, resetOutput } from "./marshal.ts";

/**
 * `apartment`: vtable callbacks run on the calling thread, which must be the
 * JavaScript thread (STA objects, the default). `free`: thread-safe callbacks;
 * calls from other threads are queued to the JavaScript thread and wait for it,
 * so they deadlock while that thread is blocked in a synchronous COM call.
 */
export type Threading = "apartment" | "free";

type Callable = (...args: any[]) => any;

interface ClassInfo {
  /** Implemented interfaces: one native block (vtable pointer) each. */
  readonly interfaces: readonly InterfaceDef<any>[];
  /** IID (including base interfaces) -> index in `interfaces`. */
  readonly lookup: ReadonlyMap<string, number>;
  readonly threading: Threading;
}

interface State {
  references: number;
  readonly info: ClassInfo;
  /** Native blocks: each holds the address of an interface vtable. */
  readonly blocks: readonly Buffer[];
  readonly addresses: readonly bigint[];
  connectionPoints: Map<string, ConnectionPoint> | null;
}

// Block address -> object. Weak: the JavaScript object owns its blocks; while
// COM holds references it is kept in `alive`.
const instances = new Map<bigint, { readonly object: WeakRef<COMObject>; readonly index: number }>();
const forget = new FinalizationRegistry<readonly bigint[]>((addresses) => {
  for (const address of addresses) instances.delete(address);
});
const alive = new Set<COMObject>();
const classInfos = new WeakMap<Function, ClassInfo>();

function interfaceChain(iface: InterfaceDef<any>): InterfaceDef<any>[] {
  const chain: InterfaceDef<any>[] = [];
  for (let current: InterfaceDef<any> | null = iface; current; current = current.base) chain.push(current);
  return chain;
}

function classInfoOf(type: typeof COMObject): ClassInfo {
  let info = classInfos.get(type);
  if (info) return info;
  const listed = type._com_interfaces_ ?? [];
  if (!listed.length) throw new TypeError(`${type.name}: _com_interfaces_ is empty`);
  for (const iface of listed) if (!derivesFrom(iface, IUnknown.iid)) throw new TypeError(`${type.name}: ${iface.name} does not derive from IUnknown`);
  const interfaces = [...listed];
  const outgoing = (type._outgoing_interfaces_ ?? []).length > 0;
  // IProvideClassInfo2 for connection points, and to describe a coclass of a type library.
  const classInfo = outgoing || (type._reg_clsid_ !== undefined && classTypeLib(type) !== null);
  const implicit = [ISupportErrorInfo, ...(outgoing ? [IConnectionPointContainer] : []), ...(classInfo ? [IProvideClassInfo2] : [])];
  for (const iface of implicit) if (!interfaces.some((candidate) => derivesFrom(candidate, iface.iid))) interfaces.push(iface);
  const lookup = new Map<string, number>();
  interfaces.forEach((iface, index) => {
    for (const member of interfaceChain(iface)) if (!lookup.has(member.iid.toString())) lookup.set(member.iid.toString(), index);
  });
  info = Object.freeze({ interfaces: Object.freeze(interfaces), lookup, threading: type._threading_ ?? "apartment" });
  classInfos.set(type, info);
  return info;
}

/** The interface of `iface`'s chain that declares vtable slot `slot`. */
function declaringInterface(iface: InterfaceDef<any>, slot: number): InterfaceDef<any> {
  let current = iface;
  while (current.base && slot < current.base.allMethods.length) current = current.base;
  return current;
}

// ── vtables ───────────────────────────────────────────────────────────

interface Vtable {
  readonly table: Buffer;
  readonly callbacks: readonly ctypes.CallbackWrapper[];
}

const vtables = new WeakMap<InterfaceDef<any>, Map<Threading, Vtable>>();
let dispatchSlot: (self: bigint | null, iface: InterfaceDef<any>, slot: number, args: unknown[]) => unknown;
let queryReference: (object: COMObject, iid: GUIDLike) => bigint | null;

function vtableFor(iface: InterfaceDef<any>, threading: Threading): bigint {
  let byThreading = vtables.get(iface);
  if (!byThreading) vtables.set(iface, (byThreading = new Map()));
  let vtable = byThreading.get(threading);
  if (!vtable) {
    const callbacks = iface.allMethods.map((spec, slot) => {
      const restype = slot === 1 || slot === 2 ? ctypes.c_uint32 : spec.restype === ctypes.HRESULT ? ctypes.c_int32 : spec.restype;
      const argtypes = [ctypes.c_void_p, ...spec.argtypes.map(callbackArgtype)];
      const handler = (self: bigint | null, ...args: unknown[]) => dispatchSlot(self, iface, slot, args);
      if (threading === "free") return ctypes.threadSafeCallback(handler as never, restype as never, argtypes, "stdcall").unref();
      return ctypes.WINFUNCTYPE(restype, ...argtypes)(handler as never) as unknown as ctypes.CallbackWrapper;
    });
    const table = Buffer.alloc(POINTER_SIZE * callbacks.length);
    callbacks.forEach((callback, index) => ctypes.writeValue(table, ctypes.c_void_p, callback.pointer, index * POINTER_SIZE));
    byThreading.set(threading, (vtable = { table, callbacks }));
  }
  return ctypes.addressof(vtable.table);
}

// ── implementations ───────────────────────────────────────────────────

type ServerSpec = MethodSpec & { readonly params?: readonly ParamInfo[] };

const paramKinds = new WeakMap<MethodSpec, readonly ParamKind[]>();

function kindsOf(spec: ServerSpec): readonly ParamKind[] {
  let kinds = paramKinds.get(spec);
  if (!kinds) paramKinds.set(spec, (kinds = (spec.params ?? []).map((param) => classifyParamType(param.type))));
  return kinds;
}

/** The JavaScript implementation of a vtable method: a method named as the spec, or a property for accessors. */
function implementationOf(object: COMObject, spec: MethodSpec): Callable | null {
  const record = object as unknown as Record<string, unknown>;
  const direct = record[spec.name];
  if (typeof direct === "function") return direct as Callable;
  if (spec.kind === "method" || !(spec.property in object)) return null;
  const property = spec.property;
  if (spec.kind === "propget") {
    const value = record[property];
    return typeof value === "function" ? (value as Callable) : () => record[property];
  }
  return (...args: unknown[]) => {
    if (args.length !== 1) throw new COMError(E_NOTIMPL, { text: `${property}: parameterized puts require _set_${property}` });
    record[property] = args[0];
  };
}

function releaseArguments(values: readonly unknown[]): void {
  for (const value of values) if (value instanceof COMPointer && !value.released) value.Release();
}

function callCommethod(object: COMObject, implementation: Callable, spec: ServerSpec, declaring: InterfaceDef<any>, args: unknown[]): number {
  const params = spec.params ?? [];
  const kinds = kindsOf(spec);
  const inputs: unknown[] = [];
  const outputs: Array<{ kind: ParamKind; address: bigint; inout: boolean } | null> = [];
  try {
    params.forEach((param, index) => {
      const native = args[index];
      if (param.lcid) return;
      if (!param.output) {
        inputs.push(decodeInput(kinds[index], native, declaring));
        return;
      }
      const address = native as bigint | null;
      if (!address) {
        outputs.push(null);
        if (param.input) inputs.push(undefined);
        return;
      }
      if (param.input) inputs.push(decodeIndirect(kinds[index], address, declaring));
      else resetOutput(kinds[index], address);
      outputs.push({ kind: kinds[index], address, inout: param.input });
    });
    const result = implementation.apply(object, inputs);
    if (!outputs.length) return typeof result === "number" ? result : S_OK;
    const values = outputs.length === 1 ? [result] : result;
    if (!Array.isArray(values) || values.length !== outputs.length) throw new TypeError(`${spec.name} must return ${outputs.length} values`);
    outputs.forEach((output, index) => {
      if (output) encodeOutput(output.kind, output.address, values[index], output.inout, declaring);
    });
    return S_OK;
  } finally {
    releaseArguments(inputs);
  }
}

/** Report an exception thrown by an implementation: `IErrorInfo` and an HRESULT. */
function reportError(object: COMObject, declaring: InterfaceDef<any>, error: unknown): number {
  if (error instanceof ReturnHRESULT && error.hresult >= 0) return error.hresult;
  const hresult = error instanceof COMError && error.hresult < 0 ? error.hresult : E_FAIL;
  const type = object.constructor as typeof COMObject;
  try {
    setErrorInfo({ description: error instanceof COMError ? (error.details?.description ?? error.text) : error instanceof Error ? error.message : String(error), source: type._reg_progid_ ?? type.name, guid: declaring.iid });
  } catch {
    // Reporting must not fail the call further.
  }
  if (!(error instanceof ReturnHRESULT)) type.onError?.(error, object);
  return hresult;
}

// ── COMObject ─────────────────────────────────────────────────────────

/**
 * A COM object implemented in JavaScript. Python equivalent: `comtypes.COMObject`.
 *
 * ```ts
 * class Thing extends COMObject {
 *   static _com_interfaces_ = [IThing];
 *   Name = "thing";                 // propget/propput Name
 *   Find(what: string) { ... }      // COMMETHOD Find: inputs in, outputs returned
 * }
 * using thing = new Thing().QueryInterface(IThing);
 * ```
 *
 * - Methods are found by the vtable name (`Find`, `_get_Name`, `_set_Name`);
 *   accessors fall back to a JavaScript property of the same name.
 * - COMMETHOD specs are converted (`in` arguments, returned `out` values; one
 *   value, or an array for several); STDMETHOD specs receive the raw arguments
 *   and may return an HRESULT. Interface arguments are lent for the call:
 *   `QueryInterface` (or `AddRef`) them to keep them.
 * - Exceptions become failure HRESULTs (`COMError.hresult`, else `E_FAIL`)
 *   with `IErrorInfo`; `ISupportErrorInfo` is implemented automatically.
 * - Interfaces deriving from `IDispatch` get a late-bound implementation
 *   (members by name, `_NewEnum` for iterables); see `server/automation.ts`.
 * - `_outgoing_interfaces_` adds connection points; `fireEvent` raises events.
 *
 * Reference counting is native; while COM holds references the object is
 * kept alive. `_final_release_()` runs when the count drops to zero.
 */
export class COMObject {
  /** Implemented interfaces. Python: `_com_interfaces_`. */
  static _com_interfaces_: readonly InterfaceDef<any>[] = [];
  /** Event interfaces raised through connection points. Python: `_outgoing_interfaces_`. */
  static _outgoing_interfaces_: readonly InterfaceDef<any>[] = [];
  static _threading_: Threading = "apartment";
  /** Registration data for class factories and LocalServer32. */
  static _reg_clsid_?: GUIDLike;
  static _reg_progid_?: string;
  static _reg_desc_?: string;
  /** Version-independent ProgID (`CLSID\{clsid}\VersionIndependentProgID`). Python: `_reg_novers_progid_`. */
  static _reg_novers_progid_?: string;
  /**
   * Type library of the class: LIBID, major and minor version. Default: the
   * library the first of `_com_interfaces_` was built from (`GetModule`,
   * generated modules). Python: `_reg_typelib_`.
   */
  static _reg_typelib_?: readonly [libid: GUIDLike, major: number, minor: number];
  /** Type library file registered with the class (default: the file `GetModule` loaded). Python: `_typelib_path_`. */
  static _typelib_path_?: string;
  /** Member invoked for `DISPID_VALUE` through `IDispatch`. */
  static _com_default_member_?: string;
  /** Called with exceptions thrown by implementations (after reporting them to the caller). */
  static onError?: (error: unknown, object: COMObject) => void;

  #state: State;

  static {
    queryReference = (object, iid) => {
      const address = object.#blockFor(iid);
      if (address === null) return null;
      object.#addRef();
      return address;
    };
    dispatchSlot = (self, iface, slot, args) => {
      const entry = self ? instances.get(self) : undefined;
      const object = entry?.object.deref();
      if (!object) return slot === 1 || slot === 2 ? 0 : E_UNEXPECTED;
      switch (slot) {
        case 0:
          return object.#queryInterface(args[0] as bigint | null, args[1] as bigint | null);
        case 1:
          return object.#addRef();
        case 2:
          return object.#release();
        default:
          return object.#call(iface, slot, args);
      }
    };
  }

  constructor() {
    const info = classInfoOf(new.target);
    const blocks = info.interfaces.map((iface) => {
      const block = Buffer.alloc(POINTER_SIZE);
      ctypes.writeValue(block, ctypes.c_void_p, vtableFor(iface, info.threading));
      return block;
    });
    const addresses = blocks.map((block) => ctypes.addressof(block));
    this.#state = { references: 0, info, blocks, addresses, connectionPoints: null };
    const reference = new WeakRef<COMObject>(this);
    addresses.forEach((address, index) => instances.set(address, { object: reference, index }));
    forget.register(this, Object.freeze([...addresses]));
  }

  #addRef(): number {
    if (this.#state.references++ === 0) alive.add(this);
    return this.#state.references;
  }

  #release(): number {
    const state = this.#state;
    if (state.references === 0) return 0;
    if (--state.references === 0) {
      alive.delete(this);
      try {
        this._final_release_();
      } catch (error) {
        (this.constructor as typeof COMObject).onError?.(error, this);
      }
    }
    return state.references;
  }

  /** Address of the block implementing `iid`, or `null`. */
  #blockFor(iid: GUIDLike): bigint | null {
    const index = this.#state.info.lookup.get(toGUID(iid).toString());
    return index === undefined ? null : this.#state.addresses[index];
  }

  #queryInterface(riid: bigint | null, out: bigint | null): number {
    if (!out) return E_POINTER;
    ctypes.writeValue(out, ctypes.c_void_p, 0n);
    if (!riid) return E_POINTER;
    const address = this.#blockFor(new GUID(new Uint8Array(readMemory(riid, 16))));
    if (address === null) return E_NOINTERFACE;
    this.#addRef();
    ctypes.writeValue(out, ctypes.c_void_p, address);
    return S_OK;
  }

  #call(iface: InterfaceDef<any>, slot: number, args: unknown[]): unknown {
    const spec = iface.allMethods[slot] as ServerSpec;
    const declaring = declaringInterface(iface, slot);
    const hresult = spec.restype === ctypes.HRESULT;
    try {
      const implementation = implementationOf(this, spec);
      if (!implementation) return hresult ? E_NOTIMPL : spec.restype === null ? undefined : 0;
      if (!spec.marshal) {
        const result = implementation.apply(this, args);
        if (!hresult) return result ?? (spec.restype === null ? undefined : 0);
        return typeof result === "number" ? result : S_OK;
      }
      return callCommethod(this, implementation, spec, declaring, args);
    } catch (error) {
      if (!hresult) {
        (this.constructor as typeof COMObject).onError?.(error, this);
        return spec.restype === null ? undefined : 0;
      }
      return reportError(this, declaring, error);
    }
  }

  /** Current native reference count. */
  get refCount(): number {
    return this.#state.references;
  }

  /** Interfaces this object answers `QueryInterface` for (listed and implicit ones). */
  get interfaces(): readonly InterfaceDef<any>[] {
    return this.#state.info.interfaces;
  }

  supportsInterface(iid: GUIDLike): boolean {
    return this.#blockFor(iid) !== null;
  }

  /** A new owned pointer to this object (`AddRef`); throws `E_NOINTERFACE`. */
  QueryInterface<T extends InterfaceDef<any>>(iface: T): InterfacePointer<T> {
    const address = this.#blockFor(iface.iid);
    if (address === null) throw new COMError(E_NOINTERFACE, { method: `${this.constructor.name}::QueryInterface(${iface.name})` });
    this.#addRef();
    return COMPointer.fromAddress(address, iface) as InterfacePointer<T>;
  }

  /** Native `AddRef`/`Release` for code managing references by hand. */
  AddRef(): number {
    return this.#addRef();
  }

  Release(): number {
    return this.#release();
  }

  /** Lends the first interface (`borrowReference`), e.g. to pass the object as an argument. */
  [comReference](): BorrowedReference {
    return { address: this.#state.addresses[0], interface: this.#state.info.interfaces[0] };
  }

  /** Called when the native reference count drops to zero. Python: `_final_release_`. */
  _final_release_(): void {}

  // ── ISupportErrorInfo ──
  InterfaceSupportsErrorInfo(riid: GUID | null): number {
    return riid && this.supportsInterface(riid) ? S_OK : S_FALSE;
  }

  // ── IDispatch (late binding over the JavaScript members) ──
  GetTypeInfoCount(pctinfo: bigint | null): number {
    if (!pctinfo) return E_POINTER;
    ctypes.writeValue(pctinfo, ctypes.c_uint32, dispatchTypeInfoSource(this) ? 1 : 0);
    return S_OK;
  }

  /** Type information of the first dispatch interface built from a type library (`GetModule`, generated modules). */
  GetTypeInfo(index: number, _lcid: number, ppTInfo: bigint | null): number {
    if (!ppTInfo) return E_POINTER;
    ctypes.writeValue(ppTInfo, ctypes.c_void_p, 0n);
    const source = index === 0 ? dispatchTypeInfoSource(this) : null;
    if (!source) return index === 0 ? E_NOTIMPL : DISP_E_BADINDEX;
    const typeInfo = typeInfoFor(source.iface, source.location);
    if (!typeInfo) return E_NOTIMPL;
    ctypes.writeValue(ppTInfo, ctypes.c_void_p, typeInfo);
    return S_OK;
  }

  GetIDsOfNames(_riid: bigint | null, rgszNames: bigint | null, cNames: number, _lcid: number, rgDispId: bigint | null): number {
    return dispatchGetIDsOfNames(this, rgszNames, cNames, rgDispId);
  }

  Invoke(dispIdMember: number, _riid: bigint | null, _lcid: number, wFlags: number, pDispParams: bigint | null, pVarResult: bigint | null, pExcepInfo: bigint | null, puArgErr: bigint | null): number {
    return dispatchInvoke(this, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
  }

  /** DISPIDs for a member name and its named arguments; override to customize late binding. */
  _ids_of_names_(names: readonly string[]): number[] {
    return defaultIdsOfNames(this, names);
  }

  /** Late-bound call; override to customize. Throw `COMError` with a `DISP_E_*` code for dispatch errors. */
  _invoke_(dispid: number, flags: number, args: InvokeArguments): unknown {
    return defaultInvoke(this, dispid, flags, args);
  }

  // ── connection points ──
  #connectionPoint(iid: GUIDLike): ConnectionPoint | null {
    const outgoing = (this.constructor as typeof COMObject)._outgoing_interfaces_ ?? [];
    const guid = toGUID(iid);
    const iface = outgoing.find((candidate) => candidate.iid.equals(guid));
    if (!iface) return null;
    const points = (this.#state.connectionPoints ??= new Map());
    let point = points.get(iface.iid.toString());
    if (!point) points.set(iface.iid.toString(), (point = new ConnectionPoint(this, iface)));
    return point;
  }

  FindConnectionPoint(riid: GUID | null): ConnectionPoint {
    const point = riid ? this.#connectionPoint(riid) : null;
    if (!point) throw new COMError(CONNECT_E_NOCONNECTION);
    return point;
  }

  EnumConnectionPoints(): never {
    throw new COMError(E_NOTIMPL);
  }

  /** Type information of the coclass `_reg_clsid_` in the class type library. */
  GetClassInfo(): InterfacePointer<typeof ITypeInfo> {
    const type = this.constructor as typeof COMObject;
    const location = classTypeLib(type);
    if (!type._reg_clsid_ || !location) throw new COMError(E_NOTIMPL);
    const address = typeInfoFor(toGUID(type._reg_clsid_), location);
    if (!address) throw new COMError(TYPE_E_ELEMENTNOTFOUND);
    return COMPointer.fromAddress(address, ITypeInfo);
  }

  GetGUID(kind: number): GUID {
    const outgoing = (this.constructor as typeof COMObject)._outgoing_interfaces_ ?? [];
    if (kind !== 1 || !outgoing.length) throw new COMError(E_NOTIMPL);
    return outgoing[0].iid;
  }

  /** Sinks connected for an outgoing interface. */
  connectedSinks(iface: InterfaceDef<any>): COMPointer[] {
    return this.#state.connectionPoints?.get(iface.iid.toString())?.sinks ?? [];
  }

  /**
   * Raise event `name` on every connected sink of the outgoing interfaces
   * declaring it: through `IDispatch::Invoke` for dispinterfaces (the spec
   * `dispid`, else `GetIDsOfNames`), otherwise through the vtable. Returns the
   * sinks' results; sink failures are thrown together after all calls.
   */
  fireEvent(name: string, ...args: unknown[]): unknown[] {
    const results: unknown[] = [];
    const errors: unknown[] = [];
    for (const iface of (this.constructor as typeof COMObject)._outgoing_interfaces_ ?? []) {
      const spec = [...iface.allMethods, ...iface.dispMethods].find((method) => method.property === name || method.name === name);
      if (!spec) continue;
      for (const sink of this.connectedSinks(iface)) {
        try {
          if (derivesFrom(iface, IDispatch.iid)) {
            const dispatch = sink as InterfacePointer<typeof IDispatch>;
            const dispid = spec.dispid ?? getIDsOfNames(dispatch, [spec.property])[0];
            results.push(invoke(dispatch, dispid, DISPATCH_METHOD, args, { name: spec.property }));
          } else {
            results.push(sink.callMethod(spec.name, ...args));
          }
        } catch (error) {
          errors.push(error);
        }
      }
    }
    if (errors.length === 1) throw errors[0];
    if (errors.length) throw new AggregateError(errors, `${errors.length} event sinks failed for ${name}`);
    return results;
  }
}

/** `QueryInterface` by IID: an owned interface pointer address, or `null`. */
export function queryObject(object: COMObject, iid: GUIDLike): bigint | null {
  return queryReference(object, iid);
}

function dispatchTypeInfoSource(object: COMObject): { iface: InterfaceDef<any>; location: TypeLibLocation } | null {
  for (const iface of object.interfaces) {
    if (!derivesFrom(iface, IDispatch.iid)) continue;
    for (let current: InterfaceDef<any> | null = iface; current; current = current.base) {
      const location = typeLibLocationOf(current);
      if (location) return { iface: current, location };
    }
  }
  return null;
}

const typeLibraries = new Map<string, COMPointer>();

/**
 * Type library of a class: `_reg_typelib_` (with `_typelib_path_`), else the
 * library of the first interface built from one.
 */
export function classTypeLib(type: typeof COMObject): TypeLibLocation | null {
  if (type._reg_typelib_) {
    const [libid, major, minor] = type._reg_typelib_;
    return { libid: toGUID(libid).toString(), major, minor, lcid: 0, ...(type._typelib_path_ ? { path: type._typelib_path_ } : {}) };
  }
  for (const iface of type._com_interfaces_ ?? []) {
    const location = typeLibLocationOf(iface);
    if (location) return type._typelib_path_ ? { ...location, path: type._typelib_path_ } : location;
  }
  return null;
}

/** An owned `ITypeInfo` address for the type `guid` (an interface or coclass), from its (cached) type library. */
function typeInfoFor(guid: InterfaceDef<any> | GUID, location: TypeLibLocation): bigint | null {
  const key = `${location.libid}:${location.major}.${location.minor}:${location.lcid}`;
  let library = typeLibraries.get(key) as InterfacePointer<typeof ITypeLib> | undefined;
  try {
    if (!library || library.released) {
      library = location.path ? loadTypeLib(location.path) : loadTypeLib({ libid: location.libid, major: location.major, minor: location.minor, lcid: location.lcid });
      typeLibraries.set(key, library);
    }
    const out = new ctypes.c_void_p();
    library.GetTypeInfoOfGuid(guid instanceof GUID ? guid : guid.iid, ctypes.byref(out));
    return out.value;
  } catch {
    return null;
  }
}

/** Objects kept alive by native references (all classes). */
export function getLiveObjectCount(predicate?: (object: COMObject) => boolean): number {
  if (!predicate) return alive.size;
  let count = 0;
  for (const object of alive) if (predicate(object)) count++;
  return count;
}

// ── standard implementations ──────────────────────────────────────────

/** `IEnumVARIANT` over a snapshot of values. */
export class VariantEnumerator extends COMObject {
  static _com_interfaces_ = [IEnumVARIANT];
  readonly items: readonly unknown[];
  #position: number;

  constructor(items: Iterable<unknown>, position = 0) {
    super();
    this.items = Object.freeze([...items]);
    this.#position = position;
  }

  Next(celt: number, rgVar: bigint | null, pCeltFetched: bigint | null): number {
    if (!rgVar) return E_POINTER;
    let fetched = 0;
    try {
      while (fetched < celt && this.#position < this.items.length) {
        const slot = Buffer.alloc(VARIANT_SIZE);
        writeVariant(slot, this.items[this.#position]);
        writeMemory(rgVar + BigInt(fetched * VARIANT_SIZE), slot);
        fetched++;
        this.#position++;
      }
    } finally {
      if (pCeltFetched) ctypes.writeValue(pCeltFetched, ctypes.c_uint32, fetched);
    }
    return fetched === celt ? S_OK : S_FALSE;
  }

  Skip(celt: number): number {
    const target = this.#position + celt;
    this.#position = Math.min(target, this.items.length);
    return target <= this.items.length ? S_OK : S_FALSE;
  }

  Reset(): number {
    this.#position = 0;
    return S_OK;
  }

  Clone(): VariantEnumerator {
    return new VariantEnumerator(this.items, this.#position);
  }
}

/** Connection point of a `COMObject` for one outgoing interface. */
export class ConnectionPoint extends COMObject {
  static _com_interfaces_ = [IConnectionPoint];
  readonly container: COMObject;
  readonly outgoing: InterfaceDef<any>;
  #sinks = new Map<number, COMPointer>();
  #nextCookie = 1;

  constructor(container: COMObject, outgoing: InterfaceDef<any>) {
    super();
    this.container = container;
    this.outgoing = outgoing;
  }

  get sinks(): COMPointer[] {
    return [...this.#sinks.values()];
  }

  GetConnectionInterface(): GUID {
    return this.outgoing.iid;
  }

  GetConnectionPointContainer(): COMObject {
    return this.container;
  }

  Advise(sink: COMPointer | null): number {
    const typed = sink?.tryQueryInterface(this.outgoing);
    if (!typed) throw new COMError(CONNECT_E_CANNOTCONNECT);
    const cookie = this.#nextCookie++;
    this.#sinks.set(cookie, typed);
    return cookie;
  }

  Unadvise(cookie: number): void {
    const sink = this.#sinks.get(cookie);
    if (!sink) throw new COMError(CONNECT_E_NOCONNECTION);
    this.#sinks.delete(cookie);
    sink.Release();
  }

  EnumConnections(): never {
    throw new COMError(E_NOTIMPL);
  }

  _final_release_(): void {
    // The container keeps its connection points; sinks stay until Unadvise.
  }
}

/** `IClassFactory` creating instances of a `COMObject` class. */
export class ClassFactory extends COMObject {
  static _com_interfaces_ = [IClassFactory];
  readonly type: new () => COMObject;
  locks = 0;

  constructor(type: new () => COMObject) {
    super();
    this.type = type;
  }

  CreateInstance(outer: bigint | null, riid: bigint | null, ppv: bigint | null): number {
    if (!ppv) return E_POINTER;
    ctypes.writeValue(ppv, ctypes.c_void_p, 0n);
    if (outer) return CLASS_E_NOAGGREGATION;
    if (!riid) return E_POINTER;
    const object = new this.type();
    const address = queryObject(object, new GUID(new Uint8Array(readMemory(riid, 16))));
    if (address === null) return E_NOINTERFACE;
    ctypes.writeValue(ppv, ctypes.c_void_p, address);
    return S_OK;
  }

  LockServer(lock: number): number {
    this.locks += lock ? 1 : -1;
    return S_OK;
  }
}
