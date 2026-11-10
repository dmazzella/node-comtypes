import * as ctypes from "node-ctypes";
import { COMError } from "../errors.ts";
import { DISP_E_BADPARAMCOUNT, DISP_E_MEMBERNOTFOUND, DISP_E_PARAMNOTOPTIONAL, DISP_E_TYPEMISMATCH, DISP_E_UNKNOWNNAME, E_INVALIDARG, S_OK } from "../hresult.ts";
import { derivesFrom } from "../interface.ts";
import type { InterfacePointer } from "../interface.ts";
import { IDispatch, IEnumVARIANT } from "../interfaces/dispatch.ts";
import { COMPointer } from "../pointer.ts";
import { DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPATCH_PROPERTYPUTREF, DISPID_NEWENUM, DISPID_VALUE } from "../automation/constants.ts";
import { getIDsOfNames, invoke } from "../automation/dispatch.ts";
import type { DispatchPointer } from "../automation/dispatch.ts";
import { dispatchTypeInfo, documentation, typeMembers } from "../automation/members.ts";
import type { MemberInfo } from "../automation/members.ts";
import { VARIANT, borrowReference, comPointer } from "../automation/variant.ts";
import type { ComObject } from "../commethod.ts";
import { checkHRESULT } from "../errors.ts";
import { rawQueryInterface } from "../low-level/vtable.ts";
import { NamedArguments } from "../automation/named.ts";

export { NamedArguments, named } from "../automation/named.ts";

/**
 * A late-bound automation object. Python equivalent: `comtypes.client.dynamic.Dispatch`.
 *
 * - `obj.Prop` reads a property; `obj.Prop = value` writes it (`PROPERTYPUTREF` for objects).
 * - `obj.Method(...args)` calls a method; parameterized properties are called too: `obj.Item(key)`.
 * - `obj(...args)` invokes the default member (`DISPID_VALUE`).
 * - `for (const item of obj)` enumerates through `_NewEnum`.
 * - `using obj = ...` / `release(obj)` releases the object.
 *
 * Member kinds come from the type information when available; otherwise a
 * read is attempted and a method caller is returned when the member is not a
 * readable property, as comtypes does.
 */
export interface DispatchObject extends Disposable, Iterable<any> {
  (...args: any[]): any;
  [member: string]: any;
  readonly [comPointer]: DispatchPointer;
}

interface State {
  readonly pointer: DispatchPointer;
  readonly dispids: Map<string, number>;
  /** `undefined`: not loaded yet; `null`: no type information. */
  members: ReadonlyMap<number, MemberInfo> | null | undefined;
  typeName: string | null | undefined;
}

const states = new WeakMap<object, State>();
const BAD_CONTEXT = new Set([DISP_E_MEMBERNOTFOUND, DISP_E_BADPARAMCOUNT, DISP_E_PARAMNOTOPTIONAL, DISP_E_TYPEMISMATCH, E_INVALIDARG]);
const inspect = Symbol.for("nodejs.util.inspect.custom");

const pointerStates = new WeakMap<COMPointer, State>();

/** State of a dispatch object, or of a typed pointer deriving from `IDispatch` (for the late-bound helpers). */
function state(object: unknown): State {
  if (object instanceof COMPointer && derivesFrom(object.interface, IDispatch.iid)) {
    let current = pointerStates.get(object);
    if (!current) pointerStates.set(object, (current = { pointer: object as DispatchPointer, dispids: new Map(), members: undefined, typeName: undefined }));
    return current;
  }
  const result = typeof object === "function" ? states.get(object) : undefined;
  if (!result) throw new TypeError("Expected a dispatch object");
  return result;
}

function loadTypeInfo(current: State): void {
  if (current.members !== undefined) return;
  using typeInfo = dispatchTypeInfo(current.pointer);
  if (!typeInfo) {
    current.members = null;
    current.typeName = null;
    return;
  }
  try {
    current.members = typeMembers(typeInfo);
    current.typeName = documentation(typeInfo).name;
  } catch {
    current.members = null;
    current.typeName = null;
  }
}

function dispidOf(current: State, name: string): number {
  const key = name.toLowerCase();
  let dispid = current.dispids.get(key);
  if (dispid === undefined) {
    dispid = getIDsOfNames(current.pointer, [name])[0];
    current.dispids.set(key, dispid);
  }
  return dispid;
}

function memberOf(current: State, dispid: number): MemberInfo | undefined {
  loadTypeInfo(current);
  return current.members?.get(dispid);
}

/** Wrap results: dispatch pointers become dynamic objects, recursively in arrays. */
function wrap(value: unknown): unknown {
  if (value instanceof COMPointer && derivesFrom(value.interface, IDispatch.iid)) return createProxy(value as DispatchPointer);
  if (Array.isArray(value)) return value.map(wrap);
  return value;
}

function call(current: State, name: string | null, dispid: number, flags: number, args: readonly unknown[]): unknown {
  let positional = args;
  let namedArguments: { dispids: number[]; values: unknown[] } | undefined;
  const last = args[args.length - 1];
  if (last instanceof NamedArguments) {
    positional = args.slice(0, -1);
    const names = Object.keys(last.values);
    if (names.length) {
      if (name === null) throw new TypeError("Named arguments require a member name");
      const ids = getIDsOfNames(current.pointer, [name, ...names]);
      namedArguments = { dispids: ids.slice(1), values: names.map((key) => last.values[key]) };
    }
  }
  return wrap(invoke(current.pointer, dispid, flags, positional, { named: namedArguments, name: name ?? undefined }));
}

function isObject(value: unknown): boolean {
  if (value instanceof COMPointer) return true;
  if (value !== null && (typeof value === "object" || typeof value === "function")) return (value as Record<symbol, unknown>)[comPointer] instanceof COMPointer;
  return false;
}

function getMember(current: State, name: string): unknown {
  const dispid = dispidOf(current, name);
  const member = memberOf(current, dispid);
  const caller = (...args: unknown[]) => call(current, name, dispid, DISPATCH_METHOD | DISPATCH_PROPERTYGET, args);
  if (member) {
    if (member.get && member.getArguments === 0) return call(current, name, dispid, DISPATCH_PROPERTYGET, []);
    if (member.method || member.get) return caller;
    return undefined;
  }
  try {
    return call(current, name, dispid, DISPATCH_PROPERTYGET, []);
  } catch (error) {
    if (error instanceof COMError && BAD_CONTEXT.has(error.hresult)) return caller;
    throw error;
  }
}

function setMember(current: State, name: string, value: unknown): void {
  const dispid = dispidOf(current, name);
  putMember(current, name, dispid, [value]);
}

function putMember(current: State, name: string, dispid: number, args: readonly unknown[]): void {
  const value = args[args.length - 1];
  if (!isObject(value)) {
    call(current, name, dispid, DISPATCH_PROPERTYPUT, args);
    return;
  }
  const member = memberOf(current, dispid);
  if (member && !member.putref) {
    call(current, name, dispid, DISPATCH_PROPERTYPUT, args);
    return;
  }
  try {
    call(current, name, dispid, DISPATCH_PROPERTYPUTREF, args);
  } catch (error) {
    if (!member && error instanceof COMError && BAD_CONTEXT.has(error.hresult)) call(current, name, dispid, DISPATCH_PROPERTYPUT, args);
    else throw error;
  }
}

function iterate(current: State): Generator<any, void, undefined> {
  return iterateNewEnum(current.pointer, wrap);
}

/** Values of a collection through `_NewEnum` (`DISPID_NEWENUM`) and `IEnumVARIANT`, converted by `convert`. */
export function* iterateNewEnum(pointer: DispatchPointer, convert: (value: unknown) => unknown): Generator<any, void, undefined> {
  const result = invoke(pointer, DISPID_NEWENUM, DISPATCH_METHOD | DISPATCH_PROPERTYGET, [], { name: "_NewEnum" });
  if (!(result instanceof COMPointer)) throw new TypeError("_NewEnum did not return an object");
  let enumerator: InterfacePointer<typeof IEnumVARIANT>;
  try {
    enumerator = result.QueryInterface(IEnumVARIANT);
  } finally {
    result.Release();
  }
  try {
    using item = new VARIANT();
    const fetched = new ctypes.c_uint32();
    for (;;) {
      const hr = enumerator.Next(1, item, ctypes.byref(fetched));
      if (hr !== S_OK || fetched.value === 0) return;
      let value: unknown;
      try {
        value = convert(item.value);
      } finally {
        item.clear();
      }
      yield value;
    }
  } finally {
    enumerator.Release();
  }
}

function describe(current: State): string {
  if (current.pointer.released) return "[Dispatch released]";
  loadTypeInfo(current);
  return `[Dispatch ${current.typeName ?? "object"}]`;
}

function createProxy(pointer: DispatchPointer): DispatchObject {
  const current: State = { pointer, dispids: new Map(), members: undefined, typeName: undefined };
  const target = (() => undefined) as unknown as DispatchObject;
  // util.inspect does not go through proxy traps: it reads the target.
  Object.defineProperty(target, inspect, { value: () => describe(current) });
  const proxy = new Proxy(target, {
    get(_target, property) {
      if (typeof property === "symbol") {
        switch (property) {
          case comPointer:
            return current.pointer;
          case Symbol.iterator:
            return () => iterate(current);
          case Symbol.dispose:
            return () => {
              if (!current.pointer.released) current.pointer.Release();
            };
          case Symbol.toPrimitive:
            return () => describe(current);
          case inspect:
            return () => describe(current);
          default:
            return undefined;
        }
      }
      // Not a thenable: `await obj` must not invoke a COM member named "then".
      if (property === "then") return undefined;
      return getMember(current, property);
    },
    set(_target, property, value) {
      if (typeof property === "symbol") return false;
      setMember(current, property, value);
      return true;
    },
    has(_target, property) {
      if (typeof property === "symbol") return property === comPointer || property === Symbol.iterator || property === Symbol.dispose;
      try {
        dispidOf(current, property);
        return true;
      } catch (error) {
        if (error instanceof COMError && error.hresult === DISP_E_UNKNOWNNAME) return false;
        throw error;
      }
    },
    apply(_target, _this, args) {
      return call(current, null, DISPID_VALUE, DISPATCH_METHOD | DISPATCH_PROPERTYGET, args);
    },
    ownKeys() {
      return [];
    },
    getOwnPropertyDescriptor() {
      return undefined;
    },
    defineProperty() {
      return false;
    },
    deleteProperty() {
      return false;
    },
  });
  states.set(proxy, current);
  return proxy;
}

/**
 * Late-bound wrapper for a COM object (pointer, dynamic object or JavaScript
 * `COMObject`). The wrapper owns its own `IDispatch` reference: the argument
 * is not consumed.
 */
export function Dispatch(object: COMPointer | DispatchObject | ComObject): DispatchObject {
  const borrowed = borrowReference(object);
  if (!borrowed) throw new TypeError("Dispatch expects a COM pointer, a dispatch object or a COMObject");
  const [hr, address] = rawQueryInterface(borrowed.address, IDispatch.iid);
  checkHRESULT(hr, "QueryInterface(IDispatch)");
  return createProxy(COMPointer.fromAddress(address!, IDispatch));
}

/** Wrap an owned `IDispatch` pointer, taking over its reference. */
/**
 * Target of the late-bound helpers (`getProperty`, `putProperty`,
 * `invokeMethod`, ...): a dispatch object or a typed pointer deriving from
 * `IDispatch`. Results are dynamic dispatch objects.
 */
export type DispatchTarget = DispatchObject | COMPointer;

/**
 * An automation object whose members are known at run time only: its best
 * interface, typed from its type library, or a dynamic dispatch object when
 * it has no type information. Returned by `CreateObject`, `GetObject` and
 * `GetActiveObject` without `interface` or `dynamic`, and by typed calls
 * declared to return `IDispatch`.
 */
export type AutomationObject = (COMPointer & Iterable<any> & { [member: string]: any }) | DispatchObject;

export function adoptDispatch(pointer: DispatchPointer): DispatchObject {
  return createProxy(pointer);
}

export function isDispatch(value: unknown): value is DispatchObject {
  return typeof value === "function" && states.has(value);
}

/** The `IDispatch` pointer of a dispatch object (still owned by the object). */
export function dispatchPointer(object: DispatchTarget): DispatchPointer {
  return state(object).pointer;
}

/** Release a dispatch object (same as `using`). */
export function release(object: DispatchTarget): void {
  const { pointer } = state(object);
  if (!pointer.released) pointer.Release();
}

/** Iterate a collection through `_NewEnum`. */
export function enumerate(object: DispatchTarget): Generator<any, void, undefined> {
  return iterate(state(object));
}

/** Property get with arguments (`DISPATCH_PROPERTYGET` only). */
export function getProperty(object: DispatchTarget, name: string, ...args: unknown[]): any {
  const current = state(object);
  return call(current, name, dispidOf(current, name), DISPATCH_PROPERTYGET, args);
}

/** Property put, also parameterized: `putProperty(dict, "Item", key, value)`. The last argument is the value. */
export function putProperty(object: DispatchTarget, name: string, ...argsAndValue: unknown[]): void {
  const current = state(object);
  call(current, name, dispidOf(current, name), DISPATCH_PROPERTYPUT, argsAndValue);
}

/** Property put by reference (`Set obj.Name = value`). */
export function putRefProperty(object: DispatchTarget, name: string, ...argsAndValue: unknown[]): void {
  const current = state(object);
  call(current, name, dispidOf(current, name), DISPATCH_PROPERTYPUTREF, argsAndValue);
}

/** Method call (`DISPATCH_METHOD` only). */
export function invokeMethod(object: DispatchTarget, name: string, ...args: unknown[]): any {
  const current = state(object);
  return call(current, name, dispidOf(current, name), DISPATCH_METHOD, args);
}

/** Member information from the type library, when available. */
export function memberInfo(object: DispatchTarget, name: string): MemberInfo | null {
  const current = state(object);
  return memberOf(current, dispidOf(current, name)) ?? null;
}

/** Name of the object's type from its type information, or `null`. */
export function typeName(object: DispatchTarget): string | null {
  const current = state(object);
  loadTypeInfo(current);
  return current.typeName ?? null;
}
