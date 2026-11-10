import * as ctypes from "node-ctypes";
import { DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPATCH_PROPERTYPUTREF, DISPID_NEWENUM, DISPID_PROPERTYPUT, DISPID_UNKNOWN, DISPID_VALUE, VT_BYREF, VT_VARIANT } from "../automation/constants.ts";
import { DISPPARAMS, EXCEPINFO } from "../automation/types.ts";
import { ByRef, VARIANT, VARIANT_SIZE, clearVariant, readVariant, writeReference, writeVariant } from "../automation/variant.ts";
import { adoptDispatch, isDispatch, release } from "../client/dynamic.ts";
import type { DispatchPointer } from "../automation/dispatch.ts";
import { COMError } from "../errors.ts";
import { DISP_E_BADPARAMCOUNT, DISP_E_EXCEPTION, DISP_E_MEMBERNOTFOUND, DISP_E_UNKNOWNNAME, E_FAIL, E_POINTER, S_OK } from "../hresult.ts";
import { derivesFrom } from "../interface.ts";
import type { InterfaceDef, MethodSpec } from "../interface.ts";
import { IDispatch } from "../interfaces/dispatch.ts";
import { IUnknown } from "../interfaces/unknown.ts";
import { allocBSTR, POINTER_SIZE, readMemory } from "../low-level/native.ts";
import { COMPointer } from "../pointer.ts";
import { COMObject, VariantEnumerator } from "./comobject.ts";
import { omittedToUndefined } from "./marshal.ts";

/** Arguments of a late-bound call: positional in natural order, named by DISPID (`DISPID_PROPERTYPUT` for puts). */
export interface InvokeArguments {
  readonly positional: readonly unknown[];
  readonly named: ReadonlyMap<number, unknown>;
  /**
   * `VT_BYREF` positional arguments as `ByRef` (same index as `positional`):
   * assigning `value` changes the caller's variable after the call.
   */
  readonly references: readonly (ByRef | undefined)[];
}

/** A dispatch-level failure (`DISP_E_*`): returned as the HRESULT, not as an exception. */
export class DispatchError extends COMError {}

/**
 * Thrown by a method implementation to return `hresult`, also a success code
 * such as `S_FALSE`, without calling `onError`; failures carry `text` as
 * error information (`IErrorInfo`, or `EXCEPINFO` through `IDispatch`).
 * Python equivalent: `comtypes.ReturnHRESULT`.
 */
export class ReturnHRESULT extends COMError {
  constructor(hresult: number, text?: string) {
    super(hresult, text === undefined ? {} : { text });
    this.name = "ReturnHRESULT";
  }
}

interface Member {
  readonly name: string;
  readonly dispid: number;
  /** Input parameter names from COMMETHOD specs (for named arguments). */
  readonly params: readonly string[];
}

interface Table {
  readonly byName: Map<string, Member>;
  readonly byId: Map<number, Member>;
  next: number;
}

const tables = new WeakMap<Function, Table>();

function specParameters(spec: MethodSpec): string[] {
  const params = (spec as MethodSpec & { readonly params?: readonly { readonly name: string; readonly input: boolean; readonly lcid: boolean }[] }).params ?? [];
  return params.filter((param) => param.input && !param.lcid).map((param) => param.name);
}

/** Members declared by the class's dual interfaces (beyond IDispatch), with their DISPIDs. */
function tableOf(object: COMObject): Table {
  const type = object.constructor;
  let table = tables.get(type);
  if (table) return table;
  const declared = new Map<string, { name: string; dispid: number | null; params: string[] }>();
  for (const iface of object.interfaces) {
    if (!derivesFrom(iface, IDispatch.iid)) continue;
    for (let current: InterfaceDef<any> | null = iface; current && !current.iid.equals(IDispatch.iid) && !current.iid.equals(IUnknown.iid); current = current.base) {
      for (const spec of [...current.methods, ...current.dispMethods.slice(current.base?.dispMethods.length ?? 0)]) {
        const key = spec.property.toLowerCase();
        const entry = declared.get(key) ?? { name: spec.property, dispid: null, params: [] };
        if (spec.dispid !== null) entry.dispid = spec.dispid;
        if (spec.kind === "method" || spec.kind === "propget") entry.params = entry.params.length ? entry.params : specParameters(spec);
        declared.set(key, entry);
      }
    }
  }
  table = { byName: new Map(), byId: new Map(), next: 1 };
  const used = new Set([...declared.values()].map((entry) => entry.dispid).filter((dispid): dispid is number => dispid !== null));
  for (const [key, entry] of declared) {
    let dispid = entry.dispid;
    if (dispid === null) {
      while (used.has(table.next)) table.next++;
      dispid = table.next++;
      used.add(dispid);
    }
    const member = { name: entry.name, dispid, params: entry.params };
    table.byName.set(key, member);
    table.byId.set(dispid, member);
  }
  tables.set(type, table);
  return table;
}

/** Public JavaScript members: own keys and prototype members up to `COMObject`, without `_`-prefixed names. */
function memberNames(object: COMObject): string[] {
  const names = new Set<string>(Object.keys(object));
  for (let prototype = Object.getPrototypeOf(object); prototype && prototype !== COMObject.prototype; prototype = Object.getPrototypeOf(prototype)) {
    for (const name of Object.getOwnPropertyNames(prototype)) names.add(name);
  }
  return [...names].filter((name) => name !== "constructor" && !name.startsWith("_"));
}

function resolveName(object: COMObject, name: string): Member | null {
  const table = tableOf(object);
  const key = name.toLowerCase();
  const known = table.byName.get(key);
  if (known) return known;
  const actual = memberNames(object).find((candidate) => candidate.toLowerCase() === key);
  if (!actual) return null;
  while (table.byId.has(table.next)) table.next++;
  const member = { name: actual, dispid: table.next++, params: [] };
  table.byName.set(key, member);
  table.byId.set(member.dispid, member);
  return member;
}

/** Default `GetIDsOfNames`: members of the dual interfaces, then the JavaScript members. */
export function defaultIdsOfNames(object: COMObject, names: readonly string[]): number[] {
  const member = resolveName(object, names[0]);
  if (!member) throw new DispatchError(DISP_E_UNKNOWNNAME);
  return names.map((name, index) => {
    if (index === 0) return member.dispid;
    const position = member.params.findIndex((param) => param.toLowerCase() === name.toLowerCase());
    if (position < 0) throw new DispatchError(DISP_E_UNKNOWNNAME);
    return position;
  });
}

function memberName(object: COMObject, dispid: number): string {
  if (dispid === DISPID_VALUE) {
    const name = (object.constructor as typeof COMObject)._com_default_member_ ?? tableOf(object).byId.get(DISPID_VALUE)?.name;
    if (name) return name;
  }
  const member = tableOf(object).byId.get(dispid);
  if (!member) throw new DispatchError(DISP_E_MEMBERNOTFOUND);
  return member.name;
}

/**
 * Default late-bound call: `_set_X`/`_setref_X`/`_get_X` methods or the
 * JavaScript member `X` (method, accessor or field). Iterable objects answer
 * `_NewEnum` with a `VariantEnumerator`.
 */
export function defaultInvoke(object: COMObject, dispid: number, flags: number, args: InvokeArguments): unknown {
  if (dispid === DISPID_NEWENUM) {
    const iterable = object as unknown as Partial<Iterable<unknown>>;
    if (typeof iterable[Symbol.iterator] !== "function") throw new DispatchError(DISP_E_MEMBERNOTFOUND);
    return new VariantEnumerator(iterable as Iterable<unknown>);
  }
  const name = memberName(object, dispid);
  const record = object as unknown as Record<string, unknown>;
  const positional = [...args.positional];
  for (const [id, value] of args.named) if (id >= 0) positional[id] = value;
  if (flags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) {
    const value = args.named.get(DISPID_PROPERTYPUT);
    const setter = flags & DISPATCH_PROPERTYPUTREF && typeof record[`_setref_${name}`] === "function" ? record[`_setref_${name}`] : record[`_set_${name}`];
    if (typeof setter === "function") {
      setter.call(object, ...positional, value);
      return undefined;
    }
    if (!positional.length && name in object && typeof record[name] !== "function") {
      record[name] = value;
      return undefined;
    }
    throw new DispatchError(DISP_E_MEMBERNOTFOUND);
  }
  const getter = record[`_get_${name}`];
  if (flags & DISPATCH_PROPERTYGET && typeof getter === "function") return getter.apply(object, positional);
  const member = record[name];
  if (typeof member === "function") {
    if (flags & DISPATCH_METHOD) return member.apply(object, positional);
    throw new DispatchError(DISP_E_MEMBERNOTFOUND);
  }
  if (!(name in object)) throw new DispatchError(DISP_E_MEMBERNOTFOUND);
  if (positional.length) throw new DispatchError(DISP_E_BADPARAMCOUNT);
  return member;
}

function readNames(rgszNames: bigint, count: number): string[] {
  return Array.from({ length: count }, (_, index) => {
    const text = ctypes.readValue(rgszNames + BigInt(index * POINTER_SIZE), ctypes.c_void_p);
    return text ? (ctypes.wstring_at(text) ?? "") : "";
  });
}

/** `IDispatch::GetIDsOfNames` for a `COMObject`, through its `_ids_of_names_`. */
export function dispatchGetIDsOfNames(object: COMObject, rgszNames: bigint | null, cNames: number, rgDispId: bigint | null): number {
  if (!rgszNames || !rgDispId) return E_POINTER;
  const write = (ids: readonly number[]) => ids.forEach((id, index) => ctypes.writeValue(rgDispId, ctypes.c_int32, id, index * 4));
  try {
    write(object._ids_of_names_(readNames(rgszNames, cNames)));
    return S_OK;
  } catch (error) {
    write(Array.from({ length: cNames }, () => DISPID_UNKNOWN));
    return error instanceof COMError && error.hresult < 0 ? error.hresult : DISP_E_UNKNOWNNAME;
  }
}

/** Incoming dispatch pointers become dynamic objects for the implementation. */
function incoming(value: unknown): unknown {
  const plain = omittedToUndefined(value);
  return plain instanceof COMPointer && derivesFrom(plain.interface, IDispatch.iid) ? adoptDispatch(plain as DispatchPointer) : plain;
}

function releaseIncoming(values: readonly unknown[]): void {
  for (const value of values) {
    if (isDispatch(value)) release(value);
    else if (value instanceof COMPointer && !value.released) value.Release();
  }
}

function fillException(pExcepInfo: bigint, object: COMObject, error: unknown): void {
  const type = object.constructor as typeof COMObject;
  const info = EXCEPINFO.from_address(pExcepInfo);
  info._buffer.fill(0);
  const description = error instanceof COMError ? (error.details?.description ?? error.text) : error instanceof Error ? error.message : String(error);
  info.bstrSource = allocBSTR(type._reg_progid_ ?? type.name);
  info.bstrDescription = allocBSTR(description);
  info.scode = error instanceof COMError && error.hresult < 0 ? error.hresult : E_FAIL;
}

/**
 * `IDispatch::Invoke` for a `COMObject`: converts `DISPPARAMS`, calls
 * `_invoke_` and writes the result. `DispatchError`s are returned as
 * HRESULTs; other exceptions fill `EXCEPINFO` (`DISP_E_EXCEPTION`).
 */
export function dispatchInvoke(object: COMObject, dispid: number, flags: number, pDispParams: bigint | null, pVarResult: bigint | null, pExcepInfo: bigint | null, _puArgErr: bigint | null): number {
  if (!pDispParams) return E_POINTER;
  const params = DISPPARAMS.from_address(pDispParams);
  const values: unknown[] = [];
  const slots: { address: bigint; vt: number; reference: ByRef }[] = [];
  const references: (ByRef | undefined)[] = [];
  try {
    for (let index = 0; index < params.cArgs; index++) {
      const variant = readMemory(params.rgvarg! + BigInt(index * VARIANT_SIZE), VARIANT_SIZE);
      const value = incoming(readVariant(variant));
      values.push(value);
      const vt = variant.readUInt16LE(0);
      if (vt & VT_BYREF && index >= params.cNamedArgs) {
        const target = POINTER_SIZE === 8 ? variant.readBigUInt64LE(8) : BigInt(variant.readUInt32LE(8));
        const reference = new ByRef(value, (vt & ~VT_BYREF) === VT_VARIANT ? null : vt & ~VT_BYREF);
        slots.push({ address: target, vt: vt & ~VT_BYREF, reference });
        references[params.cArgs - 1 - index] = reference;
      }
    }
    const named = new Map<number, unknown>();
    for (let index = 0; index < params.cNamedArgs; index++) named.set(ctypes.readValue(params.rgdispidNamedArgs! + BigInt(index * 4), ctypes.c_int32), values[index]);
    const positional = values.slice(params.cNamedArgs).reverse();
    const initial = slots.map((slot) => slot.reference.value);
    const result = object._invoke_(dispid, flags, { positional, named, references });
    slots.forEach((slot, index) => {
      if (slot.address && slot.reference.value !== initial[index]) writeReference(slot.address, slot.vt, slot.reference.value);
    });
    if (pVarResult && result !== undefined) {
      const view = VARIANT.from_address(pVarResult);
      clearVariant(view._buffer);
      writeVariant(view._buffer, result);
    }
    return S_OK;
  } catch (error) {
    if (error instanceof DispatchError) return error.hresult;
    if (error instanceof ReturnHRESULT) {
      if (error.hresult >= 0) return S_OK;
    } else {
      (object.constructor as typeof COMObject).onError?.(error, object);
    }
    if (pExcepInfo) {
      try {
        fillException(pExcepInfo, object, error);
      } catch {
        // Still report the exception HRESULT.
      }
    }
    return DISP_E_EXCEPTION;
  } finally {
    releaseIncoming(values);
  }
}
