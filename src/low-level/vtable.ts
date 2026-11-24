import * as ctypes from "node-ctypes";
import { POINTER_SIZE } from "./native.ts";

const { c_int32, c_uint32, c_void_p } = ctypes;

/** Read the vtable pointer of the COM object at `object`. */
export function vtableOf(object: bigint): bigint {
  if (!object) throw new TypeError("NULL COM pointer");
  const vtable = ctypes.readValue(object, c_void_p);
  if (!vtable) throw new TypeError("COM object has a NULL vtable");
  return vtable;
}

/** Address of the function in vtable slot `slot` of `vtable`. */
export function vtableEntryFromTable(vtable: bigint, slot: number): bigint {
  const entry = ctypes.readValue(vtable, c_void_p, slot * POINTER_SIZE);
  if (!entry) throw new TypeError(`COM vtable slot ${slot} is NULL`);
  return entry;
}

/** Address of the function in vtable slot `slot` of the COM object at `object`. */
export function vtableEntry(object: bigint, slot: number): bigint {
  return vtableEntryFromTable(vtableOf(object), slot);
}

const guidScratch = Buffer.allocUnsafe(16);

/** Format a native 16-byte GUID at `address` directly into `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}` without allocating a `GUID` Structure. */
export function readGuidString(address: bigint): string {
  ctypes.memmove(guidScratch, address, 16);
  const d1 = guidScratch.readUInt32LE(0).toString(16).padStart(8, "0");
  const d2 = guidScratch.readUInt16LE(4).toString(16).padStart(4, "0");
  const d3 = guidScratch.readUInt16LE(6).toString(16).padStart(4, "0");
  const tail = guidScratch.subarray(8, 16).toString("hex");
  return `{${d1}-${d2}-${d3}-${tail.slice(0, 4)}-${tail.slice(4)}}`.toUpperCase();
}

/**
 * Bind function pointers of one prototype, cached by address: binding costs
 * far more than calling, and vtables are shared by all objects of a class.
 */
export function functionCache<TRet extends ctypes.FFIType | null, const TArgs extends readonly ctypes.FFIType[]>(prototype: () => ctypes.FuncPtrFactory<TRet, TArgs>): (address: bigint) => ctypes.TypedFFIFunction<TArgs, TRet> {
  let factory: ctypes.FuncPtrFactory<TRet, TArgs> | null = null;
  const cache = new Map<bigint, ctypes.TypedFFIFunction<TArgs, TRet>>();
  return (address) => {
    let fn = cache.get(address);
    if (!fn) cache.set(address, (fn = (factory ??= prototype())(address)));
    return fn;
  };
}

export const QueryInterfaceAt = functionCache(() => ctypes.WINFUNCTYPE(c_int32, c_void_p, c_void_p, c_void_p));
export const RefCountAt = functionCache(() => ctypes.WINFUNCTYPE(c_uint32, c_void_p));
export const PointerArgAt = functionCache(() => ctypes.WINFUNCTYPE(c_int32, c_void_p, c_void_p));

// Raw IUnknown calls without errcheck: used by error handling, finalizers
// and lifetime tracking, where throwing or recursion is not acceptable.
export function rawQueryInterface(object: bigint, iid: { readonly _buffer: Buffer }): [hr: number, pointer: bigint | null] {
  const out = new c_void_p();
  const hr = QueryInterfaceAt(vtableEntry(object, 0))(object, iid._buffer, ctypes.byref(out));
  return [hr, hr >= 0 ? out.value : null];
}

export function rawAddRef(object: bigint): number {
  return RefCountAt(vtableEntry(object, 1))(object);
}

export function rawRelease(object: bigint): number {
  return RefCountAt(vtableEntry(object, 2))(object);
}
