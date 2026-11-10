import * as ctypes from "node-ctypes";
import { POINTER_SIZE } from "./native.ts";

const { c_int32, c_uint32, c_void_p } = ctypes;

/** Address of the function in vtable slot `slot` of the COM object at `object`. */
export function vtableEntry(object: bigint, slot: number): bigint {
  if (!object) throw new TypeError("NULL COM pointer");
  const vtable = ctypes.readValue(object, c_void_p);
  if (!vtable) throw new TypeError("COM object has a NULL vtable");
  const entry = ctypes.readValue(vtable, c_void_p, slot * POINTER_SIZE);
  if (!entry) throw new TypeError(`COM vtable slot ${slot} is NULL`);
  return entry;
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

const QueryInterfaceAt = functionCache(() => ctypes.WINFUNCTYPE(c_int32, c_void_p, c_void_p, c_void_p));
const RefCountAt = functionCache(() => ctypes.WINFUNCTYPE(c_uint32, c_void_p));

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
