import * as ctypes from "node-ctypes";
import { IMPLTYPEFLAG_FDEFAULT, IMPLTYPEFLAG_FSOURCE, TKIND_DISPATCH, TYPEFLAG_FDUAL } from "../automation/constants.ts";
import { dispatchTypeInfo, withTypeAttr } from "../automation/members.ts";
import { setOutputWrapper } from "../commethod.ts";
import { COMError } from "../errors.ts";
import { derivesFrom, findInterface } from "../interface.ts";
import type { InterfaceDef } from "../interface.ts";
import { IDispatch } from "../interfaces/dispatch.ts";
import { IProvideClassInfo, ITypeInfo, ITypeLib } from "../interfaces/typeinfo.ts";
import { COMPointer, sync } from "../pointer.ts";
import { GetModule, getModuleSync } from "../typelib/getmodule.ts";
import type { TypeInfoPointer, TypeLibPointer } from "../typelib/load.ts";
import { queryOrDispatch } from "./dual.ts";
import { adoptDispatch, iterateNewEnum, iterateNewEnumAsync } from "./dynamic.ts";
import type { DispatchPointer } from "../automation/dispatch.ts";

/** The type info of implemented interface `index` (`-1`: the vtable part of a dual interface). */
function implementedType(info: TypeInfoPointer, index: number): TypeInfoPointer {
  const href = new ctypes.c_uint32();
  sync(info).GetRefTypeOfImplType(index >>> 0, ctypes.byref(href));
  const out = new ctypes.c_void_p();
  sync(info).GetRefTypeInfo(href.value, ctypes.byref(out));
  return COMPointer.fromAddress(out.value!, ITypeInfo);
}

/** Default (non-source) interface of a coclass. */
function defaultInterface(classInfo: TypeInfoPointer): TypeInfoPointer | null {
  const count = withTypeAttr(classInfo, (attr) => attr.cImplTypes);
  const flags = new ctypes.c_int32();
  for (let index = 0; index < count; index++) {
    sync(classInfo).GetImplTypeFlags(index, ctypes.byref(flags));
    if ((flags.value & (IMPLTYPEFLAG_FDEFAULT | IMPLTYPEFLAG_FSOURCE)) === IMPLTYPEFLAG_FDEFAULT) return implementedType(classInfo, index);
  }
  return count === 1 ? implementedType(classInfo, 0) : null;
}

/** Type information of the main interface: `IProvideClassInfo`, else `IDispatch::GetTypeInfo`. */
function objectTypeInfo(pointer: COMPointer): TypeInfoPointer | null {
  const provider = pointer.tryQueryInterface(IProvideClassInfo);
  if (provider) {
    using owned = provider;
    try {
      using classInfo = sync(owned).GetClassInfo();
      const info = classInfo ? defaultInterface(classInfo) : null;
      if (info) return info;
    } catch (error) {
      if (!(error instanceof COMError)) throw error;
    }
  }
  const dispatch = pointer.tryQueryInterface(IDispatch);
  if (!dispatch) return null;
  using owned = dispatch;
  return dispatchTypeInfo(owned);
}

/** The vtable part of a dual interface described as a dispinterface, else `null`. */
function dualVtable(info: TypeInfoPointer): TypeInfoPointer | null {
  if (!withTypeAttr(info, (attr) => attr.typekind === TKIND_DISPATCH && (attr.wTypeFlags & TYPEFLAG_FDUAL) !== 0)) return null;
  try {
    return implementedType(info, -1);
  } catch (error) {
    if (error instanceof COMError) return null;
    throw error;
  }
}

function isInterfaceDef(value: unknown): value is InterfaceDef<any> {
  return typeof value === "object" && value !== null && "iid" in value && "Pointer" in value;
}

function containingLibrary(info: TypeInfoPointer): TypeLibPointer {
  const out = new ctypes.c_void_p();
  const index = new ctypes.c_uint32();
  sync(info).GetContainingTypeLib(ctypes.byref(out), ctypes.byref(index));
  return COMPointer.fromAddress(out.value!, ITypeLib);
}

/** Definition of the type described by `info`, built by `GetModule` from its type library. */
function interfaceOf(info: TypeInfoPointer): InterfaceDef<any> | null {
  const iid = withTypeAttr(info, (attr) => attr.guid.toString());
  using library = containingLibrary(info);
  for (const value of Object.values(getModuleSync(library))) if (isInterfaceDef(value) && value.iid.toString() === iid) return value;
  return findInterface(iid) ?? null;
}

/**
 * Load the type library {@link GetBestInterface} needs for `object` with
 * `await GetModule` (in slices, without blocking the event loop), so that it
 * finds the module cached. Objects without type information need nothing.
 */
export async function loadBestModule(object: COMPointer): Promise<void> {
  let library: TypeLibPointer;
  try {
    using info = objectTypeInfo(object);
    if (!info) return;
    using vtable = dualVtable(info);
    library = containingLibrary(vtable ?? info);
  } catch (error) {
    if (error instanceof COMError) return;
    throw error;
  }
  using owned = library;
  try {
    await GetModule(owned);
  } catch (error) {
    // GetBestInterface falls back to the object's own interface.
    if (!(error instanceof COMError)) throw error;
  }
}

/**
 * The most specific interface of an object, typed from its type library
 * (loaded with `GetModule`): the default interface of its coclass
 * (`IProvideClassInfo`), else the interface `IDispatch::GetTypeInfo`
 * describes. Returns a new reference; without type information, a new
 * reference with the same interface. Python equivalent:
 * `comtypes.client.GetBestInterface` (also `wrap`).
 */
export function GetBestInterface<T extends COMPointer>(object: T): COMPointer {
  try {
    using info = objectTypeInfo(object);
    if (info) {
      using vtable = dualVtable(info);
      const iface = interfaceOf(vtable ?? info);
      if (iface && iface !== object.interface) {
        // A dispinterface, or a dual interface whose vtable cannot be marshaled,
        // is reached through IDispatch::Invoke.
        const typed = queryOrDispatch(object, iface);
        if (typed) return typed;
      }
    }
  } catch (error) {
    if (!(error instanceof COMError)) throw error;
  }
  return object.QueryInterface(object.interface);
}

/** Python equivalent: `comtypes.client.wrap`. */
export const wrap = GetBestInterface;

/**
 * Values returned by typed calls: plain `IDispatch` pointers become their best
 * interface (Python: `wrap_outparam`), also inside arrays.
 */
function wrapOutparam(value: unknown): unknown {
  if (value instanceof COMPointer) {
    if (value.interface !== IDispatch || value.released) return value;
    const best = GetBestInterface(value);
    value.Release();
    // Without type information: a dynamic object (Python returns the bare IDispatch pointer).
    return best.interface === IDispatch ? adoptDispatch(best as DispatchPointer) : best;
  }
  if (Array.isArray(value) && value.some((item) => item instanceof COMPointer || Array.isArray(item))) return value.map(wrapOutparam);
  return value;
}

setOutputWrapper(wrapOutparam);

// Typed dispatch pointers iterate their `_NewEnum` collection (Python: `__iter__`),
// also with `for await` (IEnumVARIANT::Next on the COM executor).
Object.defineProperty(IDispatch.Pointer.prototype, Symbol.iterator, {
  configurable: true,
  writable: true,
  value: function (this: DispatchPointer) {
    return iterateNewEnum(this, wrapOutparam);
  },
});
Object.defineProperty(IDispatch.Pointer.prototype, Symbol.asyncIterator, {
  configurable: true,
  writable: true,
  value: function (this: DispatchPointer) {
    return iterateNewEnumAsync(this, wrapOutparam);
  },
});
