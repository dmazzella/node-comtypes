import * as ctypes from "node-ctypes";
import { IMPLTYPEFLAG_FDEFAULT, IMPLTYPEFLAG_FSOURCE, TKIND_DISPATCH, TYPEFLAG_FDUAL } from "../automation/constants.ts";
import { dispatchTypeInfo, withTypeAttr } from "../automation/members.ts";
import { setOutputWrapper } from "../commethod.ts";
import { COMError } from "../errors.ts";
import { derivesFrom, findInterface } from "../interface.ts";
import type { InterfaceDef } from "../interface.ts";
import { IDispatch } from "../interfaces/dispatch.ts";
import { IProvideClassInfo, ITypeInfo, ITypeLib } from "../interfaces/typeinfo.ts";
import { COMPointer } from "../pointer.ts";
import { GetModule } from "../typelib/getmodule.ts";
import type { TypeInfoPointer } from "../typelib/load.ts";
import { adoptDispatch, iterateNewEnum } from "./dynamic.ts";
import type { DispatchPointer } from "../automation/dispatch.ts";

/** The type info of implemented interface `index` (`-1`: the vtable part of a dual interface). */
function implementedType(info: TypeInfoPointer, index: number): TypeInfoPointer {
  const href = new ctypes.c_uint32();
  info.GetRefTypeOfImplType(index >>> 0, ctypes.byref(href));
  const out = new ctypes.c_void_p();
  info.GetRefTypeInfo(href.value, ctypes.byref(out));
  return COMPointer.fromAddress(out.value!, ITypeInfo);
}

/** Default (non-source) interface of a coclass. */
function defaultInterface(classInfo: TypeInfoPointer): TypeInfoPointer | null {
  const count = withTypeAttr(classInfo, (attr) => attr.cImplTypes);
  const flags = new ctypes.c_int32();
  for (let index = 0; index < count; index++) {
    classInfo.GetImplTypeFlags(index, ctypes.byref(flags));
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
      using classInfo = owned.GetClassInfo();
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

/** Definition of the type described by `info`, built by `GetModule` from its type library. */
function interfaceOf(info: TypeInfoPointer): InterfaceDef<any> | null {
  const iid = withTypeAttr(info, (attr) => attr.guid.toString());
  const out = new ctypes.c_void_p();
  const index = new ctypes.c_uint32();
  info.GetContainingTypeLib(ctypes.byref(out), ctypes.byref(index));
  using library = COMPointer.fromAddress(out.value!, ITypeLib);
  for (const value of Object.values(GetModule(library))) if (isInterfaceDef(value) && value.iid.toString() === iid) return value;
  return findInterface(iid) ?? null;
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
        const typed = object.tryQueryInterface(iface);
        if (typed) return typed;
        // A dispinterface is called through IDispatch::Invoke only: objects
        // may not answer QueryInterface for its DIID.
        if (iface.allMethods.length === IDispatch.allMethods.length && derivesFrom(iface, IDispatch.iid)) {
          const dispatch = object.tryQueryInterface(IDispatch);
          if (dispatch) {
            using owned = dispatch;
            return COMPointer.fromAddress(owned.address, iface, { addRef: true });
          }
        }
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

// Typed dispatch pointers iterate their `_NewEnum` collection (Python: `__iter__`).
Object.defineProperty(IDispatch.Pointer.prototype, Symbol.iterator, {
  configurable: true,
  writable: true,
  value: function (this: DispatchPointer) {
    return iterateNewEnum(this, wrapOutparam);
  },
});
