import * as ctypes from "node-ctypes";
import { COMError } from "../errors.ts";
import type { InterfacePointer } from "../interface.ts";
import { ITypeInfo, ELEMDESC, FUNCDESC, TYPEATTR, VARDESC } from "../interfaces/typeinfo.ts";
import { takeBSTR } from "../low-level/native.ts";
import { COMPointer } from "../pointer.ts";
import { INVOKE_FUNC, INVOKE_PROPERTYGET, INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF, MEMBERID_NIL, PARAMFLAG_FHASDEFAULT, PARAMFLAG_FLCID, PARAMFLAG_FOPT, PARAMFLAG_FRETVAL, VAR_DISPATCH } from "./constants.ts";
import type { DispatchPointer } from "./dispatch.ts";
import { defaultLCID } from "./dispatch.ts";

export type TypeInfoPointer = InterfacePointer<typeof ITypeInfo>;

/** How a member can be invoked, from the type information. */
export interface MemberInfo {
  readonly memid: number;
  readonly name: string | null;
  readonly method: boolean;
  readonly get: boolean;
  readonly put: boolean;
  readonly putref: boolean;
  /** Required arguments of the property get (0 for a plain property). */
  readonly getArguments: number;
}

/** `IDispatch::GetTypeInfo(0)`, or `null` when the object provides none. */
export function dispatchTypeInfo(dispatch: DispatchPointer, lcid = defaultLCID): TypeInfoPointer | null {
  try {
    const count = new ctypes.c_uint32();
    dispatch.GetTypeInfoCount(ctypes.byref(count));
    if (!count.value) return null;
    const out = new ctypes.c_void_p();
    dispatch.GetTypeInfo(0, lcid, ctypes.byref(out));
    return out.value ? COMPointer.fromAddress(out.value, ITypeInfo) : null;
  } catch (error) {
    if (error instanceof COMError) return null;
    throw error;
  }
}

/** Run `use` with the TYPEATTR of `typeInfo`, released afterwards. */
export function withTypeAttr<T>(typeInfo: TypeInfoPointer, use: (attr: TYPEATTR) => T): T {
  const out = new ctypes.c_void_p();
  typeInfo.GetTypeAttr(ctypes.byref(out));
  try {
    return use(TYPEATTR.from_address(out.value!));
  } finally {
    typeInfo.ReleaseTypeAttr(out.value);
  }
}

export function withFuncDesc<T>(typeInfo: TypeInfoPointer, index: number, use: (desc: FUNCDESC) => T): T {
  const out = new ctypes.c_void_p();
  typeInfo.GetFuncDesc(index, ctypes.byref(out));
  try {
    return use(FUNCDESC.from_address(out.value!));
  } finally {
    typeInfo.ReleaseFuncDesc(out.value);
  }
}

export function withVarDesc<T>(typeInfo: TypeInfoPointer, index: number, use: (desc: VARDESC) => T): T {
  const out = new ctypes.c_void_p();
  typeInfo.GetVarDesc(index, ctypes.byref(out));
  try {
    return use(VARDESC.from_address(out.value!));
  } finally {
    typeInfo.ReleaseVarDesc(out.value);
  }
}

export interface Documentation {
  readonly name: string | null;
  readonly docString: string | null;
  readonly helpContext: number;
  readonly helpFile: string | null;
}

/** `ITypeInfo::GetDocumentation`; `MEMBERID_NIL` documents the type itself. */
export function documentation(typeInfo: TypeInfoPointer, memid = MEMBERID_NIL): Documentation {
  const name = new ctypes.c_void_p();
  const doc = new ctypes.c_void_p();
  const context = new ctypes.c_uint32();
  const file = new ctypes.c_void_p();
  typeInfo.GetDocumentation(memid, ctypes.byref(name), ctypes.byref(doc), ctypes.byref(context), ctypes.byref(file));
  return { name: takeBSTR(name.value), docString: takeBSTR(doc.value), helpContext: context.value, helpFile: takeBSTR(file.value) };
}

/** Parameter flags of a FUNCDESC. */
export function parameterFlags(desc: FUNCDESC): number[] {
  const size = ctypes.sizeof(ELEMDESC);
  return Array.from({ length: desc.cParams }, (_, index) => ELEMDESC.from_address(desc.lprgelemdescParam! + BigInt(index * size)).paramdesc.wParamFlags);
}

interface MutableMember {
  memid: number;
  name: string | null;
  method: boolean;
  get: boolean;
  put: boolean;
  putref: boolean;
  getArguments: number;
}

function collect(typeInfo: TypeInfoPointer, members: Map<number, MutableMember>, seen: Set<string>): void {
  const { key, functions, variables, implemented } = withTypeAttr(typeInfo, (attr) => ({ key: attr.guid.toString(), functions: attr.cFuncs, variables: attr.cVars, implemented: attr.cImplTypes }));
  if (seen.has(key)) return;
  seen.add(key);
  const member = (memid: number) => {
    let entry = members.get(memid);
    if (!entry) members.set(memid, (entry = { memid, name: null, method: false, get: false, put: false, putref: false, getArguments: 0 }));
    return entry;
  };
  for (let index = 0; index < functions; index++) {
    withFuncDesc(typeInfo, index, (desc) => {
      const entry = member(desc.memid);
      const required = parameterFlags(desc).filter((flags) => !(flags & (PARAMFLAG_FOPT | PARAMFLAG_FHASDEFAULT | PARAMFLAG_FRETVAL | PARAMFLAG_FLCID))).length;
      if (desc.invkind & INVOKE_FUNC) entry.method = true;
      if (desc.invkind & INVOKE_PROPERTYGET) {
        entry.getArguments = entry.get ? Math.min(entry.getArguments, required) : required;
        entry.get = true;
      }
      if (desc.invkind & INVOKE_PROPERTYPUT) entry.put = true;
      if (desc.invkind & INVOKE_PROPERTYPUTREF) entry.putref = true;
    });
  }
  for (let index = 0; index < variables; index++) {
    withVarDesc(typeInfo, index, (desc) => {
      if (desc.varkind !== VAR_DISPATCH) return;
      const entry = member(desc.memid);
      entry.get = true;
      entry.put = true;
    });
  }
  for (const entry of members.values()) {
    if (entry.name === null) {
      try {
        entry.name = documentation(typeInfo, entry.memid).name;
      } catch {
        // Members of base interfaces are documented by their own type info.
      }
    }
  }
  for (let index = 0; index < implemented; index++) {
    const reference = new ctypes.c_uint32();
    const base = new ctypes.c_void_p();
    try {
      typeInfo.GetRefTypeOfImplType(index, ctypes.byref(reference));
      typeInfo.GetRefTypeInfo(reference.value, ctypes.byref(base));
    } catch {
      continue;
    }
    using baseInfo = COMPointer.fromAddress(base.value!, ITypeInfo);
    collect(baseInfo, members, seen);
  }
}

const cache = new Map<string, ReadonlyMap<number, MemberInfo>>();

/** Members of a type and its base types, by MEMBERID. Cached by type GUID and version. */
export function typeMembers(typeInfo: TypeInfoPointer): ReadonlyMap<number, MemberInfo> {
  const key = withTypeAttr(typeInfo, (attr) => `${attr.guid}:${attr.wMajorVerNum}.${attr.wMinorVerNum}:${attr.lcid}`);
  let members = cache.get(key);
  if (!members) {
    const collected = new Map<number, MutableMember>();
    collect(typeInfo, collected, new Set());
    members = collected;
    // GUID_NULL identifies ad-hoc type infos: never share them.
    if (!key.startsWith("{00000000-0000-0000-0000-000000000000}")) cache.set(key, members);
  }
  return members;
}
