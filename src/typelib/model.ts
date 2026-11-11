import * as ctypes from "node-ctypes";
import {
  FUNCFLAG_FHIDDEN,
  FUNCFLAG_FRESTRICTED,
  FUNC_DISPATCH,
  IMPLTYPEFLAG_FDEFAULT,
  IMPLTYPEFLAG_FRESTRICTED,
  IMPLTYPEFLAG_FSOURCE,
  INVOKE_PROPERTYGET,
  INVOKE_PROPERTYPUT,
  INVOKE_PROPERTYPUTREF,
  MEMBERID_NIL,
  PARAMFLAG_FHASDEFAULT,
  TKIND_ALIAS,
  TKIND_COCLASS,
  TKIND_DISPATCH,
  TKIND_ENUM,
  TKIND_INTERFACE,
  TKIND_MODULE,
  TKIND_RECORD,
  TKIND_UNION,
  TYPEFLAG_FCANCREATE,
  TYPEFLAG_FDUAL,
  TYPEFLAG_FDISPATCHABLE,
  VAR_CONST,
  VAR_DISPATCH,
  VAR_PERINSTANCE,
  VT_CARRAY,
  VT_PTR,
  VT_SAFEARRAY,
  VT_USERDEFINED,
} from "../automation/constants.ts";
import { dateToOADate } from "../automation/date.ts";
import { documentation, withFuncDesc, withTypeAttr, withVarDesc } from "../automation/members.ts";
import { ErrorValue, readVariant } from "../automation/variant.ts";
import type { MethodKind } from "../interface.ts";
import { ELEMDESC, ITypeInfo, ITypeLib, PARAMDESCEX, TYPEDESC } from "../interfaces/typeinfo.ts";
import { copyStruct, POINTER_SIZE, readMemory, takeBSTR } from "../low-level/native.ts";
import { COMPointer, sync } from "../pointer.ts";
import { withLibAttr } from "./load.ts";
import type { TypeInfoPointer, TypeLibPointer } from "./load.ts";
import { runSync } from "./steps.ts";
import type { Steps } from "./steps.ts";

// A serializable description of a type library: the input of GetModule's
// runtime construction and of the code generator.

/** A type in a signature, a field or an alias. */
export type TypeRef =
  | { readonly kind: "vt"; readonly vt: number }
  | { readonly kind: "ptr"; readonly to: TypeRef }
  | { readonly kind: "safearray"; readonly of: TypeRef }
  | { readonly kind: "carray"; readonly of: TypeRef; readonly dims: readonly number[] }
  | {
      readonly kind: "user";
      readonly name: string;
      readonly typekind: number;
      readonly guid: string | null;
      /** Name of the library defining an external type; `null` for this library. */
      readonly library: string | null;
      /** External interfaces: whether they are dispatchable (`IDispatch` fallback). */
      readonly dispatchable?: boolean;
    };

/** JSON-safe default values: bigints, dates and SCODEs are encoded. */
export type DefaultValue = null | boolean | number | string | { readonly $bigint: string } | { readonly $date: number } | { readonly $error: number };

export interface ParamModel {
  readonly name: string;
  readonly type: TypeRef;
  /** `PARAMFLAG_*` bits. */
  readonly flags: number;
  readonly default?: DefaultValue;
}

export interface FuncModel {
  readonly name: string;
  readonly memid: number;
  readonly invkind: MethodKind;
  /** `FUNC_DISPATCH` members are called through `IDispatch::Invoke`. */
  readonly dispatch: boolean;
  readonly params: readonly ParamModel[];
  readonly returns: TypeRef;
  /** Vtable slot (`oVft / sizeof(void*)`), `-1` for dispatch members. */
  readonly slot: number;
  readonly hidden: boolean;
  readonly restricted: boolean;
  readonly doc: string | null;
}

export interface PropertyModel {
  readonly name: string;
  readonly memid: number;
  readonly type: TypeRef;
  readonly readonly: boolean;
  readonly doc: string | null;
}

interface BaseModel {
  readonly name: string;
  readonly doc: string | null;
  /** `TYPEFLAG_*` bits. */
  readonly flags: number;
}

export interface EnumModel extends BaseModel {
  readonly kind: "enum";
  readonly members: readonly { readonly name: string; readonly value: number; readonly doc: string | null }[];
}

export interface RecordModel extends BaseModel {
  readonly kind: "record" | "union";
  readonly guid: string | null;
  readonly size: number;
  readonly alignment: number;
  readonly fields: readonly { readonly name: string; readonly type: TypeRef; readonly offset: number; readonly doc: string | null }[];
}

export interface AliasModel extends BaseModel {
  readonly kind: "alias";
  readonly type: TypeRef;
}

export interface InterfaceModel extends BaseModel {
  readonly kind: "interface";
  readonly iid: string;
  /** Base interface (`null` only for IUnknown itself). */
  readonly base: TypeRef | null;
  /** Dual interface: also reachable through `IDispatch`. */
  readonly dual: boolean;
  /** Own vtable methods, in slot order. */
  readonly methods: readonly FuncModel[];
}

export interface DispInterfaceModel extends BaseModel {
  readonly kind: "dispinterface";
  readonly iid: string;
  readonly methods: readonly FuncModel[];
  readonly properties: readonly PropertyModel[];
}

export interface CoClassModel extends BaseModel {
  readonly kind: "coclass";
  readonly clsid: string;
  readonly canCreate: boolean;
  readonly interfaces: readonly { readonly type: TypeRef; readonly default: boolean; readonly source: boolean; readonly restricted: boolean }[];
}

export interface ModuleModel extends BaseModel {
  readonly kind: "module";
  readonly constants: readonly { readonly name: string; readonly value: DefaultValue; readonly doc: string | null }[];
}

export type TypeModel = EnumModel | RecordModel | AliasModel | InterfaceModel | DispInterfaceModel | CoClassModel | ModuleModel;

export interface TypeLibModel {
  readonly name: string;
  readonly doc: string | null;
  readonly libid: string;
  readonly major: number;
  readonly minor: number;
  readonly lcid: number;
  readonly syskind: number;
  readonly types: readonly TypeModel[];
}

// ── reading ───────────────────────────────────────────────────────────

const VARFLAG_FREADONLY = 0x1;
const SIZEOF_TYPEDESC = ctypes.sizeof(TYPEDESC);

interface Context {
  readonly libid: string;
  /** External libraries by LIBID. */
  readonly libraries: Map<string, string>;
  /**
   * References to user types, by the address of the referenced ITypeInfo: the
   * same types are referenced by many members. The pointers are held until the
   * end of the read, so an address cannot be reused by another object.
   */
  readonly userTypes: Map<bigint, { readonly pointer: TypeInfoPointer; readonly type: TypeRef }>;
}

function encodeDefault(value: unknown): DefaultValue | undefined {
  if (value === undefined) return undefined;
  if (value === null || typeof value === "boolean" || typeof value === "number" || typeof value === "string") return value;
  if (typeof value === "bigint") return { $bigint: value.toString() };
  if (value instanceof Date) return { $date: dateToOADate(value) };
  if (value instanceof ErrorValue) return { $error: value.scode };
  return undefined;
}

function containingLibrary(info: TypeInfoPointer, context: Context): { libid: string; name: string } {
  const out = new ctypes.c_void_p();
  const index = new ctypes.c_uint32();
  sync(info).GetContainingTypeLib(ctypes.byref(out), ctypes.byref(index));
  using library = COMPointer.fromAddress(out.value!, ITypeLib);
  const libid = withLibAttr(library, (attr) => attr.guid.toString());
  let name = context.libraries.get(libid);
  if (name === undefined) {
    name = libraryDocumentation(library).name ?? libid;
    context.libraries.set(libid, name);
  }
  return { libid, name };
}

function libraryDocumentation(library: TypeLibPointer): { name: string | null; doc: string | null } {
  const name = new ctypes.c_void_p();
  const doc = new ctypes.c_void_p();
  const context = new ctypes.c_uint32();
  const file = new ctypes.c_void_p();
  sync(library).GetDocumentation(MEMBERID_NIL, ctypes.byref(name), ctypes.byref(doc), ctypes.byref(context), ctypes.byref(file));
  takeBSTR(file.value);
  return { name: takeBSTR(name.value), doc: takeBSTR(doc.value) };
}

function refTypeInfo(info: TypeInfoPointer, hreftype: number): TypeInfoPointer {
  const out = new ctypes.c_void_p();
  sync(info).GetRefTypeInfo(hreftype, ctypes.byref(out));
  return COMPointer.fromAddress(out.value!, ITypeInfo);
}

function readTypeDesc(info: TypeInfoPointer, desc: TYPEDESC, context: Context): TypeRef {
  switch (desc.vt) {
    case VT_PTR:
      return { kind: "ptr", to: readTypeDesc(info, copyStruct(TYPEDESC, desc.lptdesc!), context) };
    case VT_SAFEARRAY:
      return { kind: "safearray", of: readTypeDesc(info, copyStruct(TYPEDESC, desc.lptdesc!), context) };
    case VT_CARRAY: {
      // ARRAYDESC { TYPEDESC tdescElem; USHORT cDims; SAFEARRAYBOUND rgbounds[cDims]; }
      const address = desc.lptdesc!;
      const element = copyStruct(TYPEDESC, address);
      const count = ctypes.readValue(address + BigInt(SIZEOF_TYPEDESC), ctypes.c_uint16);
      const bounds = SIZEOF_TYPEDESC + 4;
      const dims = Array.from({ length: count }, (_, index) => ctypes.readValue(address + BigInt(bounds + index * 8), ctypes.c_uint32));
      return { kind: "carray", of: readTypeDesc(info, element, context), dims };
    }
    case VT_USERDEFINED:
      return userType(info, desc.hreftype, context);
    default:
      return { kind: "vt", vt: desc.vt };
  }
}

function userType(info: TypeInfoPointer, hreftype: number, context: Context): TypeRef {
  const referenced = refTypeInfo(info, hreftype);
  const cached = context.userTypes.get(referenced.address);
  if (cached) {
    referenced.Release();
    return cached.type;
  }
  let type: TypeRef;
  try {
    type = describeUserType(referenced, hreftype, context);
  } catch (error) {
    referenced.Release();
    throw error;
  }
  context.userTypes.set(referenced.address, { pointer: referenced, type });
  return type;
}

function describeUserType(referenced: TypeInfoPointer, hreftype: number, context: Context): TypeRef {
  const name = documentation(referenced).name ?? `Type${hreftype}`;
  const library = containingLibrary(referenced, context);
  const external = library.libid !== context.libid;
  return withTypeAttr(referenced, (attr): TypeRef => {
    // External aliases are replaced by their definition: no other module is needed.
    if (external && attr.typekind === TKIND_ALIAS) return readTypeDesc(referenced, attr.tdescAlias, context);
    return {
      kind: "user",
      name,
      typekind: attr.typekind,
      guid: attr.guid.isNull() ? null : attr.guid.toString(),
      library: external ? library.name : null,
      ...(external && (attr.typekind === TKIND_INTERFACE || attr.typekind === TKIND_DISPATCH) ? { dispatchable: attr.typekind === TKIND_DISPATCH || (attr.wTypeFlags & (TYPEFLAG_FDUAL | TYPEFLAG_FDISPATCHABLE)) !== 0 } : {}),
    };
  });
}

function memberNames(info: TypeInfoPointer, memid: number, count: number): string[] {
  const names = Buffer.alloc(POINTER_SIZE * count);
  const found = new ctypes.c_uint32();
  sync(info).GetNames(memid, names, count, ctypes.byref(found));
  return Array.from({ length: found.value }, (_, index) => takeBSTR(ctypes.readValue(names, ctypes.c_void_p, index * POINTER_SIZE)) ?? "");
}

const INVOKE_KINDS: Record<number, MethodKind> = { 1: "method", [INVOKE_PROPERTYGET]: "propget", [INVOKE_PROPERTYPUT]: "propput", [INVOKE_PROPERTYPUTREF]: "propputref" };

function readFunctions(info: TypeInfoPointer, count: number, context: Context): FuncModel[] {
  const functions: FuncModel[] = [];
  for (let index = 0; index < count; index++) {
    functions.push(
      withFuncDesc(info, index, (desc) => {
        const names = memberNames(info, desc.memid, desc.cParams + 1);
        const elementSize = ctypes.sizeof(ELEMDESC);
        const params = Array.from({ length: desc.cParams }, (_, position): ParamModel => {
          const element = copyStruct(ELEMDESC, desc.lprgelemdescParam! + BigInt(position * elementSize));
          const flags = element.paramdesc.wParamFlags;
          const param: ParamModel = { name: names[position + 1] || (position === desc.cParams - 1 && desc.invkind & (INVOKE_PROPERTYPUT | INVOKE_PROPERTYPUTREF) ? "value" : `arg${position}`), type: readTypeDesc(info, element.tdesc, context), flags };
          if (flags & PARAMFLAG_FHASDEFAULT && element.paramdesc.pparamdescex) {
            const value = encodeDefault(readVariant(copyStruct(PARAMDESCEX, element.paramdesc.pparamdescex).varDefaultValue._buffer));
            if (value !== undefined) return { ...param, default: value };
          }
          return param;
        });
        const doc = documentation(info, desc.memid);
        return {
          name: names[0],
          memid: desc.memid,
          invkind: INVOKE_KINDS[desc.invkind] ?? "method",
          dispatch: desc.funckind === FUNC_DISPATCH,
          params,
          returns: readTypeDesc(info, desc.elemdescFunc.tdesc, context),
          slot: desc.funckind === FUNC_DISPATCH ? -1 : desc.oVft / POINTER_SIZE,
          hidden: (desc.wFuncFlags & FUNCFLAG_FHIDDEN) !== 0,
          restricted: (desc.wFuncFlags & FUNCFLAG_FRESTRICTED) !== 0,
          doc: doc.docString,
        };
      }),
    );
  }
  return functions;
}

function implementedTypes(info: TypeInfoPointer, count: number, context: Context) {
  return Array.from({ length: count }, (_, index) => {
    const reference = new ctypes.c_uint32();
    const flags = new ctypes.c_int32();
    sync(info).GetRefTypeOfImplType(index, ctypes.byref(reference));
    sync(info).GetImplTypeFlags(index, ctypes.byref(flags));
    return { type: userType(info, reference.value, context), flags: flags.value };
  });
}

// IUnknown / IDispatch members repeated in dispinterface views.
function isInheritedDispatchMember(func: FuncModel): boolean {
  return ["QueryInterface", "AddRef", "Release", "GetTypeInfoCount", "GetTypeInfo", "GetIDsOfNames", "Invoke"].includes(func.name) && func.memid >= 0x60000000 && func.memid <= 0x60010003;
}

function readType(info: TypeInfoPointer, context: Context): TypeModel | null {
  const { name, docString } = documentation(info);
  const attr = withTypeAttr(info, (value) => ({
    typekind: value.typekind,
    guid: value.guid.toString(),
    flags: value.wTypeFlags,
    functions: value.cFuncs,
    variables: value.cVars,
    implemented: value.cImplTypes,
    size: value.cbSizeInstance,
    alignment: value.cbAlignment,
    alias: value.typekind === TKIND_ALIAS ? readTypeDesc(info, value.tdescAlias, context) : null,
  }));
  const base = { name: name ?? attr.guid, doc: docString, flags: attr.flags };
  switch (attr.typekind) {
    case TKIND_ENUM: {
      const members = Array.from({ length: attr.variables }, (_, index) =>
        withVarDesc(info, index, (desc) => ({ name: memberNames(info, desc.memid, 1)[0], value: Number(readVariant(readMemory(desc.lpvarValue!, 8 + 2 * POINTER_SIZE))), doc: documentation(info, desc.memid).docString })),
      );
      return { kind: "enum", ...base, members };
    }
    case TKIND_RECORD:
    case TKIND_UNION: {
      const fields = Array.from({ length: attr.variables }, (_, index) =>
        withVarDesc(info, index, (desc) => ({
          name: memberNames(info, desc.memid, 1)[0],
          type: readTypeDesc(info, desc.elemdescVar.tdesc, context),
          offset: desc.varkind === VAR_PERINSTANCE ? desc.oInst : 0,
          doc: documentation(info, desc.memid).docString,
        })),
      );
      return { kind: attr.typekind === TKIND_RECORD ? "record" : "union", ...base, guid: attr.guid === "{00000000-0000-0000-0000-000000000000}" ? null : attr.guid, size: attr.size, alignment: attr.alignment, fields };
    }
    case TKIND_ALIAS:
      return { kind: "alias", ...base, type: attr.alias! };
    case TKIND_INTERFACE: {
      const implemented = implementedTypes(info, attr.implemented, context);
      const methods = readFunctions(info, attr.functions, context).sort((a, b) => a.slot - b.slot);
      return { kind: "interface", ...base, iid: attr.guid, base: implemented[0]?.type ?? null, dual: (attr.flags & TYPEFLAG_FDUAL) !== 0, methods };
    }
    case TKIND_DISPATCH: {
      if (attr.flags & TYPEFLAG_FDUAL) {
        // Dual: describe the vtable interface (implemented type -1).
        const reference = new ctypes.c_uint32();
        sync(info).GetRefTypeOfImplType(-1 >>> 0, ctypes.byref(reference));
        using vtable = refTypeInfo(info, reference.value);
        const model = readType(vtable, context);
        return model && model.kind === "interface" ? { ...model, dual: true, doc: model.doc ?? docString } : model;
      }
      const methods = readFunctions(info, attr.functions, context).filter((func) => !isInheritedDispatchMember(func));
      const properties = Array.from({ length: attr.variables }, (_, index) =>
        withVarDesc(info, index, (desc) =>
          desc.varkind === VAR_DISPATCH
            ? { name: memberNames(info, desc.memid, 1)[0], memid: desc.memid, type: readTypeDesc(info, desc.elemdescVar.tdesc, context), readonly: (desc.wVarFlags & VARFLAG_FREADONLY) !== 0, doc: documentation(info, desc.memid).docString }
            : null,
        ),
      ).filter((property): property is PropertyModel => property !== null);
      return { kind: "dispinterface", ...base, iid: attr.guid, methods, properties };
    }
    case TKIND_COCLASS: {
      const interfaces = implementedTypes(info, attr.implemented, context).map(({ type, flags }) => ({
        type,
        default: (flags & IMPLTYPEFLAG_FDEFAULT) !== 0,
        source: (flags & IMPLTYPEFLAG_FSOURCE) !== 0,
        restricted: (flags & IMPLTYPEFLAG_FRESTRICTED) !== 0,
      }));
      return { kind: "coclass", ...base, clsid: attr.guid, canCreate: (attr.flags & TYPEFLAG_FCANCREATE) !== 0, interfaces };
    }
    case TKIND_MODULE: {
      const constants = Array.from({ length: attr.variables }, (_, index) =>
        withVarDesc(info, index, (desc) =>
          desc.varkind === VAR_CONST ? { name: memberNames(info, desc.memid, 1)[0], value: encodeDefault(readVariant(readMemory(desc.lpvarValue!, 8 + 2 * POINTER_SIZE))) ?? null, doc: documentation(info, desc.memid).docString } : null,
        ),
      ).filter((constant): constant is ModuleModel["constants"][number] => constant !== null);
      return { kind: "module", ...base, constants };
    }
    default:
      return null;
  }
}

/** Describe a loaded type library. */
export function readTypeLib(library: TypeLibPointer): TypeLibModel {
  return runSync(readTypeLibSteps(library));
}

/** {@link readTypeLib} in steps: yields after each type. */
export function* readTypeLibSteps(library: TypeLibPointer): Steps<TypeLibModel> {
  const attributes = withLibAttr(library, (attr) => ({ libid: attr.guid.toString(), major: attr.wMajorVerNum, minor: attr.wMinorVerNum, lcid: attr.lcid, syskind: attr.syskind }));
  const { name, doc } = libraryDocumentation(library);
  const context: Context = { libid: attributes.libid, libraries: new Map([[attributes.libid, name ?? attributes.libid]]), userTypes: new Map() };
  const types: TypeModel[] = [];
  const seen = new Set<string>();
  try {
    const count = sync(library).GetTypeInfoCount();
    for (let index = 0; index < count; index++) {
      const out = new ctypes.c_void_p();
      sync(library).GetTypeInfo(index, ctypes.byref(out));
      {
        using info = COMPointer.fromAddress(out.value!, ITypeInfo);
        const type = readType(info, context);
        if (type) {
          const key = "iid" in type ? `${type.kind === "dispinterface" ? "D" : "I"}${type.iid}` : `${type.kind}:${type.name}`;
          if (!seen.has(key) && !("iid" in type && seen.has(`I${type.iid}`))) {
            seen.add(key);
            types.push(type);
          }
        }
      }
      yield;
    }
  } finally {
    for (const { pointer } of context.userTypes.values()) pointer.Release();
  }
  return { name: name ?? attributes.libid, doc, ...attributes, types };
}
