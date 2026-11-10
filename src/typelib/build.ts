import * as ctypes from "node-ctypes";
import {
  PARAMFLAG_FIN,
  PARAMFLAG_FLCID,
  PARAMFLAG_FOPT,
  PARAMFLAG_FOUT,
  PARAMFLAG_FRETVAL,
  TKIND_ALIAS,
  TKIND_COCLASS,
  TKIND_DISPATCH,
  TKIND_ENUM,
  TKIND_INTERFACE,
  TKIND_RECORD,
  TKIND_UNION,
  VT_BOOL,
  VT_BSTR,
  VT_CY,
  VT_DATE,
  VT_DECIMAL,
  VT_DISPATCH,
  VT_ERROR,
  VT_HRESULT,
  VT_I1,
  VT_I2,
  VT_I4,
  VT_I8,
  VT_INT,
  VT_INT_PTR,
  VT_LPSTR,
  VT_LPWSTR,
  VT_R4,
  VT_R8,
  VT_RECORD,
  VT_UI1,
  VT_UI2,
  VT_UI4,
  VT_UI8,
  VT_UINT,
  VT_UINT_PTR,
  VT_UNKNOWN,
  VT_VARIANT,
  VT_VOID,
} from "../automation/constants.ts";
import { oaDateToDate } from "../automation/date.ts";
import { DECIMAL } from "../automation/decimal.ts";
import { ErrorValue, VARIANT } from "../automation/variant.ts";
import { BSTR, COMMETHOD, CURRENCY, DATE, LPOLESTR, REFGUID, SAFEARRAY, dispid, lazy } from "../commethod.ts";
import type { ComParamType, IdlFlag, ParamFlag, ParamSpec } from "../commethod.ts";
import { DISPMETHOD, DISPPROPERTY } from "../dispmethod.ts";
import { GUID } from "../guid.ts";
import { COMInterface, HRESULT, STDMETHOD, findInterface } from "../interface.ts";
import type { InterfaceDef, MethodSpec } from "../interface.ts";
import * as builtinInterfaces from "../interfaces/index.ts";
import { IDispatch } from "../interfaces/dispatch.ts";
import { IUnknown } from "../interfaces/unknown.ts";
import { setRecordTypeInfo, setTypeLibLocation } from "./coclass.ts";
import { loadTypeLib } from "./load.ts";
import { ITypeLib } from "../interfaces/typeinfo.ts";
import type { InterfacePointer } from "../interface.ts";
import type { CoClass, TypeLibLocation } from "./coclass.ts";
import type { AliasModel, DefaultValue, DispInterfaceModel, FuncModel, InterfaceModel, ParamModel, RecordModel, TypeLibModel, TypeModel, TypeRef } from "./model.ts";

/** Interfaces defined by node-comtypes, by IID: type libraries reuse them. */
export const builtinInterfaceDefs: ReadonlyMap<string, { readonly name: string; readonly definition: InterfaceDef<any> }> = new Map(
  Object.entries(builtinInterfaces)
    .filter((entry): entry is [string, InterfaceDef<any>] => typeof entry[1] === "object" && entry[1] !== null && "Pointer" in entry[1] && "iid" in entry[1])
    .map(([name, definition]) => [definition.iid.toString(), { name, definition }]),
);

/** A module built from a type library: its types by name, plus `library`. */
export interface TypeLibModule {
  readonly library: Omit<TypeLibModel, "types"> & { readonly guid: GUID; readonly path?: string };
  readonly [name: string]: unknown;
}

export function decodeDefault(value: DefaultValue | undefined): unknown {
  if (value === null || value === undefined || typeof value !== "object") return value;
  if ("$bigint" in value) return BigInt(value.$bigint);
  if ("$date" in value) return oaDateToDate(value.$date);
  return new ErrorValue(value.$error);
}

/** Types of a model by name, with alias resolution. */
export class ModelIndex {
  readonly model: TypeLibModel;
  readonly byName: ReadonlyMap<string, TypeModel>;

  constructor(model: TypeLibModel) {
    this.model = model;
    this.byName = new Map(model.types.map((type) => [type.name, type]));
  }

  /** Follow aliases of this library. */
  resolve(ref: TypeRef): TypeRef {
    let current = ref;
    for (let depth = 0; depth < 32 && current.kind === "user" && current.library === null && current.typekind === TKIND_ALIAS; depth++) {
      const alias = this.byName.get(current.name) as AliasModel | undefined;
      if (!alias) break;
      current = alias.type;
    }
    return current;
  }

  /** The model of a user type of this library. */
  local(ref: TypeRef): TypeModel | undefined {
    return ref.kind === "user" && ref.library === null ? this.byName.get(ref.name) : undefined;
  }
}

const SIMPLE: Record<number, ctypes.FFIType> = {
  [VT_I1]: ctypes.c_int8,
  [VT_UI1]: ctypes.c_uint8,
  [VT_I2]: ctypes.c_int16,
  [VT_UI2]: ctypes.c_uint16,
  [VT_I4]: ctypes.c_int32,
  [VT_INT]: ctypes.c_int32,
  [VT_ERROR]: ctypes.c_int32,
  [VT_HRESULT]: ctypes.c_int32,
  [VT_UI4]: ctypes.c_uint32,
  [VT_UINT]: ctypes.c_uint32,
  [VT_I8]: ctypes.c_int64,
  [VT_UI8]: ctypes.c_uint64,
  [VT_R4]: ctypes.c_float,
  [VT_R8]: ctypes.c_double,
  [VT_INT_PTR]: ctypes.c_ssize_t,
  [VT_UINT_PTR]: ctypes.c_size_t,
  [VT_BOOL]: ctypes.wintypes.VARIANT_BOOL,
};

type UserRef = Extract<TypeRef, { kind: "user" }>;

function isInterfaceRef(ref: TypeRef): boolean {
  return ref.kind === "user" && (ref.typekind === TKIND_INTERFACE || ref.typekind === TKIND_DISPATCH);
}

const libraries = new Map<string, InterfacePointer<typeof ITypeLib>>();

/** An owned `ITypeInfo` address for type `guid` of a library (loaded once). */
function typeInfoAddress(location: TypeLibLocation, guid: string): bigint {
  const key = `${location.libid}:${location.major}.${location.minor}:${location.lcid}`;
  let library = libraries.get(key);
  if (!library || library.released) {
    library = location.path ? loadTypeLib(location.path) : loadTypeLib({ libid: location.libid, major: location.major, minor: location.minor, lcid: location.lcid });
    libraries.set(key, library);
  }
  const out = new ctypes.c_void_p();
  library.GetTypeInfoOfGuid(new GUID(guid), ctypes.byref(out));
  return out.value!;
}

/** Builds runtime definitions from a model (shared by GetModule and generated modules). */
class Builder {
  readonly index: ModelIndex;
  readonly location: TypeLibLocation;
  readonly interfaces = new Map<string, InterfaceDef<any>>();
  readonly records = new Map<string, ctypes.StructureClass>();
  readonly #building = new Set<string>();

  constructor(model: TypeLibModel, location: TypeLibLocation) {
    this.index = new ModelIndex(model);
    this.location = location;
  }

  /** Interface definition of a user type (built-in, local, already loaded, or a fallback). */
  interfaceOf(ref: Extract<TypeRef, { kind: "user" }>): InterfaceDef<any> {
    const builtin = ref.guid ? builtinInterfaceDefs.get(ref.guid) : undefined;
    if (builtin) return builtin.definition;
    if (ref.library === null) {
      const local = this.index.byName.get(ref.name);
      if (local?.kind === "coclass") {
        const fallback = local.interfaces.find((entry) => entry.default && !entry.source) ?? local.interfaces[0];
        return fallback && fallback.type.kind === "user" ? this.interfaceOf(fallback.type) : IDispatch;
      }
      return this.interface(ref.name);
    }
    const loaded = ref.guid ? findInterface(ref.guid) : undefined;
    return loaded ?? (ref.dispatchable || ref.typekind === TKIND_DISPATCH ? IDispatch : IUnknown);
  }

  /** Parameter type for an interface reference: built-ins directly, local ones lazily (mutual references). */
  interfaceParam(ref: Extract<TypeRef, { kind: "user" }>): ComParamType {
    if (ref.library === null && !(ref.guid && builtinInterfaceDefs.has(ref.guid))) return lazy(() => this.interfaceOf(ref));
    return this.interfaceOf(ref);
  }

  /** Native layout type (record fields, pointer targets). */
  ffi(ref: TypeRef): ctypes.FFIType {
    const resolved = this.index.resolve(ref);
    switch (resolved.kind) {
      case "vt":
        if (resolved.vt === VT_VARIANT) return VARIANT;
        if (resolved.vt === VT_DATE) return ctypes.c_double;
        if (resolved.vt === VT_CY) return ctypes.c_int64;
        if (resolved.vt === VT_DECIMAL) return DECIMAL;
        return SIMPLE[resolved.vt] ?? ctypes.c_void_p;
      case "carray":
        return ctypes.array(
          this.ffi(resolved.of) as never,
          resolved.dims.reduce((total, dim) => total * dim, 1),
        ) as unknown as ctypes.FFIType;
      case "user": {
        if (resolved.typekind === TKIND_ENUM) return ctypes.c_int32;
        if (resolved.typekind === TKIND_RECORD || resolved.typekind === TKIND_UNION) return this.record(resolved);
        return ctypes.c_void_p;
      }
      default:
        return ctypes.c_void_p;
    }
  }

  /** Record class (built on first use; `GUID` maps to the node-comtypes GUID). */
  record(ref: Extract<TypeRef, { kind: "user" }>): ctypes.StructureClass {
    if (ref.name === "GUID") return GUID;
    const existing = this.records.get(ref.name);
    if (existing) return existing;
    const model = this.index.local(ref) as RecordModel | undefined;
    if (!model || (model.kind !== "record" && model.kind !== "union")) return VARIANT;
    if (this.#building.has(ref.name)) throw new TypeError(`Recursive record ${ref.name}`);
    this.#building.add(ref.name);
    try {
      const fields = model.fields.map((field) => [field.name, this.ffi(field.type)] as const);
      const Base = model.kind === "union" ? ctypes.Union : ctypes.Structure;
      const type = {
        [model.name]: class extends Base {
          static _fields_ = fields;
        },
      }[model.name];
      if (ctypes.sizeof(type) !== model.size) Object.defineProperty(type, "_pack_", { value: model.alignment });
      this.records.set(model.name, type);
      // Records with a GUID can be VT_RECORD values (IRecordInfo).
      if (model.guid) {
        const location = this.location;
        const guid = model.guid;
        setRecordTypeInfo(type, () => typeInfoAddress(location, guid));
      }
      return type;
    } finally {
      this.#building.delete(ref.name);
    }
  }

  /** Type of a value (an `in` value, an `out` target, a property). */
  value(ref: TypeRef): ComParamType {
    const resolved = this.index.resolve(ref);
    switch (resolved.kind) {
      case "vt":
        switch (resolved.vt) {
          case VT_BSTR:
            return BSTR;
          case VT_VARIANT:
            return VARIANT;
          case VT_DATE:
            return DATE;
          case VT_CY:
            return CURRENCY;
          case VT_DECIMAL:
            return DECIMAL;
          case VT_LPWSTR:
            return LPOLESTR;
          case VT_LPSTR:
            return ctypes.c_char_p;
          case VT_UNKNOWN:
            return IUnknown;
          case VT_DISPATCH:
            return IDispatch;
          default:
            return SIMPLE[resolved.vt] ?? ctypes.c_void_p;
        }
      case "ptr": {
        const target = this.index.resolve(resolved.to);
        if (isInterfaceRef(target)) return this.interfaceParam(target as UserRef);
        if (target.kind === "user" && target.typekind === TKIND_COCLASS) return this.interfaceOf(target);
        if (target.kind === "user" && target.name === "GUID") return REFGUID;
        if (target.kind === "vt" && target.vt === VT_VOID) return ctypes.c_void_p;
        return ctypes.POINTER(this.ffi(target) as never);
      }
      case "safearray": {
        const vt = this.elementVartype(resolved.of);
        return vt === VT_RECORD ? ctypes.c_void_p : SAFEARRAY(vt);
      }
      case "carray":
        return ctypes.c_void_p;
      case "user":
        if (isInterfaceRef(resolved)) return this.interfaceParam(resolved as UserRef);
        if (resolved.typekind === TKIND_COCLASS) return this.interfaceOf(resolved);
        return this.ffi(resolved);
    }
  }

  elementVartype(ref: TypeRef): number {
    const resolved = this.index.resolve(ref);
    if (resolved.kind === "vt") return resolved.vt;
    if (resolved.kind === "ptr") return this.elementVartype(resolved.to);
    if (resolved.kind === "user") {
      if (resolved.typekind === TKIND_ENUM) return VT_I4;
      if (isInterfaceRef(resolved))
        return resolved.typekind === TKIND_DISPATCH || resolved.dispatchable || this.index.local(resolved)?.kind === "dispinterface" || (this.index.local(resolved) as InterfaceModel | undefined)?.dual ? VT_DISPATCH : VT_UNKNOWN;
      return VT_RECORD;
    }
    return VT_VARIANT;
  }

  /** COMMETHOD parameter spec. */
  param(param: ParamModel): ParamSpec {
    const flags: ParamFlag[] = [];
    if (param.flags & PARAMFLAG_FIN) flags.push("in");
    if (param.flags & PARAMFLAG_FOUT) flags.push("out");
    if (param.flags & PARAMFLAG_FRETVAL) flags.push("retval");
    if (param.flags & PARAMFLAG_FLCID) flags.push("lcid");
    if (param.flags & PARAMFLAG_FOPT) flags.push("optional");
    const output = (param.flags & PARAMFLAG_FOUT) !== 0;
    const type = param.flags & PARAMFLAG_FLCID ? ctypes.c_uint32 : output ? (param.type.kind === "ptr" ? this.value(param.type.to) : ctypes.c_void_p) : this.value(param.type);
    return param.default !== undefined ? [flags, type, param.name, decodeDefault(param.default)] : [flags, type, param.name];
  }

  restype(ref: TypeRef): ctypes.FFIType | null {
    const resolved = this.index.resolve(ref);
    if (resolved.kind === "vt" && resolved.vt === VT_HRESULT) return HRESULT;
    if (resolved.kind === "vt" && resolved.vt === VT_VOID) return null;
    return this.ffi(resolved);
  }

  idlflags(func: FuncModel): IdlFlag[] {
    const flags: IdlFlag[] = [];
    if (func.invkind !== "method") flags.push(func.invkind);
    flags.push(dispid(func.memid));
    if (func.hidden) flags.push("hidden");
    if (func.restricted) flags.push("restricted");
    return flags;
  }

  interface(name: string): InterfaceDef<any> {
    const existing = this.interfaces.get(name);
    if (existing) return existing;
    const model = this.index.byName.get(name);
    if (!model || (model.kind !== "interface" && model.kind !== "dispinterface")) return IDispatch;
    const builtin = builtinInterfaceDefs.get(model.iid);
    if (builtin) {
      this.interfaces.set(name, builtin.definition);
      return builtin.definition;
    }
    if (this.#building.has(name)) throw new TypeError(`Recursive base interface ${name}`);
    this.#building.add(name);
    try {
      const definition = model.kind === "interface" ? this.vtableInterface(model) : this.dispInterface(model);
      setTypeLibLocation(definition, this.location);
      this.interfaces.set(name, definition);
      return definition;
    } finally {
      this.#building.delete(name);
    }
  }

  vtableInterface(model: InterfaceModel): InterfaceDef<any> {
    const base = model.base && model.base.kind === "user" ? this.interfaceOf(model.base) : model.iid === IUnknown.iid.toString() ? null : IUnknown;
    const methods: MethodSpec[] = [];
    const used = new Set<string>();
    let slot = base?.allMethods.length ?? 0;
    for (const func of model.methods) {
      if (func.slot < slot) continue;
      // Keep the vtable layout when the library skips slots.
      while (slot < func.slot) methods.push(STDMETHOD(HRESULT, `_slot${slot++}`));
      let spec = COMMETHOD(this.idlflags(func), this.restype(func.returns), func.name, ...func.params.map((param) => this.param(param))) as MethodSpec;
      if (used.has(spec.name)) spec = COMMETHOD(this.idlflags(func), this.restype(func.returns), `${func.name}_${slot}`, ...func.params.map((param) => this.param(param))) as MethodSpec;
      used.add(spec.name);
      methods.push(spec);
      slot++;
    }
    return COMInterface({ name: model.name, iid: model.iid, extends: base ?? undefined, methods });
  }

  dispInterface(model: DispInterfaceModel): InterfaceDef<any> {
    const members: MethodSpec[] = [];
    const used = new Set<string>();
    const add = (spec: MethodSpec) => {
      if (used.has(spec.name)) return;
      used.add(spec.name);
      members.push(spec);
    };
    for (const property of model.properties) add(DISPPROPERTY([dispid(property.memid), ...(property.readonly ? ["readonly"] : [])], this.value(property.type), property.name) as MethodSpec);
    for (const func of model.methods) {
      const returns = this.index.resolve(func.returns);
      const restype = returns.kind === "vt" && (returns.vt === VT_VOID || returns.vt === VT_HRESULT) ? null : this.value(returns);
      const params = func.params.map((param): ParamSpec => {
        const flags: ParamFlag[] = param.flags & PARAMFLAG_FOPT || param.default !== undefined ? ["in", "optional"] : ["in"];
        const target = param.type.kind === "ptr" && !isInterfaceRef(this.index.resolve(param.type.to)) ? param.type.to : param.type;
        return param.default !== undefined ? [flags, this.value(target), param.name, decodeDefault(param.default)] : [flags, this.value(target), param.name];
      });
      add(DISPMETHOD(func.invkind === "method" ? [dispid(func.memid)] : [func.invkind, dispid(func.memid)], restype, func.name, ...params) as MethodSpec);
    }
    return COMInterface({ name: model.name, iid: model.iid, extends: IDispatch, dispMethods: members });
  }

  coclass(model: Extract<TypeModel, { kind: "coclass" }>): CoClass {
    const interfaces = model.interfaces.filter((entry) => !entry.source && entry.type.kind === "user").map((entry) => ({ entry, definition: this.interfaceOf(entry.type as Extract<TypeRef, { kind: "user" }>) }));
    const sources = model.interfaces.filter((entry) => entry.source && entry.type.kind === "user").map((entry) => ({ entry, definition: this.interfaceOf(entry.type as Extract<TypeRef, { kind: "user" }>) }));
    return Object.freeze({
      kind: "coclass",
      name: model.name,
      clsid: new GUID(model.clsid),
      defaultInterface: (interfaces.find((item) => item.entry.default) ?? interfaces[0])?.definition ?? IDispatch,
      sourceInterface: (sources.find((item) => item.entry.default) ?? sources[0])?.definition ?? null,
      interfaces: Object.freeze(interfaces.map((item) => item.definition)),
      sourceInterfaces: Object.freeze(sources.map((item) => item.definition)),
      canCreate: model.canCreate,
    });
  }
}

/**
 * Build the runtime definitions of a type library model: enums (frozen
 * objects), records (`Structure` classes), interfaces (COMMETHOD),
 * dispinterfaces (DISPMETHOD / DISPPROPERTY), coclasses and module constants.
 */
export function buildModule(model: TypeLibModel, options: { readonly path?: string } = {}): TypeLibModule {
  const location: TypeLibLocation = { libid: model.libid, major: model.major, minor: model.minor, lcid: model.lcid, ...(options.path ? { path: options.path } : {}) };
  const builder = new Builder(model, location);
  const { types: _types, ...library } = model;
  const module: Record<string, unknown> = { library: Object.freeze({ ...library, guid: new GUID(model.libid), ...(options.path ? { path: options.path } : {}) }) };
  const define = (name: string, value: unknown) => {
    if (!(name in module)) module[name] = value;
  };
  for (const type of model.types) {
    switch (type.kind) {
      case "enum":
        define(type.name, Object.freeze(Object.fromEntries(type.members.map((member) => [member.name, member.value]))));
        break;
      case "record":
      case "union":
        define(type.name, builder.record({ kind: "user", name: type.name, typekind: type.kind === "record" ? TKIND_RECORD : TKIND_UNION, guid: type.guid, library: null }));
        break;
      case "interface":
      case "dispinterface":
        define(type.name, builder.interface(type.name));
        break;
      case "coclass":
        define(type.name, builder.coclass(type));
        break;
      case "module":
        for (const constant of type.constants) define(constant.name, decodeDefault(constant.value));
        break;
    }
  }
  // Aliases of local enums, records and interfaces name the same definition
  // (MIDL often names the enum itself __MIDL___MIDL_itf_...).
  for (const type of model.types) {
    if (type.kind !== "alias") continue;
    const target = builder.index.resolve(type.type);
    if (target.kind === "user" && target.library === null && target.name in module) define(type.name, module[target.name]);
  }
  return Object.freeze(module) as TypeLibModule;
}
