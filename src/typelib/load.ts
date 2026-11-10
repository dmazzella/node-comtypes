import { resolve as resolvePath } from "node:path";
import * as ctypes from "node-ctypes";
import { dispatchTypeInfo } from "../automation/members.ts";
import { borrowReference } from "../automation/variant.ts";
import type { ComObject } from "../commethod.ts";
import { COMError, checkHRESULT } from "../errors.ts";
import { GUID, toGUID } from "../guid.ts";
import type { GUIDLike } from "../guid.ts";
import { TYPE_E_CANTLOADLIBRARY } from "../hresult.ts";
import type { InterfacePointer } from "../interface.ts";
import { IDispatch } from "../interfaces/dispatch.ts";
import { ICreateTypeLib2, IProvideClassInfo, ITypeInfo, ITypeLib, TLIBATTR } from "../interfaces/typeinfo.ts";
import { ole32, oleaut32, takeBSTR } from "../low-level/native.ts";
import { HKEY_CLASSES_ROOT, readString, subkeys } from "../low-level/registry.ts";
import { SYS_WIN32, SYS_WIN64 } from "../automation/constants.ts";
import { POINTER_SIZE } from "../low-level/native.ts";
import { COMPointer } from "../pointer.ts";

export type TypeLibPointer = InterfacePointer<typeof ITypeLib>;
export type TypeInfoPointer = InterfacePointer<typeof ITypeInfo>;

/** A registered type library: LIBID, version (default: the highest registered) and LCID. */
export interface RegisteredTypeLib {
  readonly libid: GUIDLike;
  readonly major?: number;
  readonly minor?: number;
  readonly lcid?: number;
}

/** Anything that identifies a type library. */
export type TypeLibSource = string | RegisteredTypeLib | TypeLibPointer;

const REGKIND_NONE = 2;

export interface TypeLibVersion {
  readonly major: number;
  readonly minor: number;
}

/** Registered versions of a type library (`HKCR\TypeLib\{libid}`), highest first. */
export function registeredVersions(libid: GUIDLike): TypeLibVersion[] {
  return subkeys(HKEY_CLASSES_ROOT, `TypeLib\\${toGUID(libid)}`)
    .map((name) => /^([0-9a-f]+)\.([0-9a-f]+)$/i.exec(name))
    .filter((match): match is RegExpExecArray => match !== null)
    .map((match) => ({ major: parseInt(match[1], 16), minor: parseInt(match[2], 16) }))
    .sort((a, b) => b.major - a.major || b.minor - a.minor);
}

function versionOf(source: RegisteredTypeLib): TypeLibVersion {
  if (source.major !== undefined) return { major: source.major, minor: source.minor ?? 0 };
  const [highest] = registeredVersions(source.libid);
  if (!highest) throw new COMError(TYPE_E_CANTLOADLIBRARY, { text: `Type library ${toGUID(source.libid)} is not registered` });
  return highest;
}

/** `LoadRegTypeLib`: a registered type library. */
export function loadRegisteredTypeLib(source: RegisteredTypeLib): TypeLibPointer {
  const { major, minor } = versionOf(source);
  const out = new ctypes.c_void_p();
  checkHRESULT(oleaut32().LoadRegTypeLib(toGUID(source.libid)._buffer, major, minor, source.lcid ?? 0, ctypes.byref(out)), "LoadRegTypeLib");
  return COMPointer.fromAddress(out.value!, ITypeLib);
}

/** File of a registered type library (`QueryPathOfRegTypeLib`). */
export function registeredTypeLibPath(source: RegisteredTypeLib): string {
  const { major, minor } = versionOf(source);
  const out = new ctypes.c_void_p();
  checkHRESULT(oleaut32().QueryPathOfRegTypeLib(toGUID(source.libid)._buffer, major, minor, source.lcid ?? 0, ctypes.byref(out)), "QueryPathOfRegTypeLib");
  return (takeBSTR(out.value) ?? "").replace(/\0+$/, "");
}

/** The type library registered for a class (`HKCR\CLSID\{clsid}\TypeLib` and `Version`). */
export function typeLibOfClass(clsid: GUIDLike): TypeLibPointer {
  const key = `CLSID\\${toGUID(clsid)}`;
  const libid = readString(HKEY_CLASSES_ROOT, `${key}\\TypeLib`);
  if (!libid) throw new COMError(TYPE_E_CANTLOADLIBRARY, { text: `No type library registered for ${toGUID(clsid)}` });
  const version = /^(\d+)\.(\d+)$/.exec(readString(HKEY_CLASSES_ROOT, `${key}\\Version`) ?? "");
  // CLSID\{clsid}\Version is decimal; TypeLib\{libid}\x.y keys are hexadecimal.
  const registered = version ? registeredVersions(libid).find((candidate) => candidate.major === Number(version[1])) : undefined;
  return loadRegisteredTypeLib({ libid, major: registered?.major, minor: registered?.minor });
}

/** The type library of a live object, through `IProvideClassInfo` or `IDispatch::GetTypeInfo`. */
export function typeLibOf(object: COMPointer | ComObject): TypeLibPointer {
  const borrowed = borrowReference(object);
  if (!borrowed) throw new TypeError("typeLibOf expects a COM object");
  using pointer = COMPointer.fromAddress(borrowed.address, borrowed.interface, { addRef: true });
  let typeInfo: TypeInfoPointer | null = null;
  const provider = pointer.tryQueryInterface(IProvideClassInfo);
  if (provider) {
    using owned = provider;
    try {
      typeInfo = owned.GetClassInfo();
    } catch {
      typeInfo = null;
    }
  }
  if (!typeInfo) {
    const dispatch = pointer.tryQueryInterface(IDispatch);
    if (dispatch) {
      using owned = dispatch;
      typeInfo = dispatchTypeInfo(owned);
    }
  }
  if (!typeInfo) throw new COMError(TYPE_E_CANTLOADLIBRARY, { text: "The object provides no type information" });
  using info = typeInfo;
  const out = new ctypes.c_void_p();
  const index = new ctypes.c_uint32();
  info.GetContainingTypeLib(ctypes.byref(out), ctypes.byref(index));
  return COMPointer.fromAddress(out.value!, ITypeLib);
}

const FILE_EXTENSION = /\.(dll|tlb|olb|ocx|exe)$/i;

/**
 * Load a type library from:
 * - a file (`LoadTypeLibEx`, searched like `LoadLibrary`: `"scrrun.dll"`, a path, `"file.dll\\2"` for a resource index);
 * - a LIBID string or `{ libid, major?, minor?, lcid? }` (registered; default the highest version);
 * - a ProgID (`"Scripting.Dictionary"`): the type library registered for its class;
 * - an `ITypeLib` pointer (a new reference).
 */
export function loadTypeLib(source: TypeLibSource): TypeLibPointer {
  if (source instanceof COMPointer) return source.QueryInterface(ITypeLib);
  if (typeof source === "object") return loadRegisteredTypeLib(source);
  if (GUID.isGUIDString(source)) return loadRegisteredTypeLib({ libid: source });
  if (!FILE_EXTENSION.test(source) && !/[\\/]/.test(source) && source.includes(".")) {
    const clsid = new GUID();
    if (ole32().CLSIDFromProgID(source, clsid._buffer) >= 0) return typeLibOfClass(clsid);
  }
  const out = new ctypes.c_void_p();
  checkHRESULT(oleaut32().LoadTypeLibEx(source, REGKIND_NONE, ctypes.byref(out)), `LoadTypeLibEx(${source})`);
  return COMPointer.fromAddress(out.value!, ITypeLib);
}

/** Read the `TLIBATTR` of a library. */
export function withLibAttr<T>(library: TypeLibPointer, use: (attr: TLIBATTR) => T): T {
  const out = new ctypes.c_void_p();
  library.GetLibAttr(ctypes.byref(out));
  try {
    return use(TLIBATTR.from_address(out.value!));
  } finally {
    library.ReleaseTLibAttr(out.value);
  }
}

/** `SYSKIND` of this process: `SYS_WIN64` on 64-bit Windows (x64 and ARM64), else `SYS_WIN32`. */
export const SYSKIND = POINTER_SIZE === 8 ? SYS_WIN64 : SYS_WIN32;

/**
 * Start a new type library file (`CreateTypeLib2`); `SaveAllChanges` writes it.
 * Python equivalent: `comtypes.typeinfo.CreateTypeLib`.
 */
export function createTypeLib(file: string, syskind: number = SYSKIND): InterfacePointer<typeof ICreateTypeLib2> {
  const out = new ctypes.c_void_p();
  checkHRESULT(oleaut32().CreateTypeLib2(syskind, file, ctypes.byref(out)), `CreateTypeLib2(${file})`);
  return COMPointer.fromAddress(out.value!, ICreateTypeLib2);
}

/** A registered type library: what `unregisterTypeLib` needs. */
export interface TypeLibIdentity {
  readonly libid: string;
  readonly major: number;
  readonly minor: number;
  readonly lcid: number;
}

export interface RegisterTypeLibOptions {
  /** Per-user (`HKEY_CURRENT_USER`, default) or machine-wide (administrators only). */
  readonly scope?: "user" | "machine";
  readonly helpDir?: string;
}

/**
 * Register a type library file: `TypeLib\{libid}` plus, for `[oleautomation]` and
 * dual interfaces, `Interface\{iid}` with the type library marshaler as proxy/stub,
 * so that they can be called across apartments and processes. Python
 * equivalent: `comtypes.typeinfo.RegisterTypeLib`.
 */
export function registerTypeLib(file: string, options: RegisterTypeLibOptions = {}): TypeLibIdentity {
  using library = loadTypeLib(file);
  const path = registeredPathOf(file);
  const register = options.scope === "machine" ? oleaut32().RegisterTypeLib : oleaut32().RegisterTypeLibForUser;
  checkHRESULT(register(library.address, path, options.helpDir ?? null), options.scope === "machine" ? "RegisterTypeLib" : "RegisterTypeLibForUser");
  return withLibAttr(library, (attr) => ({ libid: attr.guid.toString(), major: attr.wMajorVerNum, minor: attr.wMinorVerNum, lcid: attr.lcid }));
}

/** Remove a registration made by {@link registerTypeLib}. Python equivalent: `comtypes.typeinfo.UnRegisterTypeLib`. */
export function unregisterTypeLib(library: TypeLibIdentity, options: Pick<RegisterTypeLibOptions, "scope"> = {}): void {
  const unregister = options.scope === "machine" ? oleaut32().UnRegisterTypeLib : oleaut32().UnRegisterTypeLibForUser;
  checkHRESULT(unregister(toGUID(library.libid)._buffer, library.major, library.minor, library.lcid, SYSKIND), options.scope === "machine" ? "UnRegisterTypeLib" : "UnRegisterTypeLibForUser");
}

/** Absolute path of a type library file (`RegisterTypeLib` requires one). */
function registeredPathOf(file: string): string {
  return /^[a-z]:[\/]|^[\/]{2}/i.test(file) ? file : resolvePath(file);
}
