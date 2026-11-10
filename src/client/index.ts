import * as ctypes from "node-ctypes";
import { checkHRESULT } from "../errors.ts";
import { GUID } from "../guid.ts";
import type { InterfaceDef, InterfacePointer } from "../interface.ts";
import { IDispatch } from "../interfaces/dispatch.ts";
import { IUnknown } from "../interfaces/unknown.ts";
import { ole32, oleaut32 } from "../low-level/native.ts";
import { COMPointer } from "../pointer.ts";
import { CoCreateInstance, CoCreateInstanceEx } from "../runtime.ts";
import { GetBestInterface } from "./best.ts";
import { adoptDispatch } from "./dynamic.ts";
import { isCoClass } from "../typelib/coclass.ts";
import type { CoClass } from "../typelib/coclass.ts";
import type { AutomationObject, DispatchObject } from "./dynamic.ts";

export { Dispatch, adoptDispatch, dispatchPointer, enumerate, getProperty, invokeMethod, isDispatch, memberInfo, putProperty, putRefProperty, release, typeName } from "./dynamic.ts";
export type { AutomationObject, DispatchObject, DispatchTarget } from "./dynamic.ts";
export { GetBestInterface, wrap } from "./best.ts";
export { Constants } from "./constants.ts";
export type { ConstantsObject } from "./constants.ts";
export { GetEvents, ShowEvents } from "./events.ts";
export type { EventConnection, GetEventsOptions } from "./events.ts";

/** A ProgID (`"Scripting.Dictionary"`), a CLSID string or a GUID. */
export type ClassLike = GUID | string;

/** CLSID of a ProgID or CLSID string. */
export function classID(value: ClassLike): GUID {
  if (value instanceof GUID) return value;
  if (typeof value !== "string") throw new TypeError("Expected a ProgID, a CLSID string or a GUID");
  return GUID.isGUIDString(value) ? new GUID(value) : GUID.fromProgID(value);
}

export interface ObjectOptions<T extends InterfaceDef | undefined = undefined> {
  /** Return this interface. */
  readonly interface?: T;
  /** Return a dynamic dispatch object (`GetIDsOfNames` + `Invoke`, no type library). Python: `dynamic=True`. */
  readonly dynamic?: boolean;
}

export interface CreateObjectOptions<T extends InterfaceDef | undefined = undefined> extends ObjectOptions<T> {
  /** Default `CLSCTX_SERVER`. */
  readonly clsctx?: number;
  /** Remote machine (`CoCreateInstanceEx`). */
  readonly machine?: string;
}

/**
 * The object as requested: `interface`, a dynamic dispatch object, or its best
 * interface (dynamic when it has no type information). Takes over `pointer`.
 * Python equivalent: `comtypes.client._manage`.
 */
function manage(pointer: COMPointer, options: ObjectOptions<InterfaceDef | undefined>): AutomationObject | COMPointer {
  using owned = pointer;
  if (options.interface) return owned.QueryInterface(options.interface);
  if (options.dynamic) return adoptDispatch(owned.QueryInterface(IDispatch));
  const best = GetBestInterface(owned);
  if (best.interface !== IDispatch && best.interface !== IUnknown) return best as AutomationObject;
  using untyped = best;
  const dispatch = untyped.tryQueryInterface(IDispatch);
  return dispatch ? adoptDispatch(dispatch) : (untyped.QueryInterface(IUnknown) as AutomationObject);
}

/**
 * Create a COM object. Python equivalent: `comtypes.client.CreateObject`.
 *
 * ```ts
 * using dict = CreateObject("Scripting.Dictionary");   // IDictionary, typed from scrrun.dll
 * dict.Add("a", 1);
 * using dynamic = CreateObject("Scripting.Dictionary", { dynamic: true });  // IDispatch, late bound
 * using typed = CreateObject(Scripting.Dictionary);    // coclass: statically typed default interface
 * ```
 */
export function CreateObject<T extends InterfaceDef<any>>(coclass: CoClass<T>, options?: Omit<CreateObjectOptions, "interface" | "dynamic"> & { readonly dynamic?: false }): InterfacePointer<T>;
export function CreateObject(progid: ClassLike | CoClass, options: CreateObjectOptions & { readonly dynamic: true }): DispatchObject;
export function CreateObject<T extends InterfaceDef<any>>(progid: ClassLike | CoClass, options: CreateObjectOptions<T> & { readonly interface: T }): InterfacePointer<T>;
export function CreateObject(progid: ClassLike, options?: CreateObjectOptions): AutomationObject;
export function CreateObject(progid: ClassLike | CoClass, options: CreateObjectOptions<InterfaceDef | undefined> = {}): AutomationObject | COMPointer {
  // A coclass from a type library: its default interface, typed.
  const coclass = isCoClass(progid) ? progid : null;
  const clsid = coclass ? coclass.clsid : classID(progid as ClassLike);
  const iface = options.interface ?? (options.dynamic ? IDispatch : (coclass?.defaultInterface ?? IUnknown));
  const pointer = options.machine ? CoCreateInstanceEx(clsid, iface, { clsctx: options.clsctx, machine: options.machine }) : CoCreateInstance(clsid, iface, { clsctx: options.clsctx });
  return options.interface || (coclass && !options.dynamic) ? pointer : manage(pointer, options);
}

export type GetObjectOptions<T extends InterfaceDef | undefined = undefined> = ObjectOptions<T>;

/**
 * A running object registered in the Running Object Table. Python
 * equivalent: `comtypes.client.GetActiveObject`.
 */
export function GetActiveObject(progid: ClassLike, options: GetObjectOptions & { readonly dynamic: true }): DispatchObject;
export function GetActiveObject<T extends InterfaceDef>(progid: ClassLike, options: GetObjectOptions<T> & { readonly interface: T }): InterfacePointer<T>;
export function GetActiveObject(progid: ClassLike, options?: GetObjectOptions): AutomationObject;
export function GetActiveObject(progid: ClassLike, options: GetObjectOptions<InterfaceDef | undefined> = {}): AutomationObject | COMPointer {
  const out = new ctypes.c_void_p();
  checkHRESULT(oleaut32().GetActiveObject(classID(progid)._buffer, null, ctypes.byref(out)), "GetActiveObject");
  return manage(COMPointer.fromAddress(out.value!, IUnknown), options);
}

/**
 * Bind to an object by display name (`CoGetObject`), e.g. `"winmgmts:"` or a
 * file moniker. Python equivalent: `comtypes.client.GetObject`.
 */
export function GetObject(displayName: string, options: GetObjectOptions & { readonly dynamic: true }): DispatchObject;
export function GetObject<T extends InterfaceDef>(displayName: string, options: GetObjectOptions<T> & { readonly interface: T }): InterfacePointer<T>;
export function GetObject(displayName: string, options?: GetObjectOptions): AutomationObject;
export function GetObject(displayName: string, options: GetObjectOptions<InterfaceDef | undefined> = {}): AutomationObject | COMPointer {
  const iface = options.interface ?? (options.dynamic ? IDispatch : IUnknown);
  const out = new ctypes.c_void_p();
  checkHRESULT(ole32().CoGetObject(displayName, null, iface.iid._buffer, ctypes.byref(out)), "CoGetObject");
  return manage(COMPointer.fromAddress(out.value!, iface), options);
}
