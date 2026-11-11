import * as ctypes from "node-ctypes";
import { CLSCTX_INPROC_SERVER, CLSCTX_SERVER } from "../constants.ts";
import { COMError, checkHRESULT } from "../errors.ts";
import { handOver, homeApartment } from "../executor/apartment.ts";
import { isLocalClass } from "../executor/classes.ts";
import { executorTarget } from "../executor/executor.ts";
import type { CallTarget } from "../executor/executor.ts";
import { operation } from "../executor/operations.ts";
import type { Operation } from "../executor/operations.ts";
import { GUID } from "../guid.ts";
import { E_POINTER } from "../hresult.ts";
import type { InterfaceDef, InterfacePointer } from "../interface.ts";
import { IDispatch } from "../interfaces/dispatch.ts";
import { IUnknown } from "../interfaces/unknown.ts";
import { ole32, oleaut32 } from "../low-level/native.ts";
import { COSERVERINFO, MULTI_QI } from "../low-level/types.ts";
import { COMPointer, sync as syncPointer } from "../pointer.ts";
import { CoCreateInstance, CoCreateInstanceEx } from "../runtime.ts";
import { queryOrDispatch } from "./dual.ts";
import type { SyncPointer } from "../interface.ts";
import { GetBestInterface, loadBestModule } from "./best.ts";
import { adoptDispatch, isDispatch, syncDispatch } from "./dynamic.ts";
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

/**
 * A view whose methods run synchronously on the home thread: of a typed
 * pointer (sharing its reference) or of a dynamic dispatch object. For event
 * handlers, COM servers and code that needs the lowest latency.
 */
export function sync<I extends InterfaceDef<any>>(pointer: COMPointer<I> | InterfacePointer<I>): SyncPointer<I>;
export function sync(object: DispatchObject): DispatchObject;
export function sync(object: COMPointer | DispatchObject): SyncPointer<any> | DispatchObject {
  return isDispatch(object) ? syncDispatch(object) : syncPointer(object as COMPointer);
}

export interface ObjectOptions<T extends InterfaceDef | undefined = undefined> {
  /** Return this interface. */
  readonly interface?: T;
  /** Return a dynamic dispatch object (`GetIDsOfNames` + `Invoke`, no type library). Python: `dynamic=True`. */
  readonly dynamic?: boolean;
  /**
   * Obtain the object on the home thread instead of the COM executor: an
   * in-process object then lives in the home apartment, called directly
   * (no marshaling, no executor) but always on the home thread. Needed in a
   * single-threaded home for objects whose interfaces cannot be marshaled
   * (e.g. ADSI `IDirectorySearch`); a multithreaded home shares objects with
   * the executor anyway.
   */
  readonly home?: boolean;
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
 * {@link manage}, after loading the type library of the object's best interface
 * asynchronously: the first object of a large library does not block the
 * event loop while its definitions are built.
 */
async function manageLoaded(pointer: COMPointer, options: ObjectOptions<InterfaceDef | undefined>): Promise<AutomationObject | COMPointer> {
  if (!options.interface && !options.dynamic) {
    try {
      await loadBestModule(pointer);
    } catch (error) {
      pointer.Release();
      throw error;
    }
  }
  return manage(pointer, options);
}

/**
 * `iface` of `unknown`; a dispinterface (or a dual interface whose vtable cannot
 * be marshaled across apartments) is reached through `IDispatch`.
 */
function queryInterface(unknown: COMPointer, iface: InterfaceDef<any>): COMPointer {
  return queryOrDispatch(unknown, iface) ?? unknown.QueryInterface(iface);
}

/**
 * Obtain an object on the COM executor: `start` queues the call whose out
 * parameter `out` receives an `IUnknown`. The object lives in the
 * multithreaded apartment (or in COM's host apartment for single-threaded
 * classes); a single-threaded home receives its own reference.
 */
async function obtain(method: string, current: Operation, start: (target: ctypes.ExecutorTarget, out: InstanceType<typeof ctypes.c_void_p>) => Promise<unknown>): Promise<COMPointer<typeof IUnknown>> {
  const out = new ctypes.c_void_p();
  if (current.aborted) throw current.reason;
  const target: CallTarget = { executor: executorTarget(), _job: 0 };
  const pending = start(target, out);
  if (target._job !== 0) current.queued(target.executor, target._job);
  checkHRESULT((await pending) as number, method);
  if (!out.value) throw new COMError(E_POINTER, { method });
  if (homeApartment() === "sta") await handOver([{ location: ctypes.addressof(out), kind: "pointer", iid: IUnknown.iid }]);
  return COMPointer.fromAddress(out.value!, IUnknown);
}

/** `IUnknown` from a call made on the home thread (`home: true`). */
function obtainHome(method: string, call: (out: InstanceType<typeof ctypes.c_void_p>) => number): COMPointer<typeof IUnknown> {
  const out = new ctypes.c_void_p();
  checkHRESULT(call(out), method);
  if (!out.value) throw new COMError(E_POINTER, { method });
  return COMPointer.fromAddress(out.value, IUnknown);
}

/** `CoCreateInstanceEx` for `IUnknown` on the executor (remote machines). */
function createInstanceEx(target: ctypes.ExecutorTarget, clsid: GUID, clsctx: number, machine: string, out: InstanceType<typeof ctypes.c_void_p>): Promise<number> {
  const name = ctypes.create_unicode_buffer(machine);
  const server = new COSERVERINFO();
  server.pwszName = ctypes.addressof(name);
  const result = new MULTI_QI();
  result.pIID = ctypes.addressof(IUnknown.iid);
  return ole32()
    .CoCreateInstanceEx.callOn(target, clsid._buffer, null, clsctx, server._buffer, 1, result._buffer)
    .then((hr) => {
      void name;
      if (hr < 0) return hr;
      out.value = result.pItf;
      return result.hr;
    });
}

/**
 * Create a COM object on the COM executor. Python equivalent:
 * `comtypes.client.CreateObject`.
 *
 * ```ts
 * using dict = await CreateObject("Scripting.Dictionary");   // IDictionary, typed from scrrun.dll
 * await dict.Add("a", 1);
 * using dynamic = await CreateObject("Scripting.Dictionary", { dynamic: true });  // IDispatch, late bound
 * using typed = await CreateObject(Scripting.Dictionary);    // coclass: statically typed default interface
 * ```
 */
export function CreateObject<T extends InterfaceDef<any>>(coclass: CoClass<T>, options?: Omit<CreateObjectOptions, "interface" | "dynamic"> & { readonly dynamic?: false }): Promise<InterfacePointer<T>>;
export function CreateObject(progid: ClassLike | CoClass, options: CreateObjectOptions & { readonly dynamic: true }): Promise<DispatchObject>;
export function CreateObject<T extends InterfaceDef<any>>(progid: ClassLike | CoClass, options: CreateObjectOptions<T> & { readonly interface: T }): Promise<InterfacePointer<T>>;
export function CreateObject(progid: ClassLike, options?: CreateObjectOptions): Promise<AutomationObject>;
export function CreateObject(progid: ClassLike | CoClass, options: CreateObjectOptions<InterfaceDef | undefined> = {}): Promise<AutomationObject | COMPointer> {
  return operation(async (current) => {
    // A coclass from a type library: its default interface, typed.
    const coclass = isCoClass(progid) ? progid : null;
    const clsid = coclass ? coclass.clsid : classID(progid as ClassLike);
    const clsctx = options.clsctx ?? (options.machine ? CLSCTX_SERVER & ~CLSCTX_INPROC_SERVER : CLSCTX_SERVER);
    const unknown: COMPointer =
      options.home || isLocalClass(clsid)
        ? options.machine
          ? CoCreateInstanceEx(clsid, IUnknown, { clsctx, machine: options.machine })
          : CoCreateInstance(clsid, IUnknown, { clsctx })
        : await obtain(options.machine ? "CoCreateInstanceEx" : "CoCreateInstance", current, (target, out) =>
            options.machine ? createInstanceEx(target, clsid, clsctx, options.machine, out) : ole32().CoCreateInstance.callOn(target, clsid._buffer, null, clsctx, IUnknown.iid._buffer, ctypes.byref(out)),
          );
    if (options.interface || (coclass && !options.dynamic)) {
      using owned = unknown;
      return queryInterface(owned, options.interface ?? coclass!.defaultInterface);
    }
    return manageLoaded(unknown, options);
  });
}

export type GetObjectOptions<T extends InterfaceDef | undefined = undefined> = ObjectOptions<T>;

/**
 * A running object registered in the Running Object Table, obtained on the
 * COM executor. Python equivalent: `comtypes.client.GetActiveObject`.
 */
export function GetActiveObject(progid: ClassLike, options: GetObjectOptions & { readonly dynamic: true }): Promise<DispatchObject>;
export function GetActiveObject<T extends InterfaceDef>(progid: ClassLike, options: GetObjectOptions<T> & { readonly interface: T }): Promise<InterfacePointer<T>>;
export function GetActiveObject(progid: ClassLike, options?: GetObjectOptions): Promise<AutomationObject>;
export function GetActiveObject(progid: ClassLike, options: GetObjectOptions<InterfaceDef | undefined> = {}): Promise<AutomationObject | COMPointer> {
  return operation(async (current) => {
    const clsid = classID(progid);
    const unknown = options.home
      ? obtainHome("GetActiveObject", (out) => oleaut32().GetActiveObject(clsid._buffer, null, ctypes.byref(out)))
      : await obtain("GetActiveObject", current, (target, out) => oleaut32().GetActiveObject.callOn(target, clsid._buffer, null, ctypes.byref(out)));
    return manageLoaded(unknown, options);
  });
}

/**
 * Bind to an object by display name (`CoGetObject`) on the COM executor,
 * e.g. `"winmgmts:"` or a file moniker. Python equivalent:
 * `comtypes.client.GetObject`.
 */
export function GetObject(displayName: string, options: GetObjectOptions & { readonly dynamic: true }): Promise<DispatchObject>;
export function GetObject<T extends InterfaceDef>(displayName: string, options: GetObjectOptions<T> & { readonly interface: T }): Promise<InterfacePointer<T>>;
export function GetObject(displayName: string, options?: GetObjectOptions): Promise<AutomationObject>;
export function GetObject(displayName: string, options: GetObjectOptions<InterfaceDef | undefined> = {}): Promise<AutomationObject | COMPointer> {
  return operation(async (current) => {
    const unknown = options.home
      ? obtainHome("CoGetObject", (out) => ole32().CoGetObject(displayName, null, IUnknown.iid._buffer, ctypes.byref(out)))
      : await obtain("CoGetObject", current, (target, out) => ole32().CoGetObject.callOn(target, displayName, null, IUnknown.iid._buffer, ctypes.byref(out)));
    return manageLoaded(unknown, options);
  });
}
