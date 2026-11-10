import * as ctypes from "node-ctypes";
import { APTTYPE_MAINSTA, APTTYPE_MTA, APTTYPE_NA, APTTYPE_STA, CLSCTX_INPROC_SERVER, CLSCTX_SERVER, COINIT_APARTMENTTHREADED } from "./constants.ts";
import { COMError, checkHRESULT } from "./errors.ts";
import { toGUID } from "./guid.ts";
import type { GUIDLike } from "./guid.ts";
import { CO_E_NOTINITIALIZED, E_OUTOFMEMORY, E_POINTER } from "./hresult.ts";
import type { InterfaceDef, InterfacePointer } from "./interface.ts";
import { IClassFactory } from "./interfaces/classfactory.ts";
import { IUnknown } from "./interfaces/unknown.ts";
import { ole32 } from "./low-level/native.ts";
import { COSERVERINFO, MULTI_QI } from "./low-level/types.ts";
import { COMPointer } from "./pointer.ts";
import { initCount, initialized, liveCount, uninitializing } from "./tracking.ts";

let shutdownRegistered = false;

/**
 * Initialize COM on the current thread. Default: single-threaded apartment,
 * as Python comtypes (required by Shell and most automation objects).
 * Returns `S_OK`, or `S_FALSE` when already initialized; every successful
 * call must be balanced by {@link CoUninitialize}.
 */
export function CoInitializeEx(flags: number = COINIT_APARTMENTTHREADED): number {
  const hr = checkHRESULT(ole32().CoInitializeEx(null, flags), "CoInitializeEx");
  initialized();
  if (!shutdownRegistered) {
    shutdownRegistered = true;
    // Python comtypes parity (atexit): release what is left and uninitialize.
    process.once("exit", () => {
      while (initCount() > 0) CoUninitialize();
    });
  }
  return hr;
}

/** `CoInitializeEx(COINIT_APARTMENTTHREADED)`. */
export function CoInitialize(): number {
  return CoInitializeEx(COINIT_APARTMENTTHREADED);
}

/**
 * Balance one {@link CoInitializeEx}. The last call releases every pointer
 * still alive on this thread before COM is torn down: they become unusable
 * and are never released again by the finalizer.
 */
export function CoUninitialize(): void {
  uninitializing();
  ole32().CoUninitialize();
}

export interface Apartment extends Disposable {
  /** `S_OK`, or `S_FALSE` when COM was already initialized on this thread. */
  readonly hresult: number;
}

/**
 * Scoped COM initialization:
 *
 * ```ts
 * using _ = apartment();
 * ```
 */
export function apartment(flags: number = COINIT_APARTMENTTHREADED): Apartment {
  const hresult = CoInitializeEx(flags);
  let disposed = false;
  return {
    hresult,
    [Symbol.dispose]() {
      if (disposed) return;
      disposed = true;
      CoUninitialize();
    },
  };
}

export interface ApartmentType {
  /** `APTTYPE_STA`, `APTTYPE_MTA`, `APTTYPE_NA` or `APTTYPE_MAINSTA`. */
  readonly type: number;
  readonly qualifier: number;
  readonly name: "STA" | "MTA" | "NA" | "MAINSTA";
}

/** Apartment of the current thread (`CoGetApartmentType`), or `null` when COM is not initialized. */
export function getApartmentType(): ApartmentType | null {
  const type = new ctypes.c_int32();
  const qualifier = new ctypes.c_int32();
  const hr = ole32().CoGetApartmentType(ctypes.byref(type), ctypes.byref(qualifier));
  if (hr === CO_E_NOTINITIALIZED) return null;
  checkHRESULT(hr, "CoGetApartmentType");
  const names = { [APTTYPE_STA]: "STA", [APTTYPE_MTA]: "MTA", [APTTYPE_NA]: "NA", [APTTYPE_MAINSTA]: "MAINSTA" } as const;
  const name = names[type.value as keyof typeof names];
  if (!name) throw new RangeError(`Unknown apartment type ${type.value}`);
  return { type: type.value, qualifier: qualifier.value, name };
}

/** Pointers created on this thread and not yet released (for leak checks). */
export function getLivePointerCount(): number {
  return liveCount();
}

export interface CreateInstanceOptions {
  /** Default `CLSCTX_SERVER` (in-process, local or remote), as Python comtypes. */
  readonly clsctx?: number;
  /** Controlling outer `IUnknown` for aggregation. */
  readonly outer?: COMPointer | null;
}

/**
 * Create a COM object and return the requested interface (default `IUnknown`).
 *
 * ```ts
 * using link = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
 * ```
 */
export function CoCreateInstance<T extends InterfaceDef = typeof IUnknown>(clsid: GUIDLike, iface?: T, options: CreateInstanceOptions = {}): InterfacePointer<T> {
  const target = (iface ?? IUnknown) as T;
  const out = new ctypes.c_void_p();
  checkHRESULT(ole32().CoCreateInstance(toGUID(clsid)._buffer, options.outer?.address ?? null, options.clsctx ?? CLSCTX_SERVER, target.iid._buffer, ctypes.byref(out)), "CoCreateInstance");
  return COMPointer.fromAddress(nonNull(out.value, "CoCreateInstance"), target);
}

export interface CreateInstanceExOptions extends CreateInstanceOptions {
  /** Remote machine name; local when omitted. */
  readonly machine?: string | null;
}

/** `CoCreateInstanceEx` for one interface, optionally on a remote machine. */
export function CoCreateInstanceEx<T extends InterfaceDef = typeof IUnknown>(clsid: GUIDLike, iface?: T, options: CreateInstanceExOptions = {}): InterfacePointer<T> {
  const target = (iface ?? IUnknown) as T;
  const machine = options.machine ? ctypes.create_unicode_buffer(options.machine) : null;
  const server = machine ? new COSERVERINFO() : null;
  if (server && machine) server.pwszName = ctypes.addressof(machine);
  const result = new MULTI_QI();
  result.pIID = ctypes.addressof(target.iid);
  const clsctx = options.clsctx ?? (machine ? CLSCTX_SERVER & ~CLSCTX_INPROC_SERVER : CLSCTX_SERVER);
  const hr = ole32().CoCreateInstanceEx(toGUID(clsid)._buffer, options.outer?.address ?? null, clsctx, server?._buffer ?? null, 1, result._buffer);
  // Keep the buffers referenced by address alive until the call returns.
  void [machine, target.iid];
  checkHRESULT(hr, "CoCreateInstanceEx");
  if (result.hr < 0) throw new COMError(result.hr, { method: "CoCreateInstanceEx" });
  return COMPointer.fromAddress(nonNull(result.pItf, "CoCreateInstanceEx"), target);
}

/** Class object for `clsid` (default `IClassFactory`). */
export function CoGetClassObject<T extends InterfaceDef = typeof IClassFactory>(clsid: GUIDLike, iface?: T, options: { readonly clsctx?: number } = {}): InterfacePointer<T> {
  const target = (iface ?? IClassFactory) as T;
  const out = new ctypes.c_void_p();
  checkHRESULT(ole32().CoGetClassObject(toGUID(clsid)._buffer, options.clsctx ?? CLSCTX_SERVER, null, target.iid._buffer, ctypes.byref(out)), "CoGetClassObject");
  return COMPointer.fromAddress(nonNull(out.value, "CoGetClassObject"), target);
}

/** Allocate COM task memory; throws {@link COMError} (`E_OUTOFMEMORY`) on failure. */
export function CoTaskMemAlloc(size: number | bigint): bigint {
  const pointer = ole32().CoTaskMemAlloc(size);
  if (!pointer) throw new COMError(E_OUTOFMEMORY, { method: "CoTaskMemAlloc" });
  return pointer;
}

export function CoTaskMemRealloc(pointer: bigint | null, size: number | bigint): bigint {
  const result = ole32().CoTaskMemRealloc(pointer, size);
  if (!result) throw new COMError(E_OUTOFMEMORY, { method: "CoTaskMemRealloc" });
  return result;
}

export function CoTaskMemFree(pointer: bigint | null): void {
  ole32().CoTaskMemFree(pointer);
}

function nonNull(address: bigint | null, method: string): bigint {
  if (!address) throw new COMError(E_POINTER, { method, text: `${method} returned a NULL interface pointer` });
  return address;
}
