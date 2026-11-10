import * as ctypes from "node-ctypes";
import { borrowReference } from "./automation/variant.ts";
import { adoptDispatch } from "./client/dynamic.ts";
import type { DispatchObject } from "./client/dynamic.ts";
import { COMMETHOD, REFIID } from "./commethod.ts";
import type { ComObject } from "./commethod.ts";
import { checkHRESULT } from "./errors.ts";
import { GUID } from "./guid.ts";
import { COMInterface, HRESULT, STDMETHOD } from "./interface.ts";
import type { InterfaceDef, InterfacePointer } from "./interface.ts";
import { IDispatch } from "./interfaces/dispatch.ts";
import { IUnknown } from "./interfaces/unknown.ts";
import { ole32 } from "./low-level/native.ts";
import { rawRelease } from "./low-level/vtable.ts";
import { COMPointer } from "./pointer.ts";
import { CoCreateInstance } from "./runtime.ts";

// COM pointers belong to their apartment: other threads (worker_threads, each
// with its own apartment) must receive a marshaled reference, which COM turns
// into a proxy. Calls into a single-threaded apartment are delivered through
// its message queue: the owning thread must pump (startMessagePump).

/**
 * A reference marshaled for another thread, cloneable with `postMessage`.
 * Single use: unmarshal it exactly once (or release it with `releaseMarshaled`).
 */
export interface MarshaledInterface {
  readonly kind: "node-comtypes.marshaled";
  /** `IStream` holding the marshaled data (a process-wide address). */
  readonly stream: bigint;
  readonly iid: string;
}

function source(object: COMPointer | ComObject): { address: bigint; iid: GUID } {
  const borrowed = borrowReference(object);
  if (!borrowed) throw new TypeError("Expected a COM object");
  return { address: borrowed.address, iid: borrowed.interface.iid };
}

/**
 * Marshal an interface for another thread (`CoMarshalInterThreadInterfaceInStream`).
 * Default interface: the pointer's own. Custom vtable interfaces need a
 * registered proxy/stub or type library; `IDispatch` always works.
 */
export function marshalInterface(object: COMPointer | ComObject, iface?: InterfaceDef<any>): MarshaledInterface {
  const { address, iid: own } = source(object);
  const iid = iface?.iid ?? own;
  const stream = new ctypes.c_void_p();
  checkHRESULT(ole32().CoMarshalInterThreadInterfaceInStream(iid._buffer, address, ctypes.byref(stream)), "CoMarshalInterThreadInterfaceInStream");
  return Object.freeze({ kind: "node-comtypes.marshaled", stream: stream.value!, iid: iid.toString() });
}

function checkToken(token: MarshaledInterface): void {
  if (token?.kind !== "node-comtypes.marshaled" || typeof token.stream !== "bigint") throw new TypeError("Expected a token from marshalInterface");
}

/** Unmarshal in the receiving thread (`CoGetInterfaceAndReleaseStream`); the token is consumed. */
export function unmarshalInterface<T extends InterfaceDef<any>>(token: MarshaledInterface, iface: T): InterfacePointer<T> {
  checkToken(token);
  const out = new ctypes.c_void_p();
  checkHRESULT(ole32().CoGetInterfaceAndReleaseStream(token.stream, iface.iid._buffer, ctypes.byref(out)), "CoGetInterfaceAndReleaseStream");
  return COMPointer.fromAddress(out.value!, iface) as InterfacePointer<T>;
}

/** Unmarshal as a dynamic dispatch object. */
export function unmarshalDispatch(token: MarshaledInterface): DispatchObject {
  return adoptDispatch(unmarshalInterface(token, IDispatch));
}

/** Discard a token that will not be unmarshaled. */
export function releaseMarshaled(token: MarshaledInterface): void {
  checkToken(token);
  ole32().CoReleaseMarshalData(token.stream);
  rawRelease(token.stream);
}

/** Python/C equivalent: `IGlobalInterfaceTable`. */
export const IGlobalInterfaceTable = COMInterface({
  name: "IGlobalInterfaceTable",
  iid: "{00000146-0000-0000-C000-000000000046}",
  extends: IUnknown,
  methods: [
    COMMETHOD([], HRESULT, "RegisterInterfaceInGlobal", [["in"], IUnknown, "pUnk"], [["in"], REFIID, "riid"], [["out"], ctypes.c_uint32, "pdwCookie"]),
    COMMETHOD([], HRESULT, "RevokeInterfaceFromGlobal", [["in"], ctypes.c_uint32, "dwCookie"]),
    STDMETHOD(HRESULT, "GetInterfaceFromGlobal", [ctypes.c_uint32, ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p)]),
  ],
});

export const CLSID_StdGlobalInterfaceTable: GUID = new GUID("{00000323-0000-0000-C000-000000000046}");

/**
 * The process-wide Global Interface Table: register a reference once, then
 * get a proxy in any thread (as many times as needed) with the cookie.
 */
export class GlobalInterfaceTable implements Disposable {
  readonly #table: InterfacePointer<typeof IGlobalInterfaceTable>;

  constructor() {
    this.#table = CoCreateInstance(CLSID_StdGlobalInterfaceTable, IGlobalInterfaceTable);
  }

  /** Register `object` (as `iface`, default its own interface); returns the cookie, valid in every thread. */
  register(object: COMPointer | ComObject, iface?: InterfaceDef<any>): number {
    const { iid } = source(object);
    return this.#table.RegisterInterfaceInGlobal(object, iface?.iid ?? iid);
  }

  /** A pointer usable in the calling thread. */
  get<T extends InterfaceDef<any>>(cookie: number, iface: T): InterfacePointer<T> {
    const out = new ctypes.c_void_p();
    this.#table.GetInterfaceFromGlobal(cookie, iface.iid._buffer, ctypes.byref(out));
    return COMPointer.fromAddress(out.value!, iface) as InterfacePointer<T>;
  }

  getDispatch(cookie: number): DispatchObject {
    return adoptDispatch(this.get(cookie, IDispatch));
  }

  revoke(cookie: number): void {
    this.#table.RevokeInterfaceFromGlobal(cookie);
  }

  [Symbol.dispose](): void {
    if (!this.#table.released) this.#table.Release();
  }
}
