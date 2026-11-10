import * as ctypes from "node-ctypes";
import { borrowReference } from "../automation/variant.ts";
import type { ComObject } from "../commethod.ts";
import { GUID, toGUID } from "../guid.ts";
import type { GUIDLike } from "../guid.ts";
import { COMInterface, HRESULT, STDMETHOD } from "../interface.ts";
import type { InterfaceDef, InterfacePointer } from "../interface.ts";
import { COMPointer } from "../pointer.ts";
import { IUnknown } from "./unknown.ts";

const { POINTER, c_void_p } = ctypes;

/** Services of an object by service ID. Python equivalent: `comtypes.IServiceProvider`. */
export const IServiceProvider = COMInterface({
  name: "IServiceProvider",
  iid: "{6D5140C1-7436-11CE-8034-00AA006009FA}",
  extends: IUnknown,
  methods: [STDMETHOD(HRESULT, "QueryService", [POINTER(GUID), POINTER(GUID), POINTER(c_void_p)])],
});

/**
 * `IServiceProvider::QueryService`: the service `service` of `object`, as
 * `iface`. Python equivalent: `IServiceProvider.QueryService(serviceIID, interface)`.
 */
export function queryService<T extends InterfaceDef<any>>(object: COMPointer | ComObject, service: GUIDLike, iface: T): InterfacePointer<T> {
  const borrowed = borrowReference(object);
  if (!borrowed) throw new TypeError("queryService expects a COM object");
  using pointer = COMPointer.fromAddress(borrowed.address, borrowed.interface, { addRef: true });
  using provider = pointer.QueryInterface(IServiceProvider);
  const out = new ctypes.c_void_p();
  provider.QueryService(toGUID(service), iface.iid, ctypes.byref(out));
  return COMPointer.fromAddress(out.value!, iface) as InterfacePointer<T>;
}
