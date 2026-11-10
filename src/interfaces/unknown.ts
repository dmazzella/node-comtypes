import * as ctypes from "node-ctypes";
import { GUID } from "../guid.ts";
import { COMInterface, HRESULT, STDMETHOD } from "../interface.ts";

const { POINTER, c_uint32, c_void_p } = ctypes;

/** The root COM interface. Python equivalent: `comtypes.IUnknown`. */
export const IUnknown = COMInterface({
  name: "IUnknown",
  iid: "{00000000-0000-0000-C000-000000000046}",
  methods: [STDMETHOD(HRESULT, "QueryInterface", [POINTER(GUID), POINTER(c_void_p)]), STDMETHOD(c_uint32, "AddRef"), STDMETHOD(c_uint32, "Release")],
});

export const IID_IUnknown: GUID = IUnknown.iid;
