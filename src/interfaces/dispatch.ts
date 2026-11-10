import * as ctypes from "node-ctypes";
import { DISPPARAMS, EXCEPINFO } from "../automation/types.ts";
import { VARIANT } from "../automation/variant.ts";
import { GUID } from "../guid.ts";
import { COMMETHOD, SELF } from "../commethod.ts";
import { COMInterface, HRESULT, STDMETHOD } from "../interface.ts";
import { IUnknown } from "./unknown.ts";

const { POINTER, c_int32, c_uint16, c_uint32, c_void_p } = ctypes;

/** Late-bound automation. Python equivalent: `comtypes.automation.IDispatch`. */
export const IDispatch = COMInterface({
  name: "IDispatch",
  iid: "{00020400-0000-0000-C000-000000000046}",
  extends: IUnknown,
  methods: [
    STDMETHOD(HRESULT, "GetTypeInfoCount", [POINTER(c_uint32)]),
    STDMETHOD(HRESULT, "GetTypeInfo", [c_uint32, c_uint32, POINTER(c_void_p)]),
    // rgszNames is LPOLESTR*: an array of wide string pointers.
    STDMETHOD(HRESULT, "GetIDsOfNames", [POINTER(GUID), c_void_p, c_uint32, c_uint32, POINTER(c_int32)]),
    STDMETHOD(HRESULT, "Invoke", [c_int32, POINTER(GUID), c_uint32, c_uint16, POINTER(DISPPARAMS), POINTER(VARIANT), POINTER(EXCEPINFO), POINTER(c_uint32)]),
  ],
});

export const IID_IDispatch: GUID = IDispatch.iid;

/** Enumerator returned by `_NewEnum` (`DISPID_NEWENUM`). `Next` takes raw arguments (a VARIANT array). */
export const IEnumVARIANT = COMInterface({
  name: "IEnumVARIANT",
  iid: "{00020404-0000-0000-C000-000000000046}",
  extends: IUnknown,
  methods: [STDMETHOD(HRESULT, "Next", [c_uint32, POINTER(VARIANT), POINTER(c_uint32)]), COMMETHOD([], HRESULT, "Skip", [["in"], c_uint32, "celt"]), COMMETHOD([], HRESULT, "Reset"), COMMETHOD([], HRESULT, "Clone", [["out"], SELF, "ppEnum"])],
});
