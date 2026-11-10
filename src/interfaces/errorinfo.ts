import * as ctypes from "node-ctypes";
import { BSTR, COMMETHOD, REFIID } from "../commethod.ts";
import { GUID } from "../guid.ts";
import { COMInterface, HRESULT } from "../interface.ts";
import { IUnknown } from "./unknown.ts";

const { c_uint32, c_wchar_p } = ctypes;

export const IErrorInfo = COMInterface({
  name: "IErrorInfo",
  iid: "{1CF2B120-547D-101B-8E65-08002B2BD119}",
  extends: IUnknown,
  methods: [
    COMMETHOD([], HRESULT, "GetGUID", [["out"], GUID, "pGUID"]),
    COMMETHOD([], HRESULT, "GetSource", [["out"], BSTR, "pBstrSource"]),
    COMMETHOD([], HRESULT, "GetDescription", [["out"], BSTR, "pBstrDescription"]),
    COMMETHOD([], HRESULT, "GetHelpFile", [["out"], BSTR, "pBstrHelpFile"]),
    COMMETHOD([], HRESULT, "GetHelpContext", [["out"], c_uint32, "pdwHelpContext"]),
  ],
});

export const ICreateErrorInfo = COMInterface({
  name: "ICreateErrorInfo",
  iid: "{22F03340-547D-101B-8E65-08002B2BD119}",
  extends: IUnknown,
  methods: [
    COMMETHOD([], HRESULT, "SetGUID", [["in"], REFIID, "rguid"]),
    COMMETHOD([], HRESULT, "SetSource", [["in"], c_wchar_p, "szSource"]),
    COMMETHOD([], HRESULT, "SetDescription", [["in"], c_wchar_p, "szDescription"]),
    COMMETHOD([], HRESULT, "SetHelpFile", [["in"], c_wchar_p, "szHelpFile"]),
    COMMETHOD([], HRESULT, "SetHelpContext", [["in"], c_uint32, "dwHelpContext"]),
  ],
});

/** `InterfaceSupportsErrorInfo` returns `S_OK` when supported, `S_FALSE` otherwise. */
export const ISupportErrorInfo = COMInterface({
  name: "ISupportErrorInfo",
  iid: "{DF0B3D60-548F-101B-8E65-08002B2BD119}",
  extends: IUnknown,
  methods: [COMMETHOD([], HRESULT, "InterfaceSupportsErrorInfo", [["in"], REFIID, "riid"])],
});
