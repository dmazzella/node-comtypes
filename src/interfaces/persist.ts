import * as ctypes from "node-ctypes";
import { EXCEPINFO } from "../automation/types.ts";
import { VARIANT } from "../automation/variant.ts";
import { COMMETHOD, LPOLESTR } from "../commethod.ts";
import { GUID } from "../guid.ts";
import { COMInterface, HRESULT, STDMETHOD } from "../interface.ts";
import { IUnknown } from "./unknown.ts";

const { POINTER, c_int32, c_uint16, c_uint32, c_void_p, c_wchar_p } = ctypes;

export const IPersist = COMInterface({
  name: "IPersist",
  iid: "{0000010C-0000-0000-C000-000000000046}",
  extends: IUnknown,
  methods: [COMMETHOD([], HRESULT, "GetClassID", [["out"], GUID, "pClassID"])],
});

export const IPersistFile = COMInterface({
  name: "IPersistFile",
  iid: "{0000010B-0000-0000-C000-000000000046}",
  extends: IPersist,
  methods: [
    // S_OK when dirty, S_FALSE otherwise.
    COMMETHOD([], HRESULT, "IsDirty"),
    COMMETHOD([], HRESULT, "Load", [["in"], c_wchar_p, "pszFileName"], [["in", "optional"], c_uint32, "dwMode", 0]),
    COMMETHOD([], HRESULT, "Save", [["in"], c_wchar_p, "pszFileName"], [["in", "optional"], c_int32, "fRemember", 1]),
    COMMETHOD([], HRESULT, "SaveCompleted", [["in"], c_wchar_p, "pszFileName"]),
    COMMETHOD([], HRESULT, "GetCurFile", [["out"], LPOLESTR, "ppszFileName"]),
  ],
});

/** `STGM_*` access modes for `IPersistFile::Load`. */
export const STGM_READ = 0x0;
export const STGM_WRITE = 0x1;
export const STGM_READWRITE = 0x2;

/** Errors reported while loading properties (`IPropertyBag::Read`, `IPersistPropertyBag::Load`). */
export const IErrorLog = COMInterface({
  name: "IErrorLog",
  iid: "{3127CA40-446E-11CE-8135-00AA004BB851}",
  extends: IUnknown,
  methods: [COMMETHOD([], HRESULT, "AddError", [["in"], c_wchar_p, "pszPropName"], [["in"], POINTER(EXCEPINFO), "pExcepInfo"])],
});

/**
 * Named properties as VARIANTs. `Read(name, initial, errorLog)` returns the
 * value (`[in, out] VARIANT*`); `Write(name, value)`.
 */
export const IPropertyBag = COMInterface({
  name: "IPropertyBag",
  iid: "{55272A00-42CB-11CE-8135-00AA004BB851}",
  extends: IUnknown,
  methods: [
    COMMETHOD([], HRESULT, "Read", [["in"], c_wchar_p, "pszPropName"], [["in", "out", "optional"], VARIANT, "pVar"], [["in", "optional"], IErrorLog, "pErrorLog"]),
    COMMETHOD([], HRESULT, "Write", [["in"], c_wchar_p, "pszPropName"], [["in"], POINTER(VARIANT), "pVar"]),
  ],
});

export const IPersistPropertyBag = COMInterface({
  name: "IPersistPropertyBag",
  iid: "{37D84F60-42CB-11CE-8135-00AA004BB851}",
  extends: IPersist,
  methods: [
    COMMETHOD([], HRESULT, "InitNew"),
    COMMETHOD([], HRESULT, "Load", [["in"], IPropertyBag, "pPropBag"], [["in"], IErrorLog, "pErrorLog"]),
    COMMETHOD([], HRESULT, "Save", [["in"], IPropertyBag, "pPropBag"], [["in"], c_int32, "fClearDirty"], [["in"], c_int32, "fSaveAllProperties"]),
  ],
});

/** Property description of `IPropertyBag2`. */
export class PROPBAG2 extends ctypes.Structure {
  static _fields_ = [
    ["dwType", c_uint32],
    ["vt", c_uint16],
    ["cfType", c_uint16],
    ["dwHint", c_uint32],
    ["pstrName", c_void_p],
    ["clsid", GUID],
  ] as const;
  declare dwType: number;
  declare vt: number;
  declare cfType: number;
  declare dwHint: number;
  declare pstrName: bigint | null;
  declare clsid: GUID;
}

/** `PROPBAG2.dwType`. */
export const PROPBAG2_TYPE_UNDEFINED = 0;
export const PROPBAG2_TYPE_DATA = 1;
export const PROPBAG2_TYPE_URL = 2;
export const PROPBAG2_TYPE_OBJECT = 3;
export const PROPBAG2_TYPE_STREAM = 4;
export const PROPBAG2_TYPE_STORAGE = 5;
export const PROPBAG2_TYPE_MONIKER = 6;

/** Arrays of properties: `Read`, `Write` and `GetPropertyInfo` take raw `PROPBAG2` / `VARIANT` arrays. */
export const IPropertyBag2 = COMInterface({
  name: "IPropertyBag2",
  iid: "{22F55882-280B-11D0-A8A9-00A0C90C2004}",
  extends: IUnknown,
  methods: [
    STDMETHOD(HRESULT, "Read", [c_uint32, c_void_p, c_void_p, c_void_p, c_void_p]),
    STDMETHOD(HRESULT, "Write", [c_uint32, c_void_p, c_void_p]),
    COMMETHOD([], HRESULT, "CountProperties", [["out"], c_uint32, "pcProperties"]),
    STDMETHOD(HRESULT, "GetPropertyInfo", [c_uint32, c_uint32, c_void_p, POINTER(c_uint32)]),
    COMMETHOD([], HRESULT, "LoadObject", [["in"], c_wchar_p, "pstrName"], [["in"], c_uint32, "dwHint"], [["in"], IUnknown, "pUnkObject"], [["in"], IErrorLog, "pErrLog"]),
  ],
});

export const IPersistPropertyBag2 = COMInterface({
  name: "IPersistPropertyBag2",
  iid: "{22F55881-280B-11D0-A8A9-00A0C90C2004}",
  extends: IPersist,
  methods: [
    COMMETHOD([], HRESULT, "InitNew"),
    COMMETHOD([], HRESULT, "Load", [["in"], IPropertyBag2, "pPropBag"], [["in"], IErrorLog, "pErrLog"]),
    COMMETHOD([], HRESULT, "Save", [["in"], IPropertyBag2, "pPropBag"], [["in"], c_int32, "fClearDirty"], [["in"], c_int32, "fSaveAllProperties"]),
    // S_OK when dirty, S_FALSE otherwise.
    COMMETHOD([], HRESULT, "IsDirty"),
  ],
});
