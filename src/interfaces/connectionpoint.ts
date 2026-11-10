import * as ctypes from "node-ctypes";
import { COMMETHOD, REFIID, SELF } from "../commethod.ts";
import { GUID } from "../guid.ts";
import { COMInterface, HRESULT, STDMETHOD } from "../interface.ts";
import { IProvideClassInfo } from "./typeinfo.ts";
import { IUnknown } from "./unknown.ts";

const { POINTER, c_uint32, c_void_p } = ctypes;

/** One advise connection, as returned by `IEnumConnections::Next`. */
export class CONNECTDATA extends ctypes.Structure {
  static _fields_ = [
    ["pUnk", c_void_p],
    ["dwCookie", c_uint32],
  ] as const;
  declare pUnk: bigint | null;
  declare dwCookie: number;
}

export const IEnumConnections = COMInterface({
  name: "IEnumConnections",
  iid: "{B196B287-BAB4-101A-B69C-00AA00341D07}",
  extends: IUnknown,
  methods: [
    STDMETHOD(HRESULT, "Next", [c_uint32, POINTER(CONNECTDATA), POINTER(c_uint32)]),
    COMMETHOD([], HRESULT, "Skip", [["in"], c_uint32, "cConnections"]),
    COMMETHOD([], HRESULT, "Reset"),
    COMMETHOD([], HRESULT, "Clone", [["out"], SELF, "ppEnum"]),
  ],
});

export const IConnectionPoint = COMInterface({
  name: "IConnectionPoint",
  iid: "{B196B286-BAB4-101A-B69C-00AA00341D07}",
  extends: IUnknown,
  methods: [
    COMMETHOD([], HRESULT, "GetConnectionInterface", [["out"], GUID, "pIID"]),
    // IConnectionPointContainer is declared below: the raw pointer is returned here.
    COMMETHOD([], HRESULT, "GetConnectionPointContainer", [["out"], IUnknown, "ppCPC"]),
    COMMETHOD([], HRESULT, "Advise", [["in"], IUnknown, "pUnkSink"], [["out"], c_uint32, "pdwCookie"]),
    COMMETHOD([], HRESULT, "Unadvise", [["in"], c_uint32, "dwCookie"]),
    COMMETHOD([], HRESULT, "EnumConnections", [["out"], IEnumConnections, "ppEnum"]),
  ],
});

export const IEnumConnectionPoints = COMInterface({
  name: "IEnumConnectionPoints",
  iid: "{B196B285-BAB4-101A-B69C-00AA00341D07}",
  extends: IUnknown,
  methods: [
    STDMETHOD(HRESULT, "Next", [c_uint32, POINTER(c_void_p), POINTER(c_uint32)]),
    COMMETHOD([], HRESULT, "Skip", [["in"], c_uint32, "cConnections"]),
    COMMETHOD([], HRESULT, "Reset"),
    COMMETHOD([], HRESULT, "Clone", [["out"], SELF, "ppEnum"]),
  ],
});

export const IConnectionPointContainer = COMInterface({
  name: "IConnectionPointContainer",
  iid: "{B196B284-BAB4-101A-B69C-00AA00341D07}",
  extends: IUnknown,
  methods: [COMMETHOD([], HRESULT, "EnumConnectionPoints", [["out"], IEnumConnectionPoints, "ppEnum"]), COMMETHOD([], HRESULT, "FindConnectionPoint", [["in"], REFIID, "riid"], [["out"], IConnectionPoint, "ppCP"])],
});

/** `GetGUID(GUIDKIND_DEFAULT_SOURCE_DISP_IID)`: the default event interface. */
export const IProvideClassInfo2 = COMInterface({
  name: "IProvideClassInfo2",
  iid: "{A6BC3AC0-DBAA-11CE-9DE3-00AA004BB851}",
  extends: IProvideClassInfo,
  methods: [COMMETHOD([], HRESULT, "GetGUID", [["in"], c_uint32, "dwGuidKind"], [["out"], GUID, "pGUID"])],
});

export const GUIDKIND_DEFAULT_SOURCE_DISP_IID = 1;
