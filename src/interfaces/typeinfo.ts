import * as ctypes from "node-ctypes";
import { DISPPARAMS, EXCEPINFO } from "../automation/types.ts";
import { VARIANT } from "../automation/variant.ts";
import { GUID } from "../guid.ts";
import { BSTR, COMMETHOD, LPOLESTR, REFGUID } from "../commethod.ts";
import { COMInterface, HRESULT, STDMETHOD } from "../interface.ts";
import { IUnknown } from "./unknown.ts";

const { POINTER, c_int16, c_int32, c_uint16, c_uint32, c_void_p, c_wchar_p } = ctypes;

/** Type of a parameter, return value or variable. `lptdesc` overlays `lpadesc` and `hreftype`. */
export class TYPEDESC extends ctypes.Structure {
  static _fields_ = [
    ["lptdesc", c_void_p],
    ["vt", c_uint16],
  ] as const;
  declare lptdesc: bigint | null;
  declare vt: number;
  /** `HREFTYPE` for `VT_USERDEFINED`. */
  get hreftype(): number {
    return this._buffer.readUInt32LE(0);
  }
}

/** Also used as `IDLDESC` (same layout: `dwReserved`, `wIDLFlags`). */
export class PARAMDESC extends ctypes.Structure {
  static _fields_ = [
    ["pparamdescex", c_void_p],
    ["wParamFlags", c_uint16],
  ] as const;
  declare pparamdescex: bigint | null;
  declare wParamFlags: number;
}

export class ELEMDESC extends ctypes.Structure {
  static _fields_ = [
    ["tdesc", TYPEDESC],
    ["paramdesc", PARAMDESC],
  ] as const;
  declare tdesc: TYPEDESC;
  declare paramdesc: PARAMDESC;
}

/** Default value of an optional parameter (`PARAMFLAG_FHASDEFAULT`). */
export class PARAMDESCEX extends ctypes.Structure {
  static _fields_ = [
    ["cBytes", c_uint32],
    ["varDefaultValue", VARIANT],
  ] as const;
  declare cBytes: number;
  declare varDefaultValue: VARIANT;
}

export class FUNCDESC extends ctypes.Structure {
  static _fields_ = [
    ["memid", c_int32],
    ["lprgscode", c_void_p],
    ["lprgelemdescParam", c_void_p],
    ["funckind", c_int32],
    ["invkind", c_int32],
    ["callconv", c_int32],
    ["cParams", c_int16],
    ["cParamsOpt", c_int16],
    ["oVft", c_int16],
    ["cScodes", c_int16],
    ["elemdescFunc", ELEMDESC],
    ["wFuncFlags", c_uint16],
  ] as const;
  declare memid: number;
  declare lprgscode: bigint | null;
  declare lprgelemdescParam: bigint | null;
  declare funckind: number;
  declare invkind: number;
  declare callconv: number;
  declare cParams: number;
  declare cParamsOpt: number;
  declare oVft: number;
  declare cScodes: number;
  declare elemdescFunc: ELEMDESC;
  declare wFuncFlags: number;
}

/** `lpvarValue` overlays `oInst`. */
export class VARDESC extends ctypes.Structure {
  static _fields_ = [
    ["memid", c_int32],
    ["lpstrSchema", c_void_p],
    ["lpvarValue", c_void_p],
    ["elemdescVar", ELEMDESC],
    ["wVarFlags", c_uint16],
    ["varkind", c_int32],
  ] as const;
  declare memid: number;
  declare lpstrSchema: bigint | null;
  declare lpvarValue: bigint | null;
  declare elemdescVar: ELEMDESC;
  declare wVarFlags: number;
  declare varkind: number;
  get oInst(): number {
    return this._buffer.readUInt32LE(ctypes.sizeof(c_void_p) === 8 ? 16 : 8);
  }
}

export class TYPEATTR extends ctypes.Structure {
  static _fields_ = [
    ["guid", GUID],
    ["lcid", c_uint32],
    ["dwReserved", c_uint32],
    ["memidConstructor", c_int32],
    ["memidDestructor", c_int32],
    ["lpstrSchema", c_void_p],
    ["cbSizeInstance", c_uint32],
    ["typekind", c_int32],
    ["cFuncs", c_uint16],
    ["cVars", c_uint16],
    ["cImplTypes", c_uint16],
    ["cbSizeVft", c_uint16],
    ["cbAlignment", c_uint16],
    ["wTypeFlags", c_uint16],
    ["wMajorVerNum", c_uint16],
    ["wMinorVerNum", c_uint16],
    ["tdescAlias", TYPEDESC],
    ["idldescType", PARAMDESC],
  ] as const;
  declare guid: GUID;
  declare lcid: number;
  declare dwReserved: number;
  declare memidConstructor: number;
  declare memidDestructor: number;
  declare lpstrSchema: bigint | null;
  declare cbSizeInstance: number;
  declare typekind: number;
  declare cFuncs: number;
  declare cVars: number;
  declare cImplTypes: number;
  declare cbSizeVft: number;
  declare cbAlignment: number;
  declare wTypeFlags: number;
  declare wMajorVerNum: number;
  declare wMinorVerNum: number;
  declare tdescAlias: TYPEDESC;
  declare idldescType: PARAMDESC;
}

export class TLIBATTR extends ctypes.Structure {
  static _fields_ = [
    ["guid", GUID],
    ["lcid", c_uint32],
    ["syskind", c_int32],
    ["wMajorVerNum", c_uint16],
    ["wMinorVerNum", c_uint16],
    ["wLibFlags", c_uint16],
  ] as const;
  declare guid: GUID;
  declare lcid: number;
  declare syskind: number;
  declare wMajorVerNum: number;
  declare wMinorVerNum: number;
  declare wLibFlags: number;
}

/**
 * Type information. `Release*` methods return nothing; BSTR out parameters
 * (`POINTER(c_void_p)`) must be freed with `SysFreeString`.
 */
export const ITypeInfo = COMInterface({
  name: "ITypeInfo",
  iid: "{00020401-0000-0000-C000-000000000046}",
  extends: IUnknown,
  methods: [
    STDMETHOD(HRESULT, "GetTypeAttr", [POINTER(c_void_p)]),
    STDMETHOD(HRESULT, "GetTypeComp", [POINTER(c_void_p)]),
    STDMETHOD(HRESULT, "GetFuncDesc", [c_uint32, POINTER(c_void_p)]),
    STDMETHOD(HRESULT, "GetVarDesc", [c_uint32, POINTER(c_void_p)]),
    STDMETHOD(HRESULT, "GetNames", [c_int32, c_void_p, c_uint32, POINTER(c_uint32)]),
    STDMETHOD(HRESULT, "GetRefTypeOfImplType", [c_uint32, POINTER(c_uint32)]),
    STDMETHOD(HRESULT, "GetImplTypeFlags", [c_uint32, POINTER(c_int32)]),
    STDMETHOD(HRESULT, "GetIDsOfNames", [c_void_p, c_uint32, POINTER(c_int32)]),
    STDMETHOD(HRESULT, "Invoke", [c_void_p, c_int32, c_uint16, POINTER(DISPPARAMS), POINTER(VARIANT), POINTER(EXCEPINFO), POINTER(c_uint32)]),
    STDMETHOD(HRESULT, "GetDocumentation", [c_int32, POINTER(c_void_p), POINTER(c_void_p), POINTER(c_uint32), POINTER(c_void_p)]),
    STDMETHOD(HRESULT, "GetDllEntry", [c_int32, c_int32, POINTER(c_void_p), POINTER(c_void_p), POINTER(c_uint16)]),
    STDMETHOD(HRESULT, "GetRefTypeInfo", [c_uint32, POINTER(c_void_p)]),
    STDMETHOD(HRESULT, "AddressOfMember", [c_int32, c_int32, POINTER(c_void_p)]),
    STDMETHOD(HRESULT, "CreateInstance", [c_void_p, POINTER(GUID), POINTER(c_void_p)]),
    STDMETHOD(HRESULT, "GetMops", [c_int32, POINTER(c_void_p)]),
    STDMETHOD(HRESULT, "GetContainingTypeLib", [POINTER(c_void_p), POINTER(c_uint32)]),
    STDMETHOD(null, "ReleaseTypeAttr", [c_void_p]),
    STDMETHOD(null, "ReleaseFuncDesc", [c_void_p]),
    STDMETHOD(null, "ReleaseVarDesc", [c_void_p]),
  ],
});

export const ITypeLib = COMInterface({
  name: "ITypeLib",
  iid: "{00020402-0000-0000-C000-000000000046}",
  extends: IUnknown,
  methods: [
    STDMETHOD(c_uint32, "GetTypeInfoCount"),
    STDMETHOD(HRESULT, "GetTypeInfo", [c_uint32, POINTER(c_void_p)]),
    STDMETHOD(HRESULT, "GetTypeInfoType", [c_uint32, POINTER(c_int32)]),
    STDMETHOD(HRESULT, "GetTypeInfoOfGuid", [POINTER(GUID), POINTER(c_void_p)]),
    STDMETHOD(HRESULT, "GetLibAttr", [POINTER(c_void_p)]),
    STDMETHOD(HRESULT, "GetTypeComp", [POINTER(c_void_p)]),
    STDMETHOD(HRESULT, "GetDocumentation", [c_int32, POINTER(c_void_p), POINTER(c_void_p), POINTER(c_uint32), POINTER(c_void_p)]),
    STDMETHOD(HRESULT, "IsName", [c_wchar_p, c_uint32, POINTER(c_int32)]),
    STDMETHOD(HRESULT, "FindName", [c_wchar_p, c_uint32, c_void_p, POINTER(c_int32), POINTER(c_uint16)]),
    STDMETHOD(null, "ReleaseTLibAttr", [c_void_p]),
  ],
});

export const ITypeComp = COMInterface({
  name: "ITypeComp",
  iid: "{00020403-0000-0000-C000-000000000046}",
  extends: IUnknown,
  methods: [STDMETHOD(HRESULT, "Bind", [c_wchar_p, c_uint32, c_uint16, POINTER(c_void_p), POINTER(c_int32), POINTER(c_void_p)]), STDMETHOD(HRESULT, "BindType", [c_wchar_p, c_uint32, POINTER(c_void_p), POINTER(c_void_p)])],
});

/** Class information of a coclass instance. */
export const IProvideClassInfo = COMInterface({
  name: "IProvideClassInfo",
  iid: "{B196B283-BAB4-101A-B69C-00AA00341D07}",
  extends: IUnknown,
  methods: [COMMETHOD([], HRESULT, "GetClassInfo", [["out"], ITypeInfo, "ppTI"])],
});

/**
 * Builds a type description (`ICreateTypeLib::CreateTypeInfo`). Methods taking
 * structures (`FUNCDESC`, `VARDESC`, `TYPEDESC`, `IDLDESC`) and name arrays take raw arguments.
 */
export const ICreateTypeInfo = COMInterface({
  name: "ICreateTypeInfo",
  iid: "{00020405-0000-0000-C000-000000000046}",
  extends: IUnknown,
  methods: [
    COMMETHOD([], HRESULT, "SetGuid", [["in"], REFGUID, "guid"]),
    COMMETHOD([], HRESULT, "SetTypeFlags", [["in"], c_uint32, "uTypeFlags"]),
    COMMETHOD([], HRESULT, "SetDocString", [["in"], LPOLESTR, "pStrDoc"]),
    COMMETHOD([], HRESULT, "SetHelpContext", [["in"], c_uint32, "dwHelpContext"]),
    COMMETHOD([], HRESULT, "SetVersion", [["in"], c_uint16, "wMajorVerNum"], [["in"], c_uint16, "wMinorVerNum"]),
    COMMETHOD([], HRESULT, "AddRefTypeInfo", [["in"], ITypeInfo, "pTInfo"], [["out"], c_uint32, "phRefType"]),
    STDMETHOD(HRESULT, "AddFuncDesc", [c_uint32, POINTER(FUNCDESC)]),
    COMMETHOD([], HRESULT, "AddImplType", [["in"], c_uint32, "index"], [["in"], c_uint32, "hRefType"]),
    COMMETHOD([], HRESULT, "SetImplTypeFlags", [["in"], c_uint32, "index"], [["in"], c_int32, "implTypeFlags"]),
    COMMETHOD([], HRESULT, "SetAlignment", [["in"], c_uint16, "cbAlignment"]),
    COMMETHOD([], HRESULT, "SetSchema", [["in"], LPOLESTR, "pStrSchema"]),
    STDMETHOD(HRESULT, "AddVarDesc", [c_uint32, POINTER(VARDESC)]),
    // rgszNames is LPOLESTR*: an array of wide string pointers.
    STDMETHOD(HRESULT, "SetFuncAndParamNames", [c_uint32, c_void_p, c_uint32]),
    COMMETHOD([], HRESULT, "SetVarName", [["in"], c_uint32, "index"], [["in"], LPOLESTR, "szName"]),
    STDMETHOD(HRESULT, "SetTypeDescAlias", [POINTER(TYPEDESC)]),
    COMMETHOD([], HRESULT, "DefineFuncAsDllEntry", [["in"], c_uint32, "index"], [["in"], LPOLESTR, "szDllName"], [["in"], LPOLESTR, "szProcName"]),
    COMMETHOD([], HRESULT, "SetFuncDocString", [["in"], c_uint32, "index"], [["in"], LPOLESTR, "szDocString"]),
    COMMETHOD([], HRESULT, "SetVarDocString", [["in"], c_uint32, "index"], [["in"], LPOLESTR, "szDocString"]),
    COMMETHOD([], HRESULT, "SetFuncHelpContext", [["in"], c_uint32, "index"], [["in"], c_uint32, "dwHelpContext"]),
    COMMETHOD([], HRESULT, "SetVarHelpContext", [["in"], c_uint32, "index"], [["in"], c_uint32, "dwHelpContext"]),
    COMMETHOD([], HRESULT, "SetMops", [["in"], c_uint32, "index"], [["in"], BSTR, "bstrMops"]),
    STDMETHOD(HRESULT, "SetTypeIdldesc", [POINTER(PARAMDESC)]),
    COMMETHOD([], HRESULT, "LayOut"),
  ],
});

/** Builds a type library file (`CreateTypeLib2`). Python equivalent: `comtypes.typeinfo.ICreateTypeLib`. */
export const ICreateTypeLib = COMInterface({
  name: "ICreateTypeLib",
  iid: "{00020406-0000-0000-C000-000000000046}",
  extends: IUnknown,
  methods: [
    COMMETHOD([], HRESULT, "CreateTypeInfo", [["in"], LPOLESTR, "szName"], [["in"], c_int32, "tkind"], [["out"], ICreateTypeInfo, "ppCTInfo"]),
    COMMETHOD([], HRESULT, "SetName", [["in"], LPOLESTR, "szName"]),
    COMMETHOD([], HRESULT, "SetVersion", [["in"], c_uint16, "wMajorVerNum"], [["in"], c_uint16, "wMinorVerNum"]),
    COMMETHOD([], HRESULT, "SetGuid", [["in"], REFGUID, "guid"]),
    COMMETHOD([], HRESULT, "SetDocString", [["in"], LPOLESTR, "szDoc"]),
    COMMETHOD([], HRESULT, "SetHelpFileName", [["in"], LPOLESTR, "szHelpFileName"]),
    COMMETHOD([], HRESULT, "SetHelpContext", [["in"], c_uint32, "dwHelpContext"]),
    COMMETHOD([], HRESULT, "SetLcid", [["in"], c_uint32, "lcid"]),
    COMMETHOD([], HRESULT, "SetLibFlags", [["in"], c_uint32, "uLibFlags"]),
    COMMETHOD([], HRESULT, "SaveAllChanges"),
  ],
});

export const ICreateTypeLib2 = COMInterface({
  name: "ICreateTypeLib2",
  iid: "{0002040F-0000-0000-C000-000000000046}",
  extends: ICreateTypeLib,
  methods: [
    COMMETHOD([], HRESULT, "DeleteTypeInfo", [["in"], LPOLESTR, "szName"]),
    COMMETHOD([], HRESULT, "SetCustData", [["in"], REFGUID, "guid"], [["in"], POINTER(VARIANT), "pVarVal"]),
    COMMETHOD([], HRESULT, "SetHelpStringContext", [["in"], c_uint32, "dwHelpStringContext"]),
    COMMETHOD([], HRESULT, "SetHelpStringDll", [["in"], LPOLESTR, "szFileName"]),
  ],
});
