// VARENUM (wtypes.h).
export const VT_EMPTY = 0;
export const VT_NULL = 1;
export const VT_I2 = 2;
export const VT_I4 = 3;
export const VT_R4 = 4;
export const VT_R8 = 5;
export const VT_CY = 6;
export const VT_DATE = 7;
export const VT_BSTR = 8;
export const VT_DISPATCH = 9;
export const VT_ERROR = 10;
export const VT_BOOL = 11;
export const VT_VARIANT = 12;
export const VT_UNKNOWN = 13;
export const VT_DECIMAL = 14;
export const VT_I1 = 16;
export const VT_UI1 = 17;
export const VT_UI2 = 18;
export const VT_UI4 = 19;
export const VT_I8 = 20;
export const VT_UI8 = 21;
export const VT_INT = 22;
export const VT_UINT = 23;
export const VT_VOID = 24;
export const VT_HRESULT = 25;
export const VT_PTR = 26;
export const VT_SAFEARRAY = 27;
export const VT_CARRAY = 28;
export const VT_USERDEFINED = 29;
export const VT_LPSTR = 30;
export const VT_LPWSTR = 31;
export const VT_RECORD = 36;
export const VT_INT_PTR = 37;
export const VT_UINT_PTR = 38;
export const VT_ARRAY = 0x2000;
export const VT_BYREF = 0x4000;
export const VT_TYPEMASK = 0xfff;

export const VARIANT_TRUE = -1;
export const VARIANT_FALSE = 0;

// DISPIDs.
export const DISPID_UNKNOWN = -1;
export const DISPID_VALUE = 0;
export const DISPID_PROPERTYPUT = -3;
export const DISPID_NEWENUM = -4;
export const DISPID_EVALUATE = -5;
export const DISPID_CONSTRUCTOR = -6;
export const DISPID_DESTRUCTOR = -7;
export const DISPID_COLLECT = -8;
export const MEMBERID_NIL = -1;

// IDispatch::Invoke flags (wFlags).
export const DISPATCH_METHOD = 0x1;
export const DISPATCH_PROPERTYGET = 0x2;
export const DISPATCH_PROPERTYPUT = 0x4;
export const DISPATCH_PROPERTYPUTREF = 0x8;

// Locales.
export const LOCALE_SYSTEM_DEFAULT = 0x800;
export const LOCALE_USER_DEFAULT = 0x400;
export const LOCALE_NEUTRAL = 0;

// VariantChangeType flags.
export const VARIANT_NOVALUEPROP = 0x1;
export const VARIANT_ALPHABOOL = 0x2;
export const VARIANT_NOUSEROVERRIDE = 0x4;
export const VARIANT_LOCALBOOL = 0x10;

// Type information (oaidl.h).
export const TKIND_ENUM = 0;
export const TKIND_RECORD = 1;
export const TKIND_MODULE = 2;
export const TKIND_INTERFACE = 3;
export const TKIND_DISPATCH = 4;
export const TKIND_COCLASS = 5;
export const TKIND_ALIAS = 6;
export const TKIND_UNION = 7;

export const CC_CDECL = 1;
export const CC_STDCALL = 4;

export const FUNC_VIRTUAL = 0;
export const FUNC_PUREVIRTUAL = 1;
export const FUNC_NONVIRTUAL = 2;
export const FUNC_STATIC = 3;
export const FUNC_DISPATCH = 4;

export const INVOKE_FUNC = 1;
export const INVOKE_PROPERTYGET = 2;
export const INVOKE_PROPERTYPUT = 4;
export const INVOKE_PROPERTYPUTREF = 8;

export const VAR_PERINSTANCE = 0;
export const VAR_STATIC = 1;
export const VAR_CONST = 2;
export const VAR_DISPATCH = 3;

export const PARAMFLAG_NONE = 0;
export const PARAMFLAG_FIN = 0x1;
export const PARAMFLAG_FOUT = 0x2;
export const PARAMFLAG_FLCID = 0x4;
export const PARAMFLAG_FRETVAL = 0x8;
export const PARAMFLAG_FOPT = 0x10;
export const PARAMFLAG_FHASDEFAULT = 0x20;
export const PARAMFLAG_FHASCUSTDATA = 0x40;

export const IMPLTYPEFLAG_FDEFAULT = 0x1;
export const IMPLTYPEFLAG_FSOURCE = 0x2;
export const IMPLTYPEFLAG_FRESTRICTED = 0x4;
export const IMPLTYPEFLAG_FDEFAULTVTABLE = 0x8;

export const TYPEFLAG_FAPPOBJECT = 0x1;
export const TYPEFLAG_FCANCREATE = 0x2;
export const TYPEFLAG_FLICENSED = 0x4;
export const TYPEFLAG_FPREDECLID = 0x8;
export const TYPEFLAG_FHIDDEN = 0x10;
export const TYPEFLAG_FCONTROL = 0x20;
export const TYPEFLAG_FDUAL = 0x40;
export const TYPEFLAG_FNONEXTENSIBLE = 0x80;
export const TYPEFLAG_FOLEAUTOMATION = 0x100;
export const TYPEFLAG_FRESTRICTED = 0x200;
export const TYPEFLAG_FAGGREGATABLE = 0x400;
export const TYPEFLAG_FREPLACEABLE = 0x800;
export const TYPEFLAG_FDISPATCHABLE = 0x1000;

export const FUNCFLAG_FRESTRICTED = 0x1;
export const FUNCFLAG_FSOURCE = 0x2;
export const FUNCFLAG_FBINDABLE = 0x4;
export const FUNCFLAG_FHIDDEN = 0x40;

export const SYS_WIN16 = 0;
export const SYS_WIN32 = 1;
export const SYS_MAC = 2;
export const SYS_WIN64 = 3;

/** Name of a VARTYPE, e.g. `VT_ARRAY | VT_BSTR` for `0x2008`. */
export function vartypeName(vt: number): string {
  const names: Record<number, string> = {
    [VT_EMPTY]: "VT_EMPTY",
    [VT_NULL]: "VT_NULL",
    [VT_I2]: "VT_I2",
    [VT_I4]: "VT_I4",
    [VT_R4]: "VT_R4",
    [VT_R8]: "VT_R8",
    [VT_CY]: "VT_CY",
    [VT_DATE]: "VT_DATE",
    [VT_BSTR]: "VT_BSTR",
    [VT_DISPATCH]: "VT_DISPATCH",
    [VT_ERROR]: "VT_ERROR",
    [VT_BOOL]: "VT_BOOL",
    [VT_VARIANT]: "VT_VARIANT",
    [VT_UNKNOWN]: "VT_UNKNOWN",
    [VT_DECIMAL]: "VT_DECIMAL",
    [VT_I1]: "VT_I1",
    [VT_UI1]: "VT_UI1",
    [VT_UI2]: "VT_UI2",
    [VT_UI4]: "VT_UI4",
    [VT_I8]: "VT_I8",
    [VT_UI8]: "VT_UI8",
    [VT_INT]: "VT_INT",
    [VT_UINT]: "VT_UINT",
    [VT_VOID]: "VT_VOID",
    [VT_HRESULT]: "VT_HRESULT",
    [VT_PTR]: "VT_PTR",
    [VT_SAFEARRAY]: "VT_SAFEARRAY",
    [VT_CARRAY]: "VT_CARRAY",
    [VT_USERDEFINED]: "VT_USERDEFINED",
    [VT_LPSTR]: "VT_LPSTR",
    [VT_LPWSTR]: "VT_LPWSTR",
    [VT_RECORD]: "VT_RECORD",
    [VT_INT_PTR]: "VT_INT_PTR",
    [VT_UINT_PTR]: "VT_UINT_PTR",
  };
  const base = names[vt & VT_TYPEMASK] ?? `0x${(vt & VT_TYPEMASK).toString(16)}`;
  return [vt & VT_BYREF ? "VT_BYREF" : null, vt & VT_ARRAY ? "VT_ARRAY" : null, base].filter(Boolean).join(" | ");
}
