// HRESULT values and helpers (winerror.h). Values are signed 32-bit integers,
// as returned by node-ctypes for `HRESULT` / `c_int32`.

export const S_OK = 0;
export const S_FALSE = 1;

export const E_UNEXPECTED = 0x8000ffff | 0;
export const E_NOTIMPL = 0x80004001 | 0;
export const E_NOINTERFACE = 0x80004002 | 0;
export const E_POINTER = 0x80004003 | 0;
export const E_ABORT = 0x80004004 | 0;
export const E_FAIL = 0x80004005 | 0;
export const E_PENDING = 0x8000000a | 0;
export const E_ACCESSDENIED = 0x80070005 | 0;
export const E_HANDLE = 0x80070006 | 0;
export const E_OUTOFMEMORY = 0x8007000e | 0;
export const E_INVALIDARG = 0x80070057 | 0;

export const CLASS_E_NOAGGREGATION = 0x80040110 | 0;
export const CLASS_E_CLASSNOTAVAILABLE = 0x80040111 | 0;
export const REGDB_E_CLASSNOTREG = 0x80040154 | 0;
export const CO_E_NOTINITIALIZED = 0x800401f0 | 0;
export const CO_E_CLASSSTRING = 0x800401f3 | 0;
export const CO_E_SERVER_EXEC_FAILURE = 0x80080005 | 0;
export const MK_E_UNAVAILABLE = 0x800401e3 | 0;
export const MK_E_SYNTAX = 0x800401e4 | 0;
export const RPC_E_CHANGED_MODE = 0x80010106 | 0;
export const RPC_E_WRONG_THREAD = 0x8001010e | 0;
export const RPC_E_DISCONNECTED = 0x80010108 | 0;

export const DISP_E_UNKNOWNINTERFACE = 0x80020001 | 0;
export const DISP_E_MEMBERNOTFOUND = 0x80020003 | 0;
export const DISP_E_PARAMNOTFOUND = 0x80020004 | 0;
export const DISP_E_TYPEMISMATCH = 0x80020005 | 0;
export const DISP_E_UNKNOWNNAME = 0x80020006 | 0;
export const DISP_E_NONAMEDARGS = 0x80020007 | 0;
export const DISP_E_BADVARTYPE = 0x80020008 | 0;
export const DISP_E_EXCEPTION = 0x80020009 | 0;
export const DISP_E_OVERFLOW = 0x8002000a | 0;
export const DISP_E_BADINDEX = 0x8002000b | 0;
export const DISP_E_UNKNOWNLCID = 0x8002000c | 0;
export const DISP_E_ARRAYISLOCKED = 0x8002000d | 0;
export const DISP_E_BADPARAMCOUNT = 0x8002000e | 0;
export const DISP_E_PARAMNOTOPTIONAL = 0x8002000f | 0;
export const TYPE_E_ELEMENTNOTFOUND = 0x8002802b | 0;
export const TYPE_E_CANTLOADLIBRARY = 0x80029c4a | 0;

export const CONNECT_E_NOCONNECTION = 0x80040200 | 0;
export const CONNECT_E_ADVISELIMIT = 0x80040201 | 0;
export const CONNECT_E_CANNOTCONNECT = 0x80040202 | 0;

export const FACILITY_WIN32 = 7;
export const FACILITY_ITF = 4;
export const SEVERITY_SUCCESS = 0;
export const SEVERITY_ERROR = 1;

/** Normalize a number or bigint HRESULT to a signed 32-bit integer. */
export function toHRESULT(value: number | bigint): number {
  return typeof value === "bigint" ? Number(BigInt.asIntN(32, value)) : value | 0;
}

export function SUCCEEDED(hr: number | bigint): boolean {
  return toHRESULT(hr) >= 0;
}

export function FAILED(hr: number | bigint): boolean {
  return toHRESULT(hr) < 0;
}

export function MAKE_HRESULT(severity: number, facility: number, code: number): number {
  return ((severity & 1) << 31) | ((facility & 0x1fff) << 16) | (code & 0xffff);
}

export function HRESULT_FROM_WIN32(code: number): number {
  const value = code | 0;
  return value <= 0 ? value : MAKE_HRESULT(SEVERITY_ERROR, FACILITY_WIN32, value);
}

export function HRESULT_CODE(hr: number): number {
  return hr & 0xffff;
}

export function HRESULT_FACILITY(hr: number): number {
  return (hr >> 16) & 0x1fff;
}

export function HRESULT_SEVERITY(hr: number): number {
  return (hr >>> 31) & 1;
}

/** Format as `0x8000FFFF`. */
export function formatHRESULT(hr: number | bigint): string {
  return `0x${(toHRESULT(hr) >>> 0).toString(16).toUpperCase().padStart(8, "0")}`;
}
