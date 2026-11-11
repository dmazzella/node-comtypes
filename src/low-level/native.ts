import * as ctypes from "node-ctypes";

const { c_int32, c_uint16, c_uint32, c_void_p, c_wchar_p, c_size_t } = ctypes;

export const POINTER_SIZE = ctypes.sizeof(c_void_p);
export const isWindows = process.platform === "win32";

// HRESULT-returning functions are bound with `c_int32` (not `HRESULT`): the
// raw value is returned and checked by the caller, which builds a COMError.
function bindOle32(library: ctypes.WinDLL) {
  return {
    CoInitializeEx: library.func("CoInitializeEx", c_int32, [c_void_p, c_uint32]),
    CoUninitialize: library.func("CoUninitialize", null, []),
    CoGetApartmentType: library.func("CoGetApartmentType", c_int32, [c_void_p, c_void_p]),
    CoCreateInstance: library.func("CoCreateInstance", c_int32, [c_void_p, c_void_p, c_uint32, c_void_p, c_void_p]),
    CoCreateInstanceEx: library.func("CoCreateInstanceEx", c_int32, [c_void_p, c_void_p, c_uint32, c_void_p, c_uint32, c_void_p]),
    CoGetClassObject: library.func("CoGetClassObject", c_int32, [c_void_p, c_uint32, c_void_p, c_void_p, c_void_p]),
    CoCreateGuid: library.func("CoCreateGuid", c_int32, [c_void_p]),
    CLSIDFromProgID: library.func("CLSIDFromProgID", c_int32, [c_wchar_p, c_void_p]),
    CLSIDFromString: library.func("CLSIDFromString", c_int32, [c_wchar_p, c_void_p]),
    ProgIDFromCLSID: library.func("ProgIDFromCLSID", c_int32, [c_void_p, c_void_p]),
    CoTaskMemAlloc: library.func("CoTaskMemAlloc", c_void_p, [c_size_t]),
    CoTaskMemRealloc: library.func("CoTaskMemRealloc", c_void_p, [c_void_p, c_size_t]),
    CoTaskMemFree: library.func("CoTaskMemFree", null, [c_void_p]),
    CoGetObject: library.func("CoGetObject", c_int32, [c_wchar_p, c_void_p, c_void_p, c_void_p]),
    CoMarshalInterThreadInterfaceInStream: library.func("CoMarshalInterThreadInterfaceInStream", c_int32, [c_void_p, c_void_p, c_void_p]),
    CoGetInterfaceAndReleaseStream: library.func("CoGetInterfaceAndReleaseStream", c_int32, [c_void_p, c_void_p, c_void_p]),
    CoReleaseMarshalData: library.func("CoReleaseMarshalData", c_int32, [c_void_p]),
    CoWaitForMultipleHandles: library.func("CoWaitForMultipleHandles", c_int32, [c_uint32, c_uint32, c_uint32, c_void_p, c_void_p]),
    CoRegisterClassObject: library.func("CoRegisterClassObject", c_int32, [c_void_p, c_void_p, c_uint32, c_uint32, c_void_p]),
    CoRevokeClassObject: library.func("CoRevokeClassObject", c_int32, [c_uint32]),
  };
}

function bindOleAut32(library: ctypes.WinDLL) {
  return {
    GetErrorInfo: library.func("GetErrorInfo", c_int32, [c_uint32, c_void_p]),
    SetErrorInfo: library.func("SetErrorInfo", c_int32, [c_uint32, c_void_p]),
    CreateErrorInfo: library.func("CreateErrorInfo", c_int32, [c_void_p]),
    SysAllocString: library.func("SysAllocString", c_void_p, [c_wchar_p]),
    SysAllocStringLen: library.func("SysAllocStringLen", c_void_p, [c_void_p, c_uint32]),
    SysFreeString: library.func("SysFreeString", null, [c_void_p]),
    SysStringLen: library.func("SysStringLen", c_uint32, [c_void_p]),
    VariantInit: library.func("VariantInit", null, [c_void_p]),
    VariantClear: library.func("VariantClear", c_int32, [c_void_p]),
    VariantCopy: library.func("VariantCopy", c_int32, [c_void_p, c_void_p]),
    VariantChangeType: library.func("VariantChangeType", c_int32, [c_void_p, c_void_p, c_uint16, c_uint16]),
    VariantChangeTypeEx: library.func("VariantChangeTypeEx", c_int32, [c_void_p, c_void_p, c_uint32, c_uint16, c_uint16]),
    SafeArrayCreate: library.func("SafeArrayCreate", c_void_p, [c_uint16, c_uint32, c_void_p]),
    SafeArrayDestroy: library.func("SafeArrayDestroy", c_int32, [c_void_p]),
    SafeArrayGetDim: library.func("SafeArrayGetDim", c_uint32, [c_void_p]),
    SafeArrayGetElemsize: library.func("SafeArrayGetElemsize", c_uint32, [c_void_p]),
    SafeArrayGetLBound: library.func("SafeArrayGetLBound", c_int32, [c_void_p, c_uint32, c_void_p]),
    SafeArrayGetUBound: library.func("SafeArrayGetUBound", c_int32, [c_void_p, c_uint32, c_void_p]),
    SafeArrayGetVartype: library.func("SafeArrayGetVartype", c_int32, [c_void_p, c_void_p]),
    SafeArrayAccessData: library.func("SafeArrayAccessData", c_int32, [c_void_p, c_void_p]),
    SafeArrayUnaccessData: library.func("SafeArrayUnaccessData", c_int32, [c_void_p]),
    GetActiveObject: library.func("GetActiveObject", c_int32, [c_void_p, c_void_p, c_void_p]),
    LoadTypeLibEx: library.func("LoadTypeLibEx", c_int32, [c_wchar_p, c_uint32, c_void_p]),
    LoadRegTypeLib: library.func("LoadRegTypeLib", c_int32, [c_void_p, c_uint16, c_uint16, c_uint32, c_void_p]),
    GetRecordInfoFromTypeInfo: library.func("GetRecordInfoFromTypeInfo", c_int32, [c_void_p, c_void_p]),
    SafeArrayCreateEx: library.func("SafeArrayCreateEx", c_void_p, [c_uint16, c_uint32, c_void_p, c_void_p]),
    SafeArrayGetRecordInfo: library.func("SafeArrayGetRecordInfo", c_int32, [c_void_p, c_void_p]),
    QueryPathOfRegTypeLib: library.func("QueryPathOfRegTypeLib", c_int32, [c_void_p, c_uint16, c_uint16, c_uint32, c_void_p]),
    CreateTypeLib2: library.func("CreateTypeLib2", c_int32, [c_int32, c_wchar_p, c_void_p]),
    RegisterTypeLib: library.func("RegisterTypeLib", c_int32, [c_void_p, c_wchar_p, c_wchar_p]),
    RegisterTypeLibForUser: library.func("RegisterTypeLibForUser", c_int32, [c_void_p, c_wchar_p, c_wchar_p]),
    UnRegisterTypeLib: library.func("UnRegisterTypeLib", c_int32, [c_void_p, c_uint16, c_uint16, c_uint32, c_int32]),
    UnRegisterTypeLibForUser: library.func("UnRegisterTypeLibForUser", c_int32, [c_void_p, c_uint16, c_uint16, c_uint32, c_int32]),
  };
}

export type Ole32 = ReturnType<typeof bindOle32>;
export type OleAut32 = ReturnType<typeof bindOleAut32>;

let ole32Functions: Ole32 | null = null;
let oleaut32Functions: OleAut32 | null = null;

function requireWindows(): void {
  if (!isWindows) throw new Error(`COM is only available on Windows (current platform: ${process.platform})`);
}

/** ole32.dll functions, loaded on first use. */
export function ole32(): Ole32 {
  requireWindows();
  return (ole32Functions ??= bindOle32(new ctypes.WinDLL("ole32.dll")));
}

/** oleaut32.dll functions, loaded on first use. */
export function oleaut32(): OleAut32 {
  requireWindows();
  return (oleaut32Functions ??= bindOleAut32(new ctypes.WinDLL("oleaut32.dll")));
}

/** Allocate a BSTR (embedded NULs preserved). Free it with `SysFreeString`. */
export function allocBSTR(text: string): bigint {
  const pointer = oleaut32().SysAllocStringLen(Buffer.from(text, "utf16le"), text.length);
  if (!pointer) throw new RangeError("SysAllocStringLen failed: out of memory");
  return pointer;
}

/** Read a BSTR without freeing it (`null` for a NULL BSTR, which COM treats as ""). */
export function readBSTR(bstr: bigint | null): string | null {
  if (!bstr) return null;
  // The length prefix counts bytes; embedded NULs are part of the string.
  const length = (ctypes.readValue(bstr - 4n, c_uint32) as number) >>> 1;
  return length ? (ctypes.wstring_at(bstr, length) ?? "") : "";
}

/** Read a BSTR and free it. */
export function takeBSTR(bstr: bigint | null): string | null {
  if (!bstr) return null;
  try {
    return readBSTR(bstr);
  } finally {
    oleaut32().SysFreeString(bstr);
  }
}

/**
 * A copy of the structure at `address`, for data that is only read (type
 * descriptors): cheaper than a view over native memory (`from_address`).
 * Pointers inside still refer to native memory.
 */
export function copyStruct<T>(type: new (...args: any[]) => T, address: bigint): T {
  const bytes = Buffer.allocUnsafe(ctypes.sizeof(type as unknown as ctypes.FFIType));
  ctypes.memmove(bytes, address, bytes.length);
  return (type as unknown as { from_buffer(source: Buffer): T }).from_buffer(bytes);
}

/** Copy `size` bytes of native memory at `address`. */
export function readMemory(address: bigint, size: number): Buffer {
  const buffer = Buffer.alloc(size);
  if (size) ctypes.memmove(buffer, address, size);
  return buffer;
}

/** Copy `data` to native memory at `address`. */
export function writeMemory(address: bigint, data: Uint8Array): void {
  if (data.length) ctypes.memmove(address, data, data.length);
}

/** Read a CoTaskMem-allocated wide string and free it. */
export function takeCoTaskMemString(pointer: bigint | null): string | null {
  if (!pointer) return null;
  try {
    return ctypes.wstring_at(pointer);
  } finally {
    ole32().CoTaskMemFree(pointer);
  }
}
