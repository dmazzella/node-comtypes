import * as ctypes from "node-ctypes";
import { GUID } from "./guid.ts";
import type { COMErrorDetails } from "./errors.ts";
import { oleaut32, takeBSTR } from "./low-level/native.ts";
import { functionCache, rawQueryInterface, rawRelease, vtableEntry } from "./low-level/vtable.ts";

const { c_int32, c_uint32, c_void_p } = ctypes;

const IID_ISupportErrorInfo = new GUID("{DF0B3D60-548F-101B-8E65-08002B2BD119}");
// IErrorInfo getters and ISupportErrorInfo::InterfaceSupportsErrorInfo share
// the HRESULT (this, pointer) signature.
const PointerArgAt = functionCache(() => ctypes.WINFUNCTYPE(c_int32, c_void_p, c_void_p));

function readBSTR(errorInfo: bigint, slot: number): string | null {
  const out = new c_void_p();
  if (PointerArgAt(vtableEntry(errorInfo, slot))(errorInfo, ctypes.byref(out)) < 0) return null;
  return takeBSTR(out.value);
}

/**
 * Take the current thread's error object (`GetErrorInfo`), clearing it.
 * Returns `null` when no error object is set.
 */
export function takeErrorInfo(): COMErrorDetails | null {
  const out = new c_void_p();
  if (oleaut32().GetErrorInfo(0, ctypes.byref(out)) !== 0 || !out.value) return null;
  const errorInfo = out.value;
  try {
    const guid = new GUID();
    const hasGuid = PointerArgAt(vtableEntry(errorInfo, 3))(errorInfo, guid._buffer) >= 0;
    const context = new c_uint32();
    const hasContext = PointerArgAt(vtableEntry(errorInfo, 7))(errorInfo, ctypes.byref(context)) >= 0;
    return {
      guid: hasGuid && !guid.isNull() ? guid : null,
      source: readBSTR(errorInfo, 4),
      description: readBSTR(errorInfo, 5),
      helpFile: readBSTR(errorInfo, 6),
      helpContext: hasContext ? context.value : 0,
    };
  } finally {
    rawRelease(errorInfo);
  }
}

/**
 * Error details for a failed method of interface `iid` called on `object`:
 * read only if the object declares `ISupportErrorInfo` support for `iid`,
 * as Python ctypes/comtypes do.
 */
export function errorInfoFor(object: bigint, iid: GUID): COMErrorDetails | null {
  try {
    const [hr, support] = rawQueryInterface(object, IID_ISupportErrorInfo);
    if (hr < 0 || !support) return null;
    let supported: boolean;
    try {
      supported = PointerArgAt(vtableEntry(support, 3))(support, iid._buffer) === 0;
    } finally {
      rawRelease(support);
    }
    return supported ? takeErrorInfo() : null;
  } catch {
    return null;
  }
}
