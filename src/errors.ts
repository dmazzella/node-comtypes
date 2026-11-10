import * as ctypes from "node-ctypes";
import type { GUID } from "./guid.ts";
import { formatHRESULT, toHRESULT } from "./hresult.ts";
import { isWindows } from "./low-level/native.ts";

/** Rich error information from `IErrorInfo` (Python: the `details` tuple). */
export interface COMErrorDetails {
  readonly description: string | null;
  readonly source: string | null;
  readonly helpFile: string | null;
  readonly helpContext: number;
  readonly guid: GUID | null;
  /** `EXCEPINFO.scode` of a `DISP_E_EXCEPTION` (0 when the server set only `wCode`). */
  readonly scode?: number;
  /** Index of the offending argument for `DISP_E_TYPEMISMATCH` / `DISP_E_PARAMNOTFOUND`. */
  readonly argumentIndex?: number;
}

export interface COMErrorOptions {
  /** System message for the HRESULT; defaults to `FormatMessageW`. */
  readonly text?: string;
  readonly details?: COMErrorDetails | null;
  /** Failing method, e.g. `IShellLinkW::SetPath` or `CoCreateInstance`. */
  readonly method?: string | null;
  readonly cause?: unknown;
}

/** System message for an HRESULT (`FormatMessageW`), or `null` when the system has none. */
export function systemMessage(hr: number): string | null {
  if (!isWindows) return null;
  try {
    const text = ctypes.FormatError(hr).trim();
    if (text && !text.startsWith("<") && !text.startsWith("Unknown error")) return text;
  } catch {
    // Try the Visual Basic messages below.
  }
  return hr >>> 16 === 0x800a ? visualBasicMessage(hr & 0xffff) : null;
}

const LOAD_LIBRARY_AS_DATAFILE = 0x2;
const LOAD_LIBRARY_AS_IMAGE_RESOURCE = 0x20;
let vbscript: { readonly module: bigint | null; readonly LoadStringW: ctypes.TypedFFIFunction<any, any> } | null = null;

/**
 * Text of a Visual Basic run-time error (`FACILITY_CONTROL` HRESULTs, e.g.
 * `0x800A01C9` from the Scripting runtime), from the string resources of
 * vbscript.dll when installed; `null` otherwise.
 */
function visualBasicMessage(code: number): string | null {
  try {
    if (!vbscript) {
      const LoadLibraryExW = new ctypes.WinDLL("kernel32.dll").func("LoadLibraryExW", ctypes.c_void_p, [ctypes.c_wchar_p, ctypes.c_void_p, ctypes.c_uint32]);
      vbscript = {
        module: LoadLibraryExW("vbscript.dll", null, LOAD_LIBRARY_AS_DATAFILE | LOAD_LIBRARY_AS_IMAGE_RESOURCE),
        LoadStringW: new ctypes.WinDLL("user32.dll").func("LoadStringW", ctypes.c_int32, [ctypes.c_void_p, ctypes.c_uint32, ctypes.c_void_p, ctypes.c_int32]),
      };
    }
    if (!vbscript.module) return null;
    const text = Buffer.alloc(1024);
    const length = vbscript.LoadStringW(vbscript.module, code, text, text.length / 2) as number;
    return length > 0
      ? text
          .subarray(0, length * 2)
          .toString("utf16le")
          .trim()
      : null;
  } catch {
    return null;
  }
}

/** System message for an HRESULT, or a generic text. */
export function hresultText(hr: number): string {
  return systemMessage(hr) ?? `Unknown error ${formatHRESULT(hr)}`;
}

/** A failed HRESULT. Python equivalent: `comtypes.COMError`. */
export class COMError extends Error {
  /** Signed 32-bit HRESULT. */
  readonly hresult: number;
  readonly text: string;
  readonly details: COMErrorDetails | null;
  readonly method: string | null;

  constructor(hresult: number | bigint, options: COMErrorOptions = {}) {
    const hr = toHRESULT(hresult);
    const text = options.text ?? hresultText(hr);
    const details = options.details ?? null;
    const description = details?.description ? ` - ${details.description}` : "";
    const argument = details?.argumentIndex !== undefined ? ` [argument ${details.argumentIndex}]` : "";
    const scode = details?.scode ? ` [scode ${formatHRESULT(details.scode)}]` : "";
    super(`${options.method ? `${options.method}: ` : ""}${text} (${formatHRESULT(hr)})${scode}${argument}${description}`, options.cause === undefined ? undefined : { cause: options.cause });
    this.name = "COMError";
    this.hresult = hr;
    this.text = text;
    this.details = details;
    this.method = options.method ?? null;
  }

  /** Unsigned HRESULT, as `winerror` on node-ctypes errors. */
  get winerror(): number {
    return this.hresult >>> 0;
  }
}

/** Throw a {@link COMError} when `hr` is a failure code; return it otherwise. */
export function checkHRESULT(hr: number | bigint, method?: string): number {
  const value = toHRESULT(hr);
  if (value < 0) throw new COMError(value, { method });
  return value;
}
