import * as ctypes from "node-ctypes";
import { COMError, systemMessage } from "../errors.ts";
import type { ResultLocation } from "../executor/apartment.ts";
import { GUID_NULL } from "../guid.ts";
import { DISP_E_EXCEPTION, DISP_E_PARAMNOTFOUND, DISP_E_TYPEMISMATCH, DISP_E_UNKNOWNNAME } from "../hresult.ts";
import type { InterfacePointer } from "../interface.ts";
import type { IDispatch } from "../interfaces/dispatch.ts";
import { POINTER_SIZE, takeBSTR } from "../low-level/native.ts";
import { DISPATCH_PROPERTYPUT, DISPATCH_PROPERTYPUTREF, DISPID_PROPERTYPUT, LOCALE_NEUTRAL } from "./constants.ts";
import { DISPPARAMS, EXCEPINFO } from "./types.ts";
import { ByRef, VARIANT, VARIANT_SIZE, byRefArgument, clearVariant, missing, readVariant, writeVariant } from "./variant.ts";
import { invokeSlotAsync, sync } from "../pointer.ts";
import { linkOperation } from "../executor/operations.ts";

export type DispatchPointer = InterfacePointer<typeof IDispatch>;

/** Locale passed to `GetIDsOfNames` and `Invoke`; 0 as Python comtypes. */
export let defaultLCID = LOCALE_NEUTRAL;

export function setDefaultLCID(lcid: number): void {
  defaultLCID = lcid >>> 0;
}

function wideStrings(values: readonly string[]): { pointers: Buffer; keep: Buffer[] } {
  const keep = values.map((value) => Buffer.from(`${value}\0`, "utf16le"));
  const pointers = Buffer.alloc(POINTER_SIZE * values.length);
  keep.forEach((buffer, index) => {
    const address = ctypes.addressof(buffer);
    if (POINTER_SIZE === 8) pointers.writeBigUInt64LE(address, index * 8);
    else pointers.writeUInt32LE(Number(address), index * 4);
  });
  return { pointers, keep };
}

/** `IDispatch::GetIDsOfNames`: DISPIDs of a member and, optionally, its named parameters. */
export function getIDsOfNames(dispatch: DispatchPointer, names: readonly string[], lcid = defaultLCID): number[] {
  if (!names.length) return [];
  const { pointers, keep } = wideStrings(names);
  const dispids = Buffer.alloc(4 * names.length);
  try {
    sync(dispatch).GetIDsOfNames(GUID_NULL, pointers, names.length, lcid, dispids);
  } catch (error) {
    if (error instanceof COMError && error.hresult === DISP_E_UNKNOWNNAME) {
      throw new COMError(DISP_E_UNKNOWNNAME, { method: `IDispatch::GetIDsOfNames(${names.join(", ")})`, text: `Unknown name ${JSON.stringify(names[0])}${names.length > 1 ? " or one of its named arguments" : ""}`, cause: error });
    }
    throw error;
  } finally {
    void keep;
  }
  return Array.from({ length: names.length }, (_, index) => dispids.readInt32LE(index * 4));
}

export interface InvokeOptions {
  /** Named arguments: DISPIDs (from `getIDsOfNames`) and values. */
  readonly named?: { readonly dispids: readonly number[]; readonly values: readonly unknown[] };
  readonly lcid?: number;
  /** Member name, for error messages. */
  readonly name?: string;
}

const DeferredFillIn = ctypes.WINFUNCTYPE(ctypes.c_int32, ctypes.c_void_p);

function exceptionError(info: EXCEPINFO, method: string, cause: unknown): COMError {
  if (info.pfnDeferredFillIn) DeferredFillIn(info.pfnDeferredFillIn)(info._buffer);
  const source = takeBSTR(info.bstrSource);
  // Some servers (e.g. the Scripting runtime) set only the scode.
  const description = takeBSTR(info.bstrDescription) || (info.scode ? systemMessage(info.scode) : null);
  const helpFile = takeBSTR(info.bstrHelpFile);
  return new COMError(DISP_E_EXCEPTION, { method, details: { description, source, helpFile, helpContext: info.dwHelpContext, guid: null, scode: info.scode }, cause });
}

interface PreparedInvoke {
  /** Native arguments of `IDispatch::Invoke` (after `this`). */
  readonly args: unknown[];
  /** Memory referenced by address from `DISPPARAMS`: alive as long as this object. */
  readonly keep: readonly Buffer[];
  /** The JavaScript argument values, for the executor's transfer check. */
  readonly values: readonly unknown[];
  /** Receives the result (not passed for property puts). */
  readonly result: VARIANT;
  readonly put: boolean;
  /** The error to throw for a failed call. */
  translate(error: unknown): unknown;
  /** Reads back `ByRef` arguments and frees the argument and result VARIANTs. */
  dispose(): void;
}

// Converts the arguments; on failure, frees what was already converted.
function prepareInvoke(dispid: number, flags: number, args: readonly unknown[], options: InvokeOptions): PreparedInvoke {
  const put = (flags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) !== 0;
  if (put && !args.length) throw new TypeError("A property put requires a value");
  const positional = put ? args.slice(0, -1) : [...args];
  while (positional.length && positional[positional.length - 1] === undefined) positional.pop();
  const namedDispids = [...(options.named?.dispids ?? [])];
  const namedValues = [...(options.named?.values ?? [])];
  if (namedDispids.length !== namedValues.length) throw new TypeError("Named argument DISPIDs and values differ in length");
  if (put) {
    namedDispids.unshift(DISPID_PROPERTYPUT);
    namedValues.unshift(args[args.length - 1]);
  }
  // rgvarg: named arguments first, then positional arguments in reverse order.
  const values = [...namedValues, ...positional.reverse()];
  const count = values.length;
  const variants = Buffer.alloc(VARIANT_SIZE * count);
  const dispids = Buffer.alloc(4 * namedDispids.length);
  namedDispids.forEach((id, index) => dispids.writeInt32LE(id, index * 4));
  const method = `IDispatch::Invoke(${options.name ?? `DISPID ${dispid}`})`;
  const result = new VARIANT();
  const info = new EXCEPINFO();
  const argumentError = new ctypes.c_uint32();
  const references: { finish(): void }[] = [];
  const dispose = () => {
    // ByRef arguments: read back what the callee left, then release it.
    for (const reference of references) reference.finish();
    for (let index = 0; index < count; index++) clearVariant(variants.subarray(index * VARIANT_SIZE, (index + 1) * VARIANT_SIZE));
    result.clear();
  };
  try {
    values.forEach((value, index) => {
      const slot = variants.subarray(index * VARIANT_SIZE, (index + 1) * VARIANT_SIZE);
      try {
        if (value instanceof ByRef) references.push(byRefArgument(slot, value));
        else writeVariant(slot, index >= namedValues.length && value === undefined ? missing : value);
      } catch (error) {
        const position = index < namedValues.length ? index : count - 1 - index;
        throw new TypeError(`${method}: cannot convert argument ${position}: ${(error as Error).message}`, { cause: error });
      }
    });
  } catch (error) {
    dispose();
    throw error;
  }
  const params = new DISPPARAMS();
  params.rgvarg = count ? ctypes.addressof(variants) : null;
  params.rgdispidNamedArgs = namedDispids.length ? ctypes.addressof(dispids) : null;
  params.cArgs = count;
  params.cNamedArgs = namedDispids.length;
  return {
    args: [dispid, GUID_NULL, options.lcid ?? defaultLCID, flags, params, put ? null : result, info, ctypes.byref(argumentError)],
    keep: [variants, dispids],
    values,
    result,
    put,
    translate(error) {
      if (!(error instanceof COMError)) return error;
      if (error.hresult === DISP_E_EXCEPTION) return exceptionError(info, method, error);
      if (error.hresult === DISP_E_TYPEMISMATCH || error.hresult === DISP_E_PARAMNOTFOUND) {
        const index = argumentError.value;
        const argumentIndex = index < namedValues.length ? index : count - 1 - index;
        return new COMError(error.hresult, { method, details: { description: null, source: null, helpFile: null, helpContext: 0, guid: null, argumentIndex }, cause: error });
      }
      return new COMError(error.hresult, { method, details: error.details, cause: error });
    },
    dispose,
  };
}

/**
 * `IDispatch::Invoke`. Positional arguments are converted to VARIANTs
 * (`undefined` is an omitted optional argument; `byRef(value)` is passed by
 * reference and updated after the call); for property puts the last
 * positional argument is the value. Returns the converted result; COM objects
 * are new owned pointers.
 */
export function invoke(dispatch: DispatchPointer, dispid: number, flags: number, args: readonly unknown[] = [], options: InvokeOptions = {}): unknown {
  const prepared = prepareInvoke(dispid, flags, args, options);
  try {
    try {
      (sync(dispatch).Invoke as (...args: unknown[]) => unknown)(...prepared.args);
    } catch (error) {
      throw prepared.translate(error);
    }
    return readVariant(prepared.result._buffer);
  } finally {
    prepared.dispose();
  }
}

/**
 * {@link invoke} on the COM executor: the VARIANTs are built and read on the
 * JavaScript thread, `Invoke` runs on the executor (or on the home thread
 * when the object or an argument must stay there).
 */
export function invokeAsync(dispatch: DispatchPointer, dispid: number, flags: number, args: readonly unknown[] = [], options: InvokeOptions = {}): Promise<unknown> {
  let prepared: PreparedInvoke;
  try {
    prepared = prepareInvoke(dispid, flags, args, options);
  } catch (error) {
    return Promise.reject(error);
  }
  const slot = dispatch.interface.methodIndex.get("Invoke")!;
  const results: ResultLocation[] = prepared.put ? [] : [{ location: ctypes.addressof(prepared.result._buffer), kind: "variant" }];
  const call = invokeSlotAsync(dispatch, slot, prepared.args, { inspect: prepared.values, results });
  const converted = call
    .then(
      () => readVariant(prepared.result._buffer),
      (error: unknown) => {
        throw prepared.translate(error);
      },
    )
    .finally(() => prepared.dispose());
  return linkOperation(converted, call);
}
