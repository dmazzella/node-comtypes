import * as ctypes from "node-ctypes";
import { VARIANT, borrowReference, toCurrency, clearVariant, createSafeArray, destroySafeArray, ErrorValue, readSafeArray, readVariant, Typed, VARIANT_SIZE, writeVariant } from "../automation/variant.ts";
import { VT_TYPEMASK } from "../automation/constants.ts";
import { resolveParamInterface } from "../commethod.ts";
import type { ParamKind } from "../commethod.ts";
import { currencyToString } from "../automation/decimal.ts";
import { dateToOADate, oaDateToDate } from "../automation/date.ts";
import { checkHRESULT } from "../errors.ts";
import { GUID, toGUID } from "../guid.ts";
import type { GUIDLike } from "../guid.ts";
import { DISP_E_PARAMNOTFOUND } from "../hresult.ts";
import { derivesFrom } from "../interface.ts";
import type { InterfaceDef } from "../interface.ts";
import { allocBSTR, ole32, oleaut32, readBSTR, readMemory, writeMemory } from "../low-level/native.ts";
import { rawAddRef, rawQueryInterface, rawRelease } from "../low-level/vtable.ts";
import { COMPointer } from "../pointer.ts";

// Server side of the parameter conversions: what a JavaScript implementation
// receives from a native caller, and how its results are written back.

function readPointer(address: bigint): bigint | null {
  return ctypes.readValue(address, ctypes.c_void_p);
}

function writePointer(address: bigint, value: bigint | null): void {
  ctypes.writeValue(address, ctypes.c_void_p, value ?? 0n);
}

/** Callback argument type: pointers arrive as raw addresses. */
export function callbackArgtype(type: ctypes.FFIType): ctypes.FFIType {
  return typeof type === "object" && type !== null && "_type_" in type && "_pointerTo" in type ? ctypes.c_void_p : type;
}

/** `VT_ERROR`/`DISP_E_PARAMNOTFOUND` (an omitted optional argument) becomes `undefined`. */
export function omittedToUndefined(value: unknown): unknown {
  return value instanceof ErrorValue && value.scode === DISP_E_PARAMNOTFOUND ? undefined : value;
}

/** Wrap an incoming interface pointer: the implementation receives its own reference. */
function incomingInterface(address: bigint | null, iface: InterfaceDef<any>): COMPointer | null {
  return address ? COMPointer.fromAddress(address, iface, { addRef: true }) : null;
}

/** Reference to store in an out parameter, owned by the caller. */
export function outgoingReference(value: unknown, iface: InterfaceDef<any>): bigint {
  if (value === null || value === undefined) return 0n;
  const borrowed = borrowReference(value);
  if (!borrowed) throw new TypeError(`Expected a COM object for ${iface.name}`);
  if (derivesFrom(borrowed.interface, iface.iid)) {
    rawAddRef(borrowed.address);
    return borrowed.address;
  }
  const [hr, address] = rawQueryInterface(borrowed.address, iface.iid);
  checkHRESULT(hr, `QueryInterface(${iface.name})`);
  return address ?? 0n;
}

/** Value of an `in` parameter, from the native callback argument. */
export function decodeInput(kind: ParamKind, native: unknown, self: InterfaceDef<any>): unknown {
  switch (kind.kind) {
    case "bstr":
      return readBSTR(native as bigint | null) ?? "";
    case "lpolestr":
      return native ? ctypes.wstring_at(native as bigint) : null;
    case "refguid":
      return native ? new GUID(new Uint8Array(readMemory(native as bigint, 16))) : null;
    case "safearray":
      return native ? readSafeArray(native as bigint, kind.vt) : null;
    case "interface":
      return incomingInterface(native as bigint | null, resolveParamInterface(kind, self));
    case "date":
      return oaDateToDate(native as number);
    case "currency":
      return Number(currencyToString(BigInt(native as bigint)));
    case "variant":
      return omittedToUndefined(readVariant((native as VARIANT)._buffer));
    case "struct":
      return (kind.type as unknown as typeof ctypes.Structure).from_buffer_copy((native as ctypes.Structure)._buffer) as ctypes.Structure;
    case "pointer":
      // VARIANT* is a value; other T* stay raw addresses (often arrays, size_is).
      return kind.base.kind === "variant" ? (native ? decodeIndirect(kind.base, native as bigint, self) : undefined) : native;
    default:
      return native;
  }
}

/** Current value of an `[in, out]` parameter at `address`. */
export function decodeIndirect(kind: ParamKind, address: bigint, self: InterfaceDef<any>): unknown {
  switch (kind.kind) {
    case "bstr":
      return readBSTR(readPointer(address)) ?? "";
    case "lpolestr": {
      const text = readPointer(address);
      return text ? ctypes.wstring_at(text) : null;
    }
    case "safearray":
      return readSafeArray(readPointer(address) ?? 0n, kind.vt);
    case "interface":
      return incomingInterface(readPointer(address), resolveParamInterface(kind, self));
    case "date":
      return oaDateToDate(ctypes.readValue(address, ctypes.c_double));
    case "currency":
      return Number(currencyToString(BigInt(ctypes.readValue(address, ctypes.c_int64))));
    case "variant":
      return omittedToUndefined(readVariant(readMemory(address, VARIANT_SIZE)));
    case "refguid":
      return new GUID(new Uint8Array(readMemory(address, 16)));
    case "struct":
      return (kind.type as unknown as typeof ctypes.Structure).from_buffer_copy(readMemory(address, ctypes.sizeof(kind.type as ctypes.StructureClass))) as ctypes.Structure;
    case "simple":
      return ctypes.readValue(address, kind.type);
    case "pointer":
      return decodeIndirect(kind.base, address, self);
    default:
      throw new TypeError("Unsupported [in, out] parameter type");
  }
}

/** Reset a pure `out` parameter, so that a failure leaves NULL / VT_EMPTY. */
export function resetOutput(kind: ParamKind, address: bigint): void {
  switch (kind.kind) {
    case "bstr":
    case "lpolestr":
    case "safearray":
    case "interface":
      writePointer(address, null);
      return;
    case "variant":
      writeMemory(address, Buffer.alloc(VARIANT_SIZE));
      return;
    case "pointer":
      resetOutput(kind.base, address);
      return;
  }
}

function structureBytes(type: new (...args: any[]) => ctypes.Structure, value: unknown): Buffer {
  if (value instanceof type) return value._buffer;
  if ((type === GUID || type.prototype instanceof GUID) && (typeof value === "string" || value instanceof GUID)) return toGUID(value as GUIDLike)._buffer;
  return (type as unknown as typeof ctypes.Structure).create(value as Record<string, unknown>)._buffer;
}

/** Write an `out` value at `address`; `inout` releases the previous value first. */
export function encodeOutput(kind: ParamKind, address: bigint, value: unknown, inout: boolean, self: InterfaceDef<any>): void {
  switch (kind.kind) {
    case "bstr": {
      const previous = inout ? readPointer(address) : null;
      writePointer(address, value === null || value === undefined ? null : allocBSTR(String(value)));
      if (previous) oleaut32().SysFreeString(previous);
      return;
    }
    case "lpolestr": {
      const previous = inout ? readPointer(address) : null;
      let text: bigint | null = null;
      if (value !== null && value !== undefined) {
        const bytes = Buffer.from(`${String(value)}\0`, "utf16le");
        text = ole32().CoTaskMemAlloc(bytes.length);
        if (!text) throw new RangeError("CoTaskMemAlloc failed");
        writeMemory(text, bytes);
      }
      writePointer(address, text);
      if (previous) ole32().CoTaskMemFree(previous);
      return;
    }
    case "safearray": {
      const previous = inout ? readPointer(address) : null;
      const vt = kind.vt;
      writePointer(address, value === null || value === undefined ? null : value instanceof Typed ? createSafeArray(value.vt & VT_TYPEMASK, value.value, value.dimensions) : createSafeArray(vt, value, 1));
      destroySafeArray(previous);
      return;
    }
    case "interface": {
      const previous = inout ? readPointer(address) : null;
      writePointer(address, outgoingReference(value, resolveParamInterface(kind, self)));
      if (previous) rawRelease(previous);
      return;
    }
    case "variant": {
      const view = VARIANT.from_address(address);
      if (inout) clearVariant(view._buffer);
      else view._buffer.fill(0);
      writeVariant(view._buffer, value);
      return;
    }
    case "refguid":
      writeMemory(address, toGUID(value as GUIDLike)._buffer);
      return;
    case "date":
      ctypes.writeValue(address, ctypes.c_double, value instanceof Date ? dateToOADate(value) : Number(value ?? 0));
      return;
    case "currency":
      ctypes.writeValue(address, ctypes.c_int64, toCurrency(value ?? 0));
      return;
    case "struct":
      writeMemory(address, structureBytes(kind.type, value));
      return;
    case "simple":
      ctypes.writeValue(address, kind.type, value as never);
      return;
    case "pointer":
      encodeOutput(kind.base, address, value, inout, self);
      return;
    default:
      throw new TypeError("Unsupported out parameter type");
  }
}
