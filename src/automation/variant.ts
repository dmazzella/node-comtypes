import * as ctypes from "node-ctypes";
import { checkHRESULT } from "../errors.ts";
import { GUID } from "../guid.ts";
import { DISP_E_PARAMNOTFOUND } from "../hresult.ts";
import { derivesFrom, findInterface } from "../interface.ts";
import type { InterfaceDef } from "../interface.ts";
import { allocBSTR, oleaut32, POINTER_SIZE, readBSTR, readMemory, writeMemory } from "../low-level/native.ts";
import { rawAddRef, rawQueryInterface, rawRelease } from "../low-level/vtable.ts";
import { RecordInfo, recordParts } from "./record.ts";
import { COMPointer } from "../pointer.ts";
import {
  VT_ARRAY,
  VT_BOOL,
  VT_BSTR,
  VT_BYREF,
  VT_CY,
  VT_DATE,
  VT_DECIMAL,
  VT_DISPATCH,
  VT_EMPTY,
  VT_ERROR,
  VT_I1,
  VT_I2,
  VT_I4,
  VT_I8,
  VT_INT,
  VT_INT_PTR,
  VT_NULL,
  VT_R4,
  VT_R8,
  VT_RECORD,
  VT_TYPEMASK,
  VT_UI1,
  VT_UI2,
  VT_UI4,
  VT_UI8,
  VT_UINT,
  VT_UINT_PTR,
  VT_UNKNOWN,
  VT_VARIANT,
  vartypeName,
} from "./constants.ts";
import { dateToOADate, oaDateToDate } from "./date.ts";
import { currencyToString, readDecimal, writeDecimal } from "./decimal.ts";

/** Size of a VARIANT: 16 bytes on 32-bit Windows, 24 on 64-bit. */
export const VARIANT_SIZE = 8 + 2 * POINTER_SIZE;
const DATA = 8;

const IID_IUnknown = new GUID("{00000000-0000-0000-C000-000000000046}");
const IID_IDispatch = new GUID("{00020400-0000-0000-C000-000000000046}");

/**
 * Objects exposing a COM pointer under this symbol (e.g. dynamic dispatch
 * proxies) are converted like that pointer.
 */
export const comPointer: unique symbol = Symbol.for("node-comtypes.pointer") as never;

/** A `VT_ERROR` value (SCODE). */
export class ErrorValue {
  readonly scode: number;
  constructor(scode: number) {
    this.scode = scode | 0;
  }
  [Symbol.for("nodejs.util.inspect.custom")](): string {
    return `ErrorValue(0x${(this.scode >>> 0).toString(16).toUpperCase()})`;
  }
}

/** Omitted optional argument: `VT_ERROR` with `DISP_E_PARAMNOTFOUND`. Python: `comtypes.automation.Missing`. */
export const missing: ErrorValue = Object.freeze(new ErrorValue(DISP_E_PARAMNOTFOUND));

export interface TypedOptions {
  /** For `VT_ARRAY` types: number of dimensions of the nested arrays (default 1). */
  readonly dimensions?: number;
}

/** A value with an explicit VARTYPE, e.g. `typed(VT_UI1, 7)` or `typed(VT_ARRAY | VT_BSTR, ["a"])`. */
export class Typed {
  readonly vt: number;
  readonly value: unknown;
  readonly dimensions: number;
  constructor(vt: number, value: unknown, dimensions = 1) {
    this.vt = vt;
    this.value = value;
    this.dimensions = dimensions;
  }
}

/**
 * An argument passed by reference to `IDispatch::Invoke` (VB `ByRef`): the
 * callee may change it, and `value` holds the result after the call.
 * Python/win32com equivalent: the returned byref values.
 */
export class ByRef<T = unknown> {
  value: T;
  /** VARTYPE of the referenced value; `null`: a `VARIANT` (`VT_BYREF | VT_VARIANT`). */
  readonly vt: number | null;
  constructor(value: T, vt: number | null = null) {
    this.value = value;
    this.vt = vt;
  }
}

/** `byRef(false)`, or with an explicit type `byRef("x", VT_BSTR)` (`VT_BYREF | VT_BSTR`). */
export function byRef<T>(value: T, vt: number | null = null): ByRef<T> {
  if (vt !== null && (!Number.isInteger(vt) || vt <= 0 || vt & (VT_BYREF | VT_ARRAY) || vt === VT_VARIANT)) throw new RangeError(`Unsupported ByRef VARTYPE ${vt}`);
  return new ByRef(value, vt);
}

/** Holder of a `ByRef` argument: `slot` becomes `VT_BYREF` to it; call `finish` after the call. */
export function byRefArgument(slot: Buffer, reference: ByRef): { readonly holder: Buffer; finish(): void } {
  const holder = Buffer.alloc(VARIANT_SIZE);
  writeVariant(holder, reference.vt === null ? reference.value : typed(reference.vt, reference.value));
  const vt = reference.vt ?? VT_VARIANT;
  // VT_BYREF | VT_VARIANT and VT_BYREF | VT_DECIMAL point to the whole holder; other types to its data.
  writePointer(slot, DATA, ctypes.addressof(holder) + (vt === VT_VARIANT || vt === VT_DECIMAL ? 0n : BigInt(DATA)));
  slot.writeUInt16LE(VT_BYREF | vt, 0);
  return {
    holder,
    finish() {
      try {
        reference.value = readVariant(holder) as never;
      } finally {
        clearVariant(holder);
      }
    },
  };
}

/** Replace the value referenced by a `VT_BYREF | vt` argument (server side), releasing the previous one. */
export function writeReference(address: bigint, vt: number, value: unknown): void {
  if (vt === VT_VARIANT || vt === VT_DECIMAL) {
    const view = VARIANT.from_address(address);
    if (vt === VT_VARIANT) {
      clearVariant(view._buffer);
      writeVariant(view._buffer, value);
    } else {
      writeDecimal(view._buffer, 0, value as string | number | bigint);
      view._buffer.writeUInt16LE(0, 0);
    }
    return;
  }
  const temporary = Buffer.alloc(VARIANT_SIZE);
  writeVariant(temporary, vt & VT_ARRAY ? safeArray(value as ArrayLike<unknown>, { vt: vt & VT_TYPEMASK }) : typed(vt, value));
  const size = vt & VT_ARRAY ? POINTER_SIZE : elementSize(vt);
  // Release what the reference held: BSTR, interface or array.
  if (needsClear(vt)) {
    const previous = Buffer.alloc(VARIANT_SIZE);
    previous.writeUInt16LE(vt, 0);
    writePointer(previous, DATA, readPointer(readMemory(address, POINTER_SIZE), 0));
    clearVariant(previous);
  }
  writeMemory(address, temporary.subarray(DATA, DATA + size));
}

export function typed(vt: number, value: unknown, options: TypedOptions = {}): Typed {
  if (!Number.isInteger(vt) || vt < 0 || vt > 0xffff || vt & VT_BYREF) throw new RangeError(`Unsupported VARTYPE ${vt}`);
  const dimensions = options.dimensions ?? 1;
  if (!Number.isInteger(dimensions) || dimensions < 1) throw new RangeError("dimensions must be a positive integer");
  return new Typed(vt, value, dimensions);
}

/**
 * A SAFEARRAY argument. Nested arrays become an N-dimensional array when
 * `dimensions` > 1 (`values[i][j]` is element `(i, j)`), otherwise arrays of arrays.
 */
export function safeArray(values: ArrayLike<unknown>, options: TypedOptions & { readonly vt?: number } = {}): Typed {
  return typed(VT_ARRAY | (options.vt ?? VT_VARIANT), values, options);
}

// ── memory helpers ────────────────────────────────────────────────────

function readPointer(buffer: Buffer, offset: number): bigint {
  return POINTER_SIZE === 8 ? buffer.readBigUInt64LE(offset) : BigInt(buffer.readUInt32LE(offset));
}

function writePointer(buffer: Buffer, offset: number, value: bigint): void {
  if (POINTER_SIZE === 8) buffer.writeBigUInt64LE(value, offset);
  else buffer.writeUInt32LE(Number(value), offset);
}

/** Size of one element of VARTYPE `vt` (SAFEARRAY element or BYREF target). */
export function elementSize(vt: number): number {
  switch (vt) {
    case VT_I1:
    case VT_UI1:
      return 1;
    case VT_I2:
    case VT_UI2:
    case VT_BOOL:
      return 2;
    case VT_I4:
    case VT_UI4:
    case VT_INT:
    case VT_UINT:
    case VT_R4:
    case VT_ERROR:
      return 4;
    case VT_R8:
    case VT_CY:
    case VT_DATE:
    case VT_I8:
    case VT_UI8:
      return 8;
    case VT_BSTR:
    case VT_DISPATCH:
    case VT_UNKNOWN:
    case VT_INT_PTR:
    case VT_UINT_PTR:
      return POINTER_SIZE;
    case VT_DECIMAL:
      return 16;
    case VT_VARIANT:
      return VARIANT_SIZE;
    default:
      throw new TypeError(`Unsupported element type ${vartypeName(vt)}`);
  }
}

/** VARTYPEs that own resources released by `VariantClear`. */
function needsClear(vt: number): boolean {
  if (vt & VT_BYREF) return false;
  if (vt & VT_ARRAY) return true;
  return vt === VT_BSTR || vt === VT_DISPATCH || vt === VT_UNKNOWN || vt === VT_RECORD;
}

// ── reading ───────────────────────────────────────────────────────────

function wrapInterface(address: bigint, iid: GUID): COMPointer | null {
  if (!address) return null;
  const iface: InterfaceDef<any> = findInterface(iid) ?? (findInterface(IID_IUnknown) as InterfaceDef<any>);
  rawAddRef(address);
  return COMPointer.fromAddress(address, iface);
}

/** Decode a value of VARTYPE `vt` stored at `offset` (not a whole VARIANT). */
function readScalar(vt: number, buffer: Buffer, offset: number): unknown {
  switch (vt) {
    case VT_EMPTY:
      return undefined;
    case VT_NULL:
      return null;
    case VT_I1:
      return buffer.readInt8(offset);
    case VT_UI1:
      return buffer.readUInt8(offset);
    case VT_I2:
      return buffer.readInt16LE(offset);
    case VT_UI2:
      return buffer.readUInt16LE(offset);
    case VT_I4:
    case VT_INT:
      return buffer.readInt32LE(offset);
    case VT_UI4:
    case VT_UINT:
      return buffer.readUInt32LE(offset);
    case VT_ERROR:
      return new ErrorValue(buffer.readInt32LE(offset));
    case VT_R4:
      return buffer.readFloatLE(offset);
    case VT_R8:
      return buffer.readDoubleLE(offset);
    case VT_BOOL:
      return buffer.readInt16LE(offset) !== 0;
    case VT_CY:
      return Number(currencyToString(buffer.readBigInt64LE(offset)));
    case VT_DATE:
      return oaDateToDate(buffer.readDoubleLE(offset));
    case VT_BSTR:
      return readBSTR(readPointer(buffer, offset)) ?? "";
    case VT_DISPATCH:
      return wrapInterface(readPointer(buffer, offset), IID_IDispatch);
    case VT_UNKNOWN:
      return wrapInterface(readPointer(buffer, offset), IID_IUnknown);
    case VT_I8:
      return buffer.readBigInt64LE(offset);
    case VT_UI8:
      return buffer.readBigUInt64LE(offset);
    case VT_INT_PTR:
    case VT_UINT_PTR:
      return readPointer(buffer, offset);
    case VT_DECIMAL:
      return Number(readDecimal(buffer, offset));
    case VT_VARIANT:
      return readVariant(buffer.subarray(offset, offset + VARIANT_SIZE));
    default:
      throw new TypeError(`Unsupported VARIANT type ${vartypeName(vt)}`);
  }
}

/** Convert the VARIANT in `buffer` to a JavaScript value. COM objects are returned as new owned pointers. */
export function readVariant(buffer: Buffer): unknown {
  const vt = buffer.readUInt16LE(0);
  // VT_RECORD (also by reference): pvRecord and pRecInfo.
  if ((vt & ~VT_BYREF) === VT_RECORD) {
    const record = readPointer(buffer, DATA);
    const info = readPointer(buffer, DATA + POINTER_SIZE);
    return record && info ? RecordInfo.borrow(info).read(record) : null;
  }
  if (vt & VT_BYREF) {
    const target = readPointer(buffer, DATA);
    const base = vt & ~VT_BYREF;
    if (!target) return null;
    if (base & VT_ARRAY) return readSafeArray(readPointer(readMemory(target, POINTER_SIZE), 0), base & VT_TYPEMASK);
    return readScalar(base, readMemory(target, elementSize(base)), 0);
  }
  if (vt & VT_ARRAY) return readSafeArray(readPointer(buffer, DATA), vt & VT_TYPEMASK);
  if (vt === VT_DECIMAL) return Number(readDecimal(buffer, 0));
  if (vt === VT_VARIANT) throw new TypeError("VT_VARIANT requires VT_BYREF");
  return readScalar(vt, buffer, DATA);
}

/** Exact text of a `VT_DECIMAL` or `VT_CY` VARIANT (the `value` getter returns a number). */
export function readVariantExact(buffer: Buffer): string {
  const vt = buffer.readUInt16LE(0);
  if (vt === VT_DECIMAL) return readDecimal(buffer, 0);
  if (vt === VT_CY) return currencyToString(buffer.readBigInt64LE(DATA));
  throw new TypeError(`readVariantExact expects VT_DECIMAL or VT_CY, not ${vartypeName(vt)}`);
}

/** Convert a SAFEARRAY to nested JavaScript arrays (1-D: a flat array). */
export function readSafeArray(psa: bigint, vtHint = VT_VARIANT): unknown[] | null {
  if (!psa) return null;
  const api = oleaut32();
  const dimensions = api.SafeArrayGetDim(psa);
  const vtOut = new ctypes.c_uint16();
  const vt = api.SafeArrayGetVartype(psa, ctypes.byref(vtOut)) >= 0 && vtOut.value ? vtOut.value : vtHint;
  const size = api.SafeArrayGetElemsize(psa);
  const lengths: number[] = [];
  for (let dimension = 1; dimension <= dimensions; dimension++) {
    const lower = new ctypes.c_int32();
    const upper = new ctypes.c_int32();
    checkHRESULT(api.SafeArrayGetLBound(psa, dimension, ctypes.byref(lower)), "SafeArrayGetLBound");
    checkHRESULT(api.SafeArrayGetUBound(psa, dimension, ctypes.byref(upper)), "SafeArrayGetUBound");
    lengths.push(Math.max(0, upper.value - lower.value + 1));
  }
  const count = lengths.reduce((total, length) => total * length, 1);
  const data = new ctypes.c_void_p();
  checkHRESULT(api.SafeArrayAccessData(psa, ctypes.byref(data)), "SafeArrayAccessData");
  let bytes: Buffer;
  let records: unknown[] | null = null;
  try {
    if (vt === VT_RECORD) {
      // Records are read in place, through the array's IRecordInfo.
      const info = new ctypes.c_void_p();
      checkHRESULT(api.SafeArrayGetRecordInfo(psa, ctypes.byref(info)), "SafeArrayGetRecordInfo");
      const recordInfo = RecordInfo.borrow(info.value!);
      rawRelease(info.value!);
      records = Array.from({ length: count }, (_, index) => recordInfo.read((data.value ?? 0n) + BigInt(index * size)));
    }
    bytes = records ? Buffer.alloc(0) : readMemory(data.value ?? 0n, count * size);
  } finally {
    api.SafeArrayUnaccessData(psa);
  }
  const element = (index: number) => (records ? records[index] : readScalar(vt, bytes, index * size));
  if (dimensions <= 1) return Array.from({ length: lengths[0] ?? 0 }, (_, index) => element(index));
  // Column-major storage: the first index varies fastest.
  const build = (dimension: number, base: number, stride: number): unknown[] => {
    const length = lengths[dimension];
    if (dimension === dimensions - 1) return Array.from({ length }, (_, index) => element(base + index * stride));
    return Array.from({ length }, (_, index) => build(dimension + 1, base + index * stride, stride * length));
  };
  return build(0, 0, 1);
}

// ── writing ───────────────────────────────────────────────────────────

/**
 * Objects implemented in JavaScript (`COMObject`) lend an interface pointer
 * through a method under this symbol, without adding a reference.
 */
export const comReference: unique symbol = Symbol.for("node-comtypes.reference") as never;

/** An interface pointer lent by its owner: valid while the owner holds it, not owned by the caller. */
export interface BorrowedReference {
  readonly address: bigint;
  readonly interface: InterfaceDef<any>;
}

/** The interface pointer behind a COM pointer, a dynamic object or a JavaScript COM object, if any. */
export function borrowReference(value: unknown): BorrowedReference | null {
  if (value instanceof COMPointer) return { address: value.address, interface: value.interface };
  if (value !== null && (typeof value === "object" || typeof value === "function")) {
    const record = value as Record<symbol, unknown>;
    const pointer = record[comPointer];
    if (pointer instanceof COMPointer) return { address: pointer.address, interface: pointer.interface };
    const lend = record[comReference];
    if (typeof lend === "function") return (lend as () => BorrowedReference | null).call(value);
  }
  return null;
}

function typedArrayVartype(value: ArrayBufferView): number | null {
  if (value instanceof Uint8Array) return VT_UI1;
  if (value instanceof Int8Array) return VT_I1;
  if (value instanceof Int16Array) return VT_I2;
  if (value instanceof Uint16Array) return VT_UI2;
  if (value instanceof Int32Array) return VT_I4;
  if (value instanceof Uint32Array) return VT_UI4;
  if (value instanceof Float32Array) return VT_R4;
  if (value instanceof Float64Array) return VT_R8;
  if (value instanceof BigInt64Array) return VT_I8;
  if (value instanceof BigUint64Array) return VT_UI8;
  return null;
}

/** VARTYPE chosen for a JavaScript value. */
export function inferVartype(value: unknown): number {
  switch (typeof value) {
    case "undefined":
      return VT_EMPTY;
    case "boolean":
      return VT_BOOL;
    case "number":
      return Number.isInteger(value) && value >= -0x80000000 && value <= 0x7fffffff && !Object.is(value, -0) ? VT_I4 : VT_R8;
    case "bigint":
      return value >= -(2n ** 63n) && value < 2n ** 63n ? VT_I8 : VT_UI8;
    case "string":
      return VT_BSTR;
  }
  if (value === null) return VT_NULL;
  if (value instanceof Typed) return value.vt;
  if (value instanceof ErrorValue) return VT_ERROR;
  if (value instanceof Date) return VT_DATE;
  if (value instanceof VARIANT) return value.vt;
  const reference = borrowReference(value);
  if (reference) return derivesFrom(reference.interface, IID_IDispatch) ? VT_DISPATCH : VT_UNKNOWN;
  if (ArrayBuffer.isView(value)) {
    const vt = typedArrayVartype(value);
    if (vt !== null) return VT_ARRAY | vt;
  }
  if (Array.isArray(value)) return VT_ARRAY | VT_VARIANT;
  if (recordParts(value)) return VT_RECORD;
  throw new TypeError(`Cannot convert ${Object.prototype.toString.call(value)} to a VARIANT`);
}

function toInteger(value: unknown, min: number, max: number, vt: number): number {
  const number = typeof value === "bigint" ? Number(value) : typeof value === "boolean" ? (value ? 1 : 0) : (value as number);
  if (typeof number !== "number" || !Number.isInteger(number) || number < min || number > max) throw new RangeError(`${String(value)} is not a valid ${vartypeName(vt)}`);
  return number;
}

function toBigInt(value: unknown, min: bigint, max: bigint, vt: number): bigint {
  let result: bigint;
  try {
    result = typeof value === "bigint" ? value : BigInt(value as number);
  } catch {
    throw new RangeError(`${String(value)} is not a valid ${vartypeName(vt)}`);
  }
  if (result < min || result > max) throw new RangeError(`${String(value)} is out of range for ${vartypeName(vt)}`);
  return result;
}

/** CY units (value x 10 000) of a number, exact decimal string or bigint (whole units). */
export function toCurrency(value: unknown): bigint {
  if (typeof value === "bigint") return value * 10000n;
  const text = typeof value === "number" ? value.toFixed(4) : String(value).trim();
  const match = /^([+-])?(\d+)(?:\.(\d{0,4})\d*)?$/.exec(text);
  if (!match) throw new RangeError(`${String(value)} is not a valid VT_CY`);
  const units = BigInt(match[2]) * 10000n + BigInt((match[3] ?? "").padEnd(4, "0"));
  return toBigInt(match[1] === "-" ? -units : units, -(2n ** 63n), 2n ** 63n - 1n, VT_CY);
}

/** COM pointer to store for `VT_DISPATCH`/`VT_UNKNOWN`, with one reference owned by the destination. */
function interfaceReference(value: unknown, vt: number): bigint {
  if (value === null || value === undefined) return 0n;
  const reference = borrowReference(value);
  if (!reference) throw new TypeError(`${vartypeName(vt)} requires a COM object`);
  if (vt === VT_DISPATCH && !derivesFrom(reference.interface, IID_IDispatch)) {
    const [hr, dispatch] = rawQueryInterface(reference.address, IID_IDispatch);
    checkHRESULT(hr, "IUnknown::QueryInterface(IDispatch)");
    return dispatch ?? 0n;
  }
  rawAddRef(reference.address);
  return reference.address;
}

/**
 * Encode `value` as VARTYPE `vt` at `offset` (not a whole VARIANT). Resources
 * (BSTR, interface references, SAFEARRAYs) become owned by the destination.
 */
function writeScalar(vt: number, buffer: Buffer, offset: number, value: unknown): void {
  switch (vt) {
    case VT_EMPTY:
    case VT_NULL:
      return;
    case VT_I1:
      buffer.writeInt8(toInteger(value, -0x80, 0x7f, vt), offset);
      return;
    case VT_UI1:
      buffer.writeUInt8(toInteger(value, 0, 0xff, vt), offset);
      return;
    case VT_I2:
      buffer.writeInt16LE(toInteger(value, -0x8000, 0x7fff, vt), offset);
      return;
    case VT_UI2:
      buffer.writeUInt16LE(toInteger(value, 0, 0xffff, vt), offset);
      return;
    case VT_I4:
    case VT_INT:
      buffer.writeInt32LE(toInteger(value, -0x80000000, 0x7fffffff, vt), offset);
      return;
    case VT_UI4:
    case VT_UINT:
      buffer.writeUInt32LE(toInteger(value, 0, 0xffffffff, vt), offset);
      return;
    case VT_ERROR:
      buffer.writeInt32LE(value instanceof ErrorValue ? value.scode | 0 : toInteger(value, -0x80000000, 0xffffffff, vt) | 0, offset);
      return;
    case VT_R4:
      buffer.writeFloatLE(Number(value), offset);
      return;
    case VT_R8:
      buffer.writeDoubleLE(Number(value), offset);
      return;
    case VT_BOOL:
      buffer.writeInt16LE(value ? -1 : 0, offset);
      return;
    case VT_CY:
      buffer.writeBigInt64LE(toCurrency(value), offset);
      return;
    case VT_DATE:
      buffer.writeDoubleLE(value instanceof Date ? dateToOADate(value) : Number(value), offset);
      return;
    case VT_BSTR:
      writePointer(buffer, offset, value === null || value === undefined ? 0n : allocBSTR(String(value)));
      return;
    case VT_DISPATCH:
    case VT_UNKNOWN:
      writePointer(buffer, offset, interfaceReference(value, vt));
      return;
    case VT_I8:
      buffer.writeBigInt64LE(toBigInt(value, -(2n ** 63n), 2n ** 63n - 1n, vt), offset);
      return;
    case VT_UI8:
      buffer.writeBigUInt64LE(toBigInt(value, 0n, 2n ** 64n - 1n, vt), offset);
      return;
    case VT_INT_PTR:
    case VT_UINT_PTR:
      writePointer(buffer, offset, BigInt(value as number | bigint));
      return;
    case VT_DECIMAL:
      writeDecimal(buffer, offset, value as string | number | bigint);
      return;
    case VT_VARIANT:
      writeVariant(buffer.subarray(offset, offset + VARIANT_SIZE), value);
      return;
    default:
      throw new TypeError(`Unsupported VARIANT type ${vartypeName(vt)}`);
  }
}

/**
 * Store `value` in the VARIANT in `buffer`, which must be empty (`VT_EMPTY`,
 * e.g. zero-filled or cleared). The VARIANT owns any allocated resource.
 */
export function writeVariant(buffer: Buffer, value: unknown): void {
  if (value instanceof ByRef) throw new TypeError("byRef() values are IDispatch arguments, not VARIANT values");
  if (value instanceof VARIANT) {
    checkHRESULT(oleaut32().VariantCopy(buffer, value._buffer), "VariantCopy");
    return;
  }
  const vt = inferVartype(value);
  const payload = value instanceof Typed ? value.value : value;
  if (vt & VT_ARRAY) {
    const dimensions = value instanceof Typed ? value.dimensions : 1;
    writePointer(buffer, DATA, createSafeArray(vt & VT_TYPEMASK, payload, dimensions));
  } else if (vt === VT_DECIMAL) {
    writeDecimal(buffer, 0, payload as string | number | bigint);
  } else if (vt === VT_RECORD) {
    const parts = recordParts(payload);
    if (!parts) throw new TypeError("VT_RECORD requires a record value: record(type, values)");
    const record = parts.info.create();
    try {
      parts.info.write(record, parts.values);
    } catch (error) {
      parts.info.destroy(record);
      throw error;
    }
    rawAddRef(parts.info.address);
    writePointer(buffer, DATA, record);
    writePointer(buffer, DATA + POINTER_SIZE, parts.info.address);
  } else {
    writeScalar(vt, buffer, DATA, payload);
  }
  buffer.writeUInt16LE(vt, 0);
}

function shapeOf(values: unknown, dimensions: number): number[] {
  const shape: number[] = [];
  let level: unknown = values;
  for (let dimension = 0; dimension < dimensions; dimension++) {
    if (!isList(level)) throw new TypeError(`SAFEARRAY dimension ${dimension + 1} is not an array`);
    shape.push(level.length);
    level = level[0];
  }
  return shape;
}

function isList(value: unknown): value is ArrayLike<unknown> {
  return Array.isArray(value) || (ArrayBuffer.isView(value) && !(value instanceof DataView));
}

/** Create a SAFEARRAY of `vt` from (nested) JavaScript arrays or a typed array. */
export function createSafeArray(vt: number, values: unknown, dimensions = 1): bigint {
  const api = oleaut32();
  const shape = shapeOf(values, dimensions);
  const bounds = Buffer.alloc(8 * dimensions);
  shape.forEach((length, dimension) => bounds.writeUInt32LE(length, dimension * 8));
  if (vt === VT_RECORD) return createRecordArray(values, dimensions, shape, bounds);
  const psa = api.SafeArrayCreate(vt, dimensions, bounds);
  if (!psa) throw new TypeError(`SafeArrayCreate failed for ${vartypeName(VT_ARRAY | vt)}`);
  const size = elementSize(vt);
  const count = shape.reduce((total, length) => total * length, 1);
  const data = Buffer.alloc(count * size);
  const target = new ctypes.c_void_p();
  checkHRESULT(api.SafeArrayAccessData(psa, ctypes.byref(target)), "SafeArrayAccessData");
  let filled = false;
  try {
    if (dimensions === 1 && ArrayBuffer.isView(values) && typedArrayVartype(values) === vt) {
      data.set(new Uint8Array(values.buffer, values.byteOffset, values.byteLength));
    } else {
      // Column-major storage: the first index varies fastest.
      const fill = (level: unknown, dimension: number, base: number, stride: number) => {
        if (!isList(level) || level.length !== shape[dimension]) throw new TypeError("SAFEARRAY values must be rectangular");
        for (let index = 0; index < level.length; index++) {
          const flat = base + index * stride;
          if (dimension === dimensions - 1) writeScalar(vt, data, flat * size, level[index]);
          else fill(level[index], dimension + 1, flat, stride * shape[dimension]);
        }
      };
      fill(values, 0, 0, 1);
    }
    filled = true;
  } finally {
    // Partially written elements are copied too, so that SafeArrayDestroy frees them on failure.
    writeMemory(target.value ?? 0n, data);
    api.SafeArrayUnaccessData(psa);
    if (!filled) api.SafeArrayDestroy(psa);
  }
  return psa;
}

/** SAFEARRAY of records (`SafeArrayCreateEx` with the `IRecordInfo` of the first element). */
function createRecordArray(values: unknown, dimensions: number, shape: readonly number[], bounds: Buffer): bigint {
  const api = oleaut32();
  const flat: unknown[] = [];
  const fill = (level: unknown, dimension: number, base: number, stride: number) => {
    if (!isList(level) || level.length !== shape[dimension]) throw new TypeError("SAFEARRAY values must be rectangular");
    for (let index = 0; index < level.length; index++) {
      const position = base + index * stride;
      if (dimension === dimensions - 1) flat[position] = level[index];
      else fill(level[index], dimension + 1, position, stride * shape[dimension]);
    }
  };
  fill(values, 0, 0, 1);
  const first = flat.map(recordParts).find((parts) => parts !== null);
  if (!first) throw new TypeError("A VT_RECORD SAFEARRAY needs record values: record(type, values)");
  const psa = api.SafeArrayCreateEx(VT_RECORD, dimensions, bounds, first.info.address);
  if (!psa) throw new TypeError("SafeArrayCreateEx failed for VT_RECORD");
  const size = api.SafeArrayGetElemsize(psa);
  const target = new ctypes.c_void_p();
  checkHRESULT(api.SafeArrayAccessData(psa, ctypes.byref(target)), "SafeArrayAccessData");
  try {
    flat.forEach((value, index) => {
      const parts = recordParts(value);
      if (!parts) throw new TypeError(`SAFEARRAY element ${index} is not a record`);
      first.info.write((target.value ?? 0n) + BigInt(index * size), parts.values);
    });
  } catch (error) {
    api.SafeArrayUnaccessData(psa);
    destroySafeArray(psa);
    throw error;
  }
  api.SafeArrayUnaccessData(psa);
  return psa;
}

/** Destroy a SAFEARRAY and the resources owned by its elements. */
export function destroySafeArray(psa: bigint | null): void {
  if (psa) checkHRESULT(oleaut32().SafeArrayDestroy(psa), "SafeArrayDestroy");
}

/** Release the resources of the VARIANT in `buffer` (`VariantClear`) and reset it to `VT_EMPTY`. */
export function clearVariant(buffer: Buffer): void {
  if (needsClear(buffer.readUInt16LE(0))) checkHRESULT(oleaut32().VariantClear(buffer), "VariantClear");
  buffer.fill(0, 0, VARIANT_SIZE);
}

// ── VARIANT ───────────────────────────────────────────────────────────

/**
 * OLE Automation VARIANT. Python equivalent: `comtypes.automation.VARIANT`.
 *
 * `value` converts to and from JavaScript:
 *
 * | JavaScript                 | VARIANT                                   |
 * | -------------------------- | ----------------------------------------- |
 * | `undefined` / `null`       | `VT_EMPTY` / `VT_NULL`                    |
 * | `boolean`                  | `VT_BOOL`                                 |
 * | `number`                   | `VT_I4` if a 32-bit integer, else `VT_R8` |
 * | `bigint`                   | `VT_I8` (`VT_UI8` above 2^63 - 1)         |
 * | `string`                   | `VT_BSTR`                                 |
 * | `Date`                     | `VT_DATE` (local time)                    |
 * | COM pointer / dispatch     | `VT_DISPATCH` or `VT_UNKNOWN`             |
 * | `Array` / typed array      | `VT_ARRAY` of `VT_VARIANT` / element type |
 * | `ErrorValue`, `missing`    | `VT_ERROR`                                |
 * | `typed(vt, value)`         | explicit VARTYPE                          |
 *
 * A VARIANT holding a BSTR, an interface or an array owns it: call `clear()`
 * or use `using`.
 */
export class VARIANT extends ctypes.Structure {
  static _fields_ = [
    ["vt", ctypes.c_uint16],
    ["wReserved1", ctypes.c_uint16],
    ["wReserved2", ctypes.c_uint16],
    ["wReserved3", ctypes.c_uint16],
    ["data", ctypes.array(ctypes.c_uint64, POINTER_SIZE === 8 ? 2 : 1)],
  ] as const;
  declare vt: number;
  declare wReserved1: number;
  declare wReserved2: number;
  declare wReserved3: number;

  /**
   * `new VARIANT(value)` converts `value`. As for any node-ctypes structure, a
   * `Buffer` (or a plain object of fields) is wrapped instead: node-ctypes
   * builds views (`from_address`, callback arguments) this way. Use
   * {@link VARIANT.from} to convert a `Buffer` into a `VT_ARRAY | VT_UI1`.
   */
  constructor(value?: unknown) {
    if (Buffer.isBuffer(value) || (value !== null && typeof value === "object" && Object.getPrototypeOf(value) === Object.prototype && !recordParts(value))) {
      super(value);
      return;
    }
    super();
    if (value !== undefined) writeVariant(this._buffer, value);
  }

  /** A new VARIANT holding `value`, whatever its type (`Buffer` included). */
  static from(value: unknown): VARIANT {
    const variant = new VARIANT();
    if (value !== undefined) writeVariant(variant._buffer, value);
    return variant;
  }

  /** The value converted to JavaScript; COM objects are new owned pointers. */
  get value(): unknown {
    return readVariant(this._buffer);
  }

  set value(value: unknown) {
    this.clear();
    writeVariant(this._buffer, value);
  }

  /** Exact text of a `VT_DECIMAL` or `VT_CY` value. */
  get exact(): string {
    return readVariantExact(this._buffer);
  }

  /** `VariantClear`: release owned resources and reset to `VT_EMPTY`. */
  clear(): void {
    clearVariant(this._buffer);
  }

  /**
   * Convert in place with `VariantChangeType`, or `VariantChangeTypeEx` when
   * `lcid` is given (text conversions otherwise follow the user locale).
   */
  changeType(vt: number, flags = 0, lcid?: number): this {
    const api = oleaut32();
    if (lcid === undefined) checkHRESULT(api.VariantChangeType(this._buffer, this._buffer, flags, vt), "VariantChangeType");
    else checkHRESULT(api.VariantChangeTypeEx(this._buffer, this._buffer, lcid, flags, vt), "VariantChangeTypeEx");
    return this;
  }

  /** Deep copy (`VariantCopy`). */
  copy(): VARIANT {
    const copy = new VARIANT();
    checkHRESULT(oleaut32().VariantCopy(copy._buffer, this._buffer), "VariantCopy");
    return copy;
  }

  [Symbol.dispose](): void {
    this.clear();
  }

  toString(): string {
    return `VARIANT(${vartypeName(this.vt)})`;
  }

  [Symbol.for("nodejs.util.inspect.custom")](): string {
    return this.toString();
  }
}
