import * as ctypes from "node-ctypes";
import { COMError, checkHRESULT } from "../errors.ts";
import { GUID } from "../guid.ts";
import { E_INVALIDARG } from "../hresult.ts";
import { IUnknown } from "../interfaces/unknown.ts";
import { oleaut32, POINTER_SIZE, takeBSTR } from "../low-level/native.ts";
import { functionCache, vtableEntry } from "../low-level/vtable.ts";
import { COMPointer } from "../pointer.ts";
import { recordTypeInfoOf } from "../typelib/coclass.ts";
import { VARIANT_SIZE, clearVariant, readVariant, writeVariant } from "./variant.ts";

// User-defined types (VT_RECORD) through IRecordInfo, called by vtable slot
// (no interface definition, so that the VARIANT codec can use it).

const { c_int32, c_uint32, c_void_p, c_wchar_p } = ctypes;
const INVOKE_PROPERTYPUT = 4;
const SLOT = { RecordInit: 3, RecordClear: 4, GetGuid: 6, GetName: 7, GetSize: 8, GetField: 10, PutField: 12, GetFieldNames: 14, RecordCreate: 16, RecordDestroy: 18 } as const;

const PointerArgAt = functionCache(() => ctypes.WINFUNCTYPE(c_int32, c_void_p, c_void_p));
const FieldNamesAt = functionCache(() => ctypes.WINFUNCTYPE(c_int32, c_void_p, c_void_p, c_void_p));
const GetFieldAt = functionCache(() => ctypes.WINFUNCTYPE(c_int32, c_void_p, c_void_p, c_wchar_p, c_void_p));
const PutFieldAt = functionCache(() => ctypes.WINFUNCTYPE(c_int32, c_void_p, c_uint32, c_void_p, c_wchar_p, c_void_p));
const CreateAt = functionCache(() => ctypes.WINFUNCTYPE(c_void_p, c_void_p));

/** The `IRecordInfo` of a value read from a `VT_RECORD`, kept for writing it back. */
export const recordInfo: unique symbol = Symbol.for("node-comtypes.recordInfo") as never;

// IRecordInfo address -> wrapper. Record infos are per type (oleaut32 caches
// them), so this stays small; the pointers are released with the apartment.
const shared = new Map<bigint, RecordInfo>();

/** A user-defined type described by `IRecordInfo`. */
export class RecordInfo {
  /** Owned `IRecordInfo` reference. */
  readonly pointer: COMPointer;
  #fields: readonly string[] | null = null;
  #name: string | null = null;

  private constructor(pointer: COMPointer) {
    this.pointer = pointer;
  }

  /** Wrap an `IRecordInfo` address (one shared, owned wrapper per record type and apartment). */
  static borrow(address: bigint): RecordInfo {
    let info = shared.get(address);
    if (!info || info.pointer.released) shared.set(address, (info = new RecordInfo(COMPointer.fromAddress(address, IUnknown, { addRef: true }))));
    return info;
  }

  /** `GetRecordInfoFromTypeInfo`: the type must be a record with a GUID. */
  static fromTypeInfo(typeInfo: bigint): RecordInfo {
    const out = new c_void_p();
    checkHRESULT(oleaut32().GetRecordInfoFromTypeInfo(typeInfo, ctypes.byref(out)), "GetRecordInfoFromTypeInfo");
    const info = RecordInfo.borrow(out.value!);
    COMPointer.fromAddress(out.value!, IUnknown).Release();
    return info;
  }

  get address(): bigint {
    return this.pointer.address;
  }

  get name(): string {
    if (this.#name === null) {
      const out = new c_void_p();
      checkHRESULT(PointerArgAt(vtableEntry(this.address, SLOT.GetName))(this.address, ctypes.byref(out)), "IRecordInfo::GetName");
      this.#name = takeBSTR(out.value) ?? "";
    }
    return this.#name;
  }

  get guid(): GUID {
    const guid = new GUID();
    checkHRESULT(PointerArgAt(vtableEntry(this.address, SLOT.GetGuid))(this.address, guid._buffer), "IRecordInfo::GetGuid");
    return guid;
  }

  get size(): number {
    const size = new c_uint32();
    checkHRESULT(PointerArgAt(vtableEntry(this.address, SLOT.GetSize))(this.address, ctypes.byref(size)), "IRecordInfo::GetSize");
    return size.value;
  }

  get fields(): readonly string[] {
    if (this.#fields) return this.#fields;
    const count = new c_uint32(0);
    const call = FieldNamesAt(vtableEntry(this.address, SLOT.GetFieldNames));
    checkHRESULT(call(this.address, ctypes.byref(count), null), "IRecordInfo::GetFieldNames");
    const names = Buffer.alloc(POINTER_SIZE * count.value);
    checkHRESULT(call(this.address, ctypes.byref(count), names), "IRecordInfo::GetFieldNames");
    this.#fields = Object.freeze(Array.from({ length: count.value }, (_, index) => takeBSTR(ctypes.readValue(names, c_void_p, index * POINTER_SIZE)) ?? ""));
    return this.#fields;
  }

  /** `RecordCreate`: a new initialized record, destroyed by `VariantClear` / `SafeArrayDestroy` once stored. */
  create(): bigint {
    const record = CreateAt(vtableEntry(this.address, SLOT.RecordCreate))(this.address);
    if (!record) throw new COMError(E_INVALIDARG, { method: "IRecordInfo::RecordCreate" });
    return record;
  }

  destroy(record: bigint): void {
    PointerArgAt(vtableEntry(this.address, SLOT.RecordDestroy))(this.address, record);
  }

  /** Read every field of the record at `record`. */
  read(record: bigint): RecordValue {
    const value: Record<string, unknown> = {};
    const slot = Buffer.alloc(VARIANT_SIZE);
    const get = GetFieldAt(vtableEntry(this.address, SLOT.GetField));
    for (const field of this.fields) {
      checkHRESULT(get(this.address, record, field, slot), `IRecordInfo::GetField(${field})`);
      try {
        value[field] = readVariant(slot);
      } finally {
        clearVariant(slot);
      }
    }
    Object.defineProperty(value, recordInfo, { value: this, enumerable: false });
    return value as RecordValue;
  }

  /** Write the fields present in `values` (by name, case-insensitive) into the record at `record`. */
  write(record: bigint, values: Readonly<Record<string, unknown>>): void {
    const put = PutFieldAt(vtableEntry(this.address, SLOT.PutField));
    const keys = new Map(Object.keys(values).map((key) => [key.toLowerCase(), key]));
    const slot = Buffer.alloc(VARIANT_SIZE);
    for (const field of this.fields) {
      const key = keys.get(field.toLowerCase());
      if (key === undefined || values[key] === undefined) continue;
      writeVariant(slot, values[key]);
      try {
        checkHRESULT(put(this.address, INVOKE_PROPERTYPUT, record, field, slot), `IRecordInfo::PutField(${field})`);
      } finally {
        clearVariant(slot);
      }
    }
  }
}

/** A record value: its fields, plus the hidden `IRecordInfo` that describes it. */
export type RecordValue = Record<string, unknown> & { readonly [recordInfo]: RecordInfo };

/** Where a record type comes from: an `IRecordInfo`, a record class of `GetModule` / a generated module, or an `ITypeInfo` address. */
export type RecordSource = RecordInfo | object | bigint;

const infos = new WeakMap<object, RecordInfo>();

export function recordInfoOf(source: RecordSource): RecordInfo {
  if (source instanceof RecordInfo) return source;
  if (typeof source === "bigint") return RecordInfo.fromTypeInfo(source);
  let info = infos.get(source);
  if (info && !info.pointer.released) return info;
  const typeInfo = recordTypeInfoOf(source);
  if (!typeInfo) throw new TypeError("Expected a RecordInfo, a record class built from a type library or an ITypeInfo address");
  const address = typeInfo();
  try {
    info = RecordInfo.fromTypeInfo(address);
  } finally {
    COMPointer.fromAddress(address, IUnknown).Release();
  }
  infos.set(source, info);
  return info;
}

/**
 * A record value for `VT_RECORD`: `record(Drawing.Point, { x: 1, y: 2 })`.
 * The type needs a GUID in its type library (an `IRecordInfo` requirement).
 */
export function record(type: RecordSource, values: Readonly<Record<string, unknown>> = {}): RecordValue {
  const result: Record<string, unknown> = { ...values };
  Object.defineProperty(result, recordInfo, { value: recordInfoOf(type), enumerable: false });
  return result as RecordValue;
}

/** The `IRecordInfo` and field values of a JavaScript record value, if it is one. */
export function recordParts(value: unknown): { readonly info: RecordInfo; readonly values: Readonly<Record<string, unknown>> } | null {
  if (value === null || typeof value !== "object") return null;
  const info = (value as Partial<RecordValue>)[recordInfo];
  if (info instanceof RecordInfo) return { info, values: value as Record<string, unknown> };
  // Instances of record classes built from a type library.
  const type = (value as object).constructor;
  if (type && recordTypeInfoOf(type) && typeof (value as { toObject?: unknown }).toObject === "function") return { info: recordInfoOf(type), values: (value as { toObject(): Record<string, unknown> }).toObject() };
  return null;
}
