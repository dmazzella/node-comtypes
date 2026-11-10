import * as ctypes from "node-ctypes";
import { ole32, takeCoTaskMemString } from "./low-level/native.ts";
import { checkHRESULT } from "./errors.ts";

const GUID_PATTERN = /^\{?([0-9a-f]{8})-([0-9a-f]{4})-([0-9a-f]{4})-([0-9a-f]{4})-([0-9a-f]{12})\}?$/i;

/**
 * COM GUID (16 bytes). Python equivalent: `comtypes.GUID`.
 *
 * ```ts
 * new GUID("{00000000-0000-0000-C000-000000000046}");
 * new GUID();          // GUID_NULL
 * new GUID(buffer);    // view of 16 bytes (other Uint8Arrays are copied)
 * ```
 */
export class GUID extends ctypes.Structure {
  static _fields_ = [
    ["Data1", ctypes.c_uint32],
    ["Data2", ctypes.c_uint16],
    ["Data3", ctypes.c_uint16],
    ["Data4", ctypes.array(ctypes.c_uint8, 8)],
  ] as const;
  declare Data1: number;
  declare Data2: number;
  declare Data3: number;
  declare Data4: ctypes.ArrayInstance<typeof ctypes.c_uint8>;

  constructor(init?: string | Uint8Array | GUID) {
    // A Buffer is wrapped (node-ctypes views: from_address, nested fields); other inputs are copied.
    if (Buffer.isBuffer(init)) {
      if (init.length !== 16) throw new TypeError("GUID requires a GUID string, another GUID or exactly 16 bytes");
      super(init);
      return;
    }
    super();
    if (init === undefined) return;
    if (typeof init === "string") {
      const match = GUID_PATTERN.exec(init.trim());
      if (!match) throw new SyntaxError(`Invalid GUID string: ${JSON.stringify(init)}`);
      this.Data1 = parseInt(match[1], 16);
      this.Data2 = parseInt(match[2], 16);
      this.Data3 = parseInt(match[3], 16);
      Buffer.from(match[4] + match[5], "hex").copy(this._buffer, 8);
      return;
    }
    const bytes = init instanceof GUID ? init._buffer : init;
    if (!(bytes instanceof Uint8Array) || bytes.length !== 16) throw new TypeError("GUID requires a GUID string, another GUID or exactly 16 bytes");
    this._buffer.set(bytes);
  }

  /** A new random GUID (`CoCreateGuid`). Python: `GUID.create_new()`. */
  static createNew(): GUID {
    const guid = new GUID();
    checkHRESULT(ole32().CoCreateGuid(guid._buffer), "CoCreateGuid");
    return guid;
  }

  /** CLSID registered for a ProgID (`CLSIDFromProgID`). Python: `GUID.from_progid()`. */
  static fromProgID(progid: string): GUID {
    const guid = new GUID();
    checkHRESULT(ole32().CLSIDFromProgID(progid, guid._buffer), "CLSIDFromProgID");
    return guid;
  }

  /** ProgID registered for this CLSID (`ProgIDFromCLSID`). Python: `GUID.as_progid()`. */
  asProgID(): string {
    const out = new ctypes.c_void_p();
    checkHRESULT(ole32().ProgIDFromCLSID(this._buffer, ctypes.byref(out)), "ProgIDFromCLSID");
    return takeCoTaskMemString(out.value) ?? "";
  }

  /** Whether `value` is a well-formed GUID string. */
  static isGUIDString(value: string): boolean {
    return GUID_PATTERN.test(value.trim());
  }

  equals(other: GUIDLike | null | undefined): boolean {
    if (other === null || other === undefined) return false;
    return toGUID(other)._buffer.equals(this._buffer);
  }

  isNull(): boolean {
    return this._buffer.every((byte) => byte === 0);
  }

  /** `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}` (uppercase, as `StringFromGUID2`). */
  toString(): string {
    const hex = (value: number, width: number) => value.toString(16).padStart(width, "0");
    const tail = this._buffer.subarray(8, 16).toString("hex");
    return `{${hex(this.Data1, 8)}-${hex(this.Data2, 4)}-${hex(this.Data3, 4)}-${tail.slice(0, 4)}-${tail.slice(4)}}`.toUpperCase();
  }

  [Symbol.for("nodejs.util.inspect.custom")](): string {
    return `GUID("${this.toString()}")`;
  }
}

/** A GUID or a GUID string. */
export type GUIDLike = GUID | string;

export function toGUID(value: GUIDLike): GUID {
  if (value instanceof GUID) return value;
  if (typeof value === "string") return new GUID(value);
  throw new TypeError("Expected a GUID or a GUID string");
}

export const GUID_NULL: GUID = new GUID();
