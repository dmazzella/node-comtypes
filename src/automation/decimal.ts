import * as ctypes from "node-ctypes";

/**
 * 96-bit scaled decimal (`VT_DECIMAL`). It overlays a whole VARIANT: the
 * first 2 bytes are the VARIANT `vt`.
 */
export class DECIMAL extends ctypes.Structure {
  static _fields_ = [
    ["wReserved", ctypes.c_uint16],
    ["scale", ctypes.c_uint8],
    ["sign", ctypes.c_uint8],
    ["Hi32", ctypes.c_uint32],
    ["Lo64", ctypes.c_uint64],
  ] as const;
  declare wReserved: number;
  declare scale: number;
  declare sign: number;
  declare Hi32: number;
  declare Lo64: bigint;
}

export const DECIMAL_NEG = 0x80;
const MAX_MANTISSA = (1n << 96n) - 1n;
const MAX_SCALE = 28;
const PATTERN = /^([+-])?(\d*)(?:\.(\d*))?(?:e([+-]?\d+))?$/i;

/** Exact decimal string of the DECIMAL at `offset`. */
export function readDecimal(buffer: Buffer, offset = 0): string {
  const scale = buffer.readUInt8(offset + 2);
  const negative = (buffer.readUInt8(offset + 3) & DECIMAL_NEG) !== 0;
  const mantissa = (BigInt(buffer.readUInt32LE(offset + 4)) << 64n) | buffer.readBigUInt64LE(offset + 8);
  if (scale > MAX_SCALE) throw new RangeError(`Invalid DECIMAL scale ${scale}`);
  let digits = mantissa.toString();
  if (scale > 0) {
    digits = digits.padStart(scale + 1, "0");
    digits = `${digits.slice(0, -scale)}.${digits.slice(-scale)}`.replace(/\.?0+$/, "");
  }
  return negative && mantissa !== 0n ? `-${digits}` : digits;
}

/**
 * Write `value` as a DECIMAL at `offset`, leaving the first 2 bytes (the
 * VARIANT `vt`) to the caller. Values with more than 28 decimals are rounded
 * half away from zero.
 */
export function writeDecimal(buffer: Buffer, offset: number, value: string | number | bigint): void {
  let negative = false;
  let mantissa: bigint;
  let scale = 0;
  if (typeof value === "bigint") {
    negative = value < 0n;
    mantissa = negative ? -value : value;
  } else {
    if (typeof value === "number" && !Number.isFinite(value)) throw new RangeError("DECIMAL cannot represent NaN or Infinity");
    const text = typeof value === "number" ? numberText(value) : value.trim();
    const match = PATTERN.exec(text);
    if (!match || (!match[2] && !match[3])) throw new SyntaxError(`Invalid decimal number: ${JSON.stringify(value)}`);
    negative = match[1] === "-";
    const fraction = match[3] ?? "";
    mantissa = BigInt((match[2] || "0") + fraction);
    scale = fraction.length - Number(match[4] ?? 0);
    if (scale < 0) {
      mantissa *= 10n ** BigInt(-scale);
      scale = 0;
    }
  }
  while (scale > 0 && (scale > MAX_SCALE || mantissa > MAX_MANTISSA)) {
    mantissa = (mantissa + 5n) / 10n;
    scale--;
  }
  if (mantissa > MAX_MANTISSA) throw new RangeError("Value outside the DECIMAL range");
  buffer.writeUInt8(scale, offset + 2);
  buffer.writeUInt8(negative && mantissa !== 0n ? DECIMAL_NEG : 0, offset + 3);
  buffer.writeUInt32LE(Number(mantissa >> 64n), offset + 4);
  buffer.writeBigUInt64LE(mantissa & 0xffffffffffffffffn, offset + 8);
}

/** Shortest round-trip text of a number, without exponent loss (`1e-7` stays exact). */
function numberText(value: number): string {
  return Object.is(value, -0) ? "0" : String(value);
}

/** Exact string of a currency (`VT_CY`) value: a 64-bit integer scaled by 10 000. */
export function currencyToString(value: bigint): string {
  const negative = value < 0n;
  const digits = (negative ? -value : value).toString().padStart(5, "0");
  const text = `${digits.slice(0, -4)}.${digits.slice(-4)}`.replace(/\.?0+$/, "");
  return negative ? `-${text}` : text;
}
