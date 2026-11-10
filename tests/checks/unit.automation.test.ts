import assert from "node:assert/strict";
import { describe, it } from "node:test";
import * as ctypes from "node-ctypes";

import {
  DECIMAL,
  DISPPARAMS,
  ELEMDESC,
  EXCEPINFO,
  ErrorValue,
  FUNCDESC,
  PARAMDESCEX,
  SAFEARRAYBOUND,
  TLIBATTR,
  TYPEATTR,
  TYPEDESC,
  VARDESC,
  VARIANT,
  VARIANT_SIZE,
  VT_ARRAY,
  VT_BOOL,
  VT_BSTR,
  VT_BYREF,
  VT_CY,
  VT_DATE,
  VT_DECIMAL,
  VT_EMPTY,
  VT_ERROR,
  VT_I1,
  VT_I4,
  VT_I8,
  VT_NULL,
  VT_R4,
  VT_R8,
  VT_UI1,
  VT_UI2,
  VT_UI8,
  VT_VARIANT,
  DISP_E_PARAMNOTFOUND,
  currencyToString,
  dateToOADate,
  inferVartype,
  missing,
  oaDateToDate,
  readDecimal,
  typed,
  vartypeName,
  writeDecimal,
  GUID,
} from "@node-ctypes/comtypes";

const x64 = ctypes.sizeof(ctypes.c_void_p) === 8;

describe("automation layouts", () => {
  it("VARIANT is 16 bytes on 32-bit and 24 on 64-bit, 8-aligned", () => {
    assert.equal(ctypes.sizeof(VARIANT), x64 ? 24 : 16);
    assert.equal(VARIANT_SIZE, ctypes.sizeof(VARIANT));
    assert.equal(ctypes.alignment(VARIANT), 8);
  });

  it("DECIMAL overlays a VARIANT", () => {
    assert.equal(ctypes.sizeof(DECIMAL), 16);
    const fields = Object.fromEntries(DECIMAL.fields.map((field) => [field.name, field.offset]));
    assert.deepEqual(fields, { wReserved: 0, scale: 2, sign: 3, Hi32: 4, Lo64: 8 });
  });

  it("automation structures match the Windows SDK sizes", () => {
    const expected: Array<[string, ctypes.StructureClass, number, number]> = [
      ["DISPPARAMS", DISPPARAMS, 24, 16],
      ["EXCEPINFO", EXCEPINFO, 64, 32],
      ["SAFEARRAYBOUND", SAFEARRAYBOUND, 8, 8],
      ["TYPEDESC", TYPEDESC, 16, 8],
      ["ELEMDESC", ELEMDESC, 32, 16],
      ["FUNCDESC", FUNCDESC, 88, 52],
      ["VARDESC", VARDESC, 64, 36],
      ["TYPEATTR", TYPEATTR, 96, 76],
      ["TLIBATTR", TLIBATTR, 32, 32],
      ["PARAMDESCEX", PARAMDESCEX, 32, 24],
    ];
    for (const [name, type, size64, size32] of expected) assert.equal(ctypes.sizeof(type), x64 ? size64 : size32, name);
  });
});

describe("DECIMAL", () => {
  const roundTrip = (value: string | number | bigint) => {
    const buffer = Buffer.alloc(16);
    writeDecimal(buffer, 0, value);
    return readDecimal(buffer, 0);
  };

  it("round-trips exact values", () => {
    assert.equal(roundTrip("0"), "0");
    assert.equal(roundTrip("-0.000"), "0");
    assert.equal(roundTrip("123.4500"), "123.45");
    assert.equal(roundTrip("-0.0000000000000000000000000001"), "-0.0000000000000000000000000001");
    assert.equal(roundTrip("79228162514264337593543950335"), "79228162514264337593543950335");
    assert.equal(roundTrip(12345678901234567890n), "12345678901234567890");
    assert.equal(roundTrip(0.1), "0.1");
    assert.equal(roundTrip(1e-7), "0.0000001");
    assert.equal(roundTrip("1.5e3"), "1500");
  });

  it("encodes scale, sign and the 96-bit mantissa", () => {
    const buffer = Buffer.alloc(16);
    writeDecimal(buffer, 0, "-1.25");
    assert.equal(buffer.readUInt8(2), 2);
    assert.equal(buffer.readUInt8(3), 0x80);
    assert.equal(buffer.readUInt32LE(4), 0);
    assert.equal(buffer.readBigUInt64LE(8), 125n);
  });

  it("rounds excess precision and rejects out-of-range values", () => {
    assert.equal(roundTrip("0.12345678901234567890123456789"), "0.1234567890123456789012345679");
    assert.throws(() => roundTrip("79228162514264337593543950336"), RangeError);
    assert.throws(() => roundTrip("abc"), SyntaxError);
    assert.throws(() => roundTrip(Number.NaN), RangeError);
  });

  it("formats currency exactly", () => {
    assert.equal(currencyToString(15000n), "1.5");
    assert.equal(currencyToString(-1n), "-0.0001");
    assert.equal(currencyToString(0n), "0");
    assert.equal(currencyToString(9223372036854775807n), "922337203685477.5807");
  });
});

describe("OLE Automation DATE", () => {
  it("converts local wall-clock dates", () => {
    assert.equal(dateToOADate(new Date(1899, 11, 30)), 0);
    assert.equal(dateToOADate(new Date(2000, 0, 1)), 36526);
    assert.equal(dateToOADate(new Date(1900, 0, 1, 6)), 2.25);
    assert.equal(dateToOADate(new Date(1899, 11, 29, 12)), -1.5);
    assert.equal(dateToOADate(new Date(1899, 11, 28, 18)), -2.75);
  });

  it("converts back, including dates before 1899-12-30", () => {
    assert.equal(oaDateToDate(36526).getTime(), new Date(2000, 0, 1).getTime());
    assert.equal(oaDateToDate(-1.5).getTime(), new Date(1899, 11, 29, 12).getTime());
    assert.equal(oaDateToDate(2.25).getTime(), new Date(1900, 0, 1, 6).getTime());
    const date = new Date(2024, 1, 29, 23, 59, 58, 123);
    assert.equal(oaDateToDate(dateToOADate(date)).getTime(), date.getTime());
    const old = new Date(2000, 0, 1);
    old.setFullYear(150);
    assert.equal(oaDateToDate(dateToOADate(old)).getTime(), old.getTime());
  });

  it("rejects invalid or out-of-range values", () => {
    assert.throws(() => dateToOADate(new Date(Number.NaN)), RangeError);
    assert.throws(() => oaDateToDate(Number.POSITIVE_INFINITY), RangeError);
    assert.throws(() => oaDateToDate(3e6), RangeError);
  });
});

describe("VARIANT scalar conversions", () => {
  const value = (input: unknown) => new VARIANT(input).value;

  it("infers VARTYPEs from JavaScript values", () => {
    assert.equal(inferVartype(undefined), VT_EMPTY);
    assert.equal(inferVartype(null), VT_NULL);
    assert.equal(inferVartype(true), VT_BOOL);
    assert.equal(inferVartype(42), VT_I4);
    assert.equal(inferVartype(-0x80000000), VT_I4);
    assert.equal(inferVartype(0x80000000), VT_R8);
    assert.equal(inferVartype(1.5), VT_R8);
    assert.equal(inferVartype(-0), VT_R8);
    assert.equal(inferVartype(1n), VT_I8);
    assert.equal(inferVartype(2n ** 63n), VT_UI8);
    assert.equal(inferVartype("x"), VT_BSTR);
    assert.equal(inferVartype(new Date()), VT_DATE);
    assert.equal(inferVartype(missing), VT_ERROR);
    assert.equal(inferVartype([1]), VT_ARRAY | VT_VARIANT);
    assert.equal(inferVartype(new Uint8Array(1)), VT_ARRAY | VT_UI1);
    assert.equal(inferVartype(typed(VT_UI2, 1)), VT_UI2);
    assert.throws(() => inferVartype({}), TypeError);
    assert.throws(() => inferVartype(Symbol("x")), TypeError);
  });

  it("round-trips scalars without native calls", () => {
    assert.equal(value(undefined), undefined);
    assert.equal(value(null), null);
    assert.equal(value(true), true);
    assert.equal(value(false), false);
    assert.equal(value(-7), -7);
    assert.equal(value(2.5), 2.5);
    assert.equal(value(-(2n ** 63n)), -(2n ** 63n));
    assert.equal(value(2n ** 64n - 1n), 2n ** 64n - 1n);
    const date = new Date(2021, 5, 1, 8, 30);
    assert.equal((value(date) as Date).getTime(), date.getTime());
    const error = value(missing) as ErrorValue;
    assert.ok(error instanceof ErrorValue);
    assert.equal(error.scode, DISP_E_PARAMNOTFOUND);
  });

  it("stores VT_BOOL as VARIANT_TRUE and the vt last", () => {
    const variant = new VARIANT(true);
    assert.equal(variant.vt, VT_BOOL);
    assert.equal(variant._buffer.readInt16LE(8), -1);
    assert.equal(new VARIANT(42)._buffer.readInt32LE(8), 42);
  });

  it("honours explicit types", () => {
    assert.equal(value(typed(VT_I1, -5)), -5);
    assert.equal(value(typed(VT_UI1, 255)), 255);
    assert.equal(value(typed(VT_R4, 0.5)), 0.5);
    assert.equal(value(typed(VT_I8, 5)), 5n);
    assert.equal(value(typed(VT_R8, 3)), 3);
    assert.equal(value(typed(VT_CY, "12.3456")), 12.3456);
    assert.equal(new VARIANT(typed(VT_CY, -1.5)).exact, "-1.5");
    const decimal = new VARIANT(typed(VT_DECIMAL, "123456789012345678901234.5678"));
    assert.equal(decimal.vt, VT_DECIMAL);
    assert.equal(decimal.exact, "123456789012345678901234.5678");
    assert.equal(decimal.value, 123456789012345678901234.5678);
    assert.throws(() => typed(VT_I4 | VT_BYREF, 1), RangeError);
    assert.throws(() => new VARIANT(typed(VT_UI1, 256)), RangeError);
    assert.throws(() => new VARIANT(typed(VT_I4, 1.5)), RangeError);
  });

  it("clears scalars and names VARTYPEs", () => {
    const variant = new VARIANT(1.5);
    variant.clear();
    assert.equal(variant.vt, VT_EMPTY);
    variant.value = 3;
    assert.equal(variant.value, 3);
    assert.equal(String(new VARIANT(1)), "VARIANT(VT_I4)");
    assert.equal(vartypeName(VT_ARRAY | VT_BSTR), "VT_ARRAY | VT_BSTR");
    assert.equal(vartypeName(VT_BYREF | VT_VARIANT), "VT_BYREF | VT_VARIANT");
  });
});

describe("structure views", () => {
  it("VARIANT and GUID wrap Buffers, as node-ctypes views do", () => {
    const backing = Buffer.alloc(VARIANT_SIZE);
    new VARIANT(-42)._buffer.copy(backing);
    const view = VARIANT.from_address(ctypes.addressof(backing));
    assert.equal(view.vt, VT_I4);
    assert.equal(view.value, -42);
    view.value = 7.5;
    assert.equal(backing.readUInt16LE(0), VT_R8, "writes go to the wrapped memory");
    assert.equal(new VARIANT(backing).value, 7.5);
    // VARIANT.from(Buffer) creates a SAFEARRAY (oleaut32, Windows only): tests/src/variant.test.ts.
    const bytes = Buffer.alloc(16);
    const guid = new GUID(bytes);
    guid.Data1 = 0x12345678;
    assert.equal(bytes.readUInt32LE(0), 0x12345678);
    assert.equal(new GUID(new Uint8Array(bytes)).Data1, 0x12345678);
  });
});
