import assert from "node:assert/strict";
import { describe, it } from "node:test";

import {
  COMPointer,
  CreateObject,
  IDispatch,
  IUnknown,
  VARIANT,
  VT_ARRAY,
  VT_BSTR,
  VT_DATE,
  VT_DECIMAL,
  VT_DISPATCH,
  VT_I2,
  VT_I4,
  VT_R8,
  VT_UI1,
  VT_UNKNOWN,
  VT_VARIANT,
  apartment,
  dispatchPointer,
  getLivePointerCount,
  safeArray,
  typed,
} from "@node-ctypes/comtypes";
import { allocBSTR, oleaut32, readBSTR } from "@node-ctypes/comtypes/low-level";
import { COMObject } from "@node-ctypes/comtypes/server";

const skip = process.platform !== "win32" && "COM requires Windows";

function refCount(pointer: COMPointer): number {
  pointer.AddRef();
  return pointer.callMethod("Release") as number;
}

describe("BSTR", { skip }, () => {
  it("preserves embedded NULs and non-BMP characters", () => {
    for (const text of ["", "plain", "a\0b", "𝄞 – ü"]) {
      const bstr = allocBSTR(text);
      try {
        assert.equal(readBSTR(bstr), text);
        assert.equal(oleaut32().SysStringLen(bstr), text.length);
      } finally {
        oleaut32().SysFreeString(bstr);
      }
    }
    assert.equal(readBSTR(null), null);
  });

  it("VARIANT owns its BSTR", () => {
    using variant = new VARIANT("héllo\0world");
    assert.equal(variant.vt, VT_BSTR);
    assert.equal(variant.value, "héllo\0world");
    variant.value = "replaced";
    assert.equal(variant.value, "replaced");
  });
});

describe("SAFEARRAY", { skip }, () => {
  it("destroys a partially filled array when a value cannot be converted", () => {
    using _ = apartment();
    class Item extends COMObject {
      static _com_interfaces_ = [IDispatch];
    }
    const item = new Item();
    assert.throws(() => new VARIANT(safeArray([item, 42], { vt: VT_DISPATCH })), TypeError);
    assert.equal(item.refCount, 0, "the reference written before the failure is released");
  });

  it("converts JavaScript arrays to VARIANT arrays and back", () => {
    const values = [1, "two", 3.5, true, null, undefined, 9n, [1, ["nested"]]];
    using variant = new VARIANT(values);
    assert.equal(variant.vt, VT_ARRAY | VT_VARIANT);
    assert.deepEqual(variant.value, values);
    using empty = new VARIANT([]);
    assert.deepEqual(empty.value, []);
  });

  it("uses the element type of typed arrays and Buffers", () => {
    using bytes = VARIANT.from(Buffer.from([0, 127, 255]));
    assert.equal(bytes.vt, VT_ARRAY | VT_UI1);
    assert.deepEqual(bytes.value, [0, 127, 255]);
    using doubles = new VARIANT(new Float64Array([0.5, -1]));
    assert.equal(doubles.vt, VT_ARRAY | VT_R8);
    assert.deepEqual(doubles.value, [0.5, -1]);
    using shorts = new VARIANT(safeArray([1, -2, 3], { vt: VT_I2 }));
    assert.deepEqual(shorts.value, [1, -2, 3]);
    using strings = new VARIANT(safeArray(["a", "", "c"], { vt: VT_BSTR }));
    assert.deepEqual(strings.value, ["a", "", "c"]);
  });

  it("keeps the index order of multi-dimensional arrays", () => {
    const matrix = [
      [1, 2, 3],
      [4, 5, 6],
    ];
    using two = new VARIANT(typed(VT_ARRAY | VT_I4, matrix, { dimensions: 2 }));
    assert.deepEqual(two.value, matrix);
    // Column-major storage: (0,0), (1,0), (0,1), ...
    const cube = [
      [
        ["a", "b"],
        ["c", "d"],
      ],
      [
        ["e", "f"],
        ["g", "h"],
      ],
    ];
    using three = new VARIANT(safeArray(cube, { dimensions: 3 }));
    assert.deepEqual(three.value, cube);
    assert.throws(() => new VARIANT(safeArray([[1, 2], [3]], { dimensions: 2 })), /rectangular/);
  });

  it("reads arrays returned by automation objects", () => {
    using _ = apartment();
    using dict = CreateObject("Scripting.Dictionary");
    dict.Add("k1", [1, 2]);
    dict.Add("k2", "v");
    assert.deepEqual(dict.Keys(), ["k1", "k2"]);
    assert.deepEqual(dict.Items(), [[1, 2], "v"]);
  });
});

describe("VARIANT native operations", { skip }, () => {
  it("changes types with VariantChangeType", () => {
    using number = new VARIANT(42);
    number.changeType(VT_BSTR);
    assert.equal(number.value, "42");
    using text = new VARIANT("12345678901234567890.125");
    text.changeType(VT_DECIMAL, 0, 0x409);
    assert.equal(text.exact, "12345678901234567890.125");
    using date = new VARIANT(new Date(2000, 0, 1, 12));
    assert.equal(date.vt, VT_DATE);
    date.changeType(VT_R8);
    assert.equal(date.value, 36526.5);
  });

  it("copies deeply", () => {
    using source = new VARIANT(["a", ["b"]]);
    using copy = source.copy();
    source.clear();
    assert.deepEqual(copy.value, ["a", ["b"]]);
  });

  it("holds interface references", () => {
    using _ = apartment();
    using dict = CreateObject("Scripting.Dictionary");
    const pointer = dispatchPointer(dict);
    const base = refCount(pointer);
    {
      using variant = new VARIANT(dict);
      assert.equal(variant.vt, VT_DISPATCH);
      assert.equal(refCount(pointer), base + 1);
      using read = variant.value as COMPointer;
      assert.ok(read.isSameObject(pointer));
      assert.equal(read.interface, IDispatch);
    }
    assert.equal(refCount(pointer), base);
    using unknown = pointer.QueryInterface(IUnknown);
    using variant = new VARIANT(unknown);
    assert.equal(variant.vt, VT_UNKNOWN);
    using dispatch = new VARIANT(typed(VT_DISPATCH, unknown));
    assert.equal(dispatch.vt, VT_DISPATCH);
  });

  it("does not leak across many conversions", () => {
    using _ = apartment();
    const baseline = getLivePointerCount();
    for (let i = 0; i < 500; i++) {
      using dict = CreateObject("Scripting.Dictionary");
      dict.Add("list", [i, "x", [i]]);
      dict.Add("self", dict);
      const self = dict.Item("self");
      assert.equal(self.Item("list")[0], i);
      self[Symbol.dispose]();
      dict.RemoveAll();
    }
    assert.equal(getLivePointerCount(), baseline);
  });
});
