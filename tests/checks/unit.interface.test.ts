import assert from "node:assert/strict";
import { describe, it } from "node:test";
import * as ctypes from "node-ctypes";

import { COMInterface, COMPointer, HRESULT, IClassFactory, IPersist, IPersistFile, IShellLinkW, IStream, IUnknown, STATSTG, STDMETHOD } from "@node-ctypes/comtypes";

describe("COMInterface", () => {
  it("lays out inherited methods first", () => {
    assert.deepEqual(
      IUnknown.allMethods.map((method) => method.name),
      ["QueryInterface", "AddRef", "Release"],
    );
    assert.equal(IPersist.methodIndex.get("GetClassID"), 3);
    assert.equal(IPersistFile.methodIndex.get("IsDirty"), 4);
    assert.equal(IPersistFile.methodIndex.get("GetCurFile"), 8);
    assert.equal(IPersistFile.allMethods.length, 9);
    assert.equal(IShellLinkW.allMethods.length, 21);
    assert.equal(IShellLinkW.methodIndex.get("SetPath"), 20);
    assert.equal(IStream.allMethods.length, 14);
    assert.equal(IStream.methodIndex.get("Clone"), 13);
    assert.equal(IClassFactory.methodIndex.get("LockServer"), 4);
  });

  it("keeps base and own methods", () => {
    assert.equal(IPersistFile.base, IPersist);
    assert.equal(IPersist.base, IUnknown);
    assert.equal(IUnknown.base, null);
    assert.deepEqual(
      IPersist.methods.map((method) => method.name),
      ["GetClassID"],
    );
    assert.equal(IShellLinkW.iid.toString(), "{000214F9-0000-0000-C000-000000000046}");
  });

  it("generates a pointer class hierarchy", () => {
    assert.equal(IPersistFile.Pointer.name, "POINTER(IPersistFile)");
    assert.ok(IPersistFile.Pointer.prototype instanceof IPersist.Pointer);
    assert.ok(IPersist.Pointer.prototype instanceof IUnknown.Pointer);
    assert.ok(IUnknown.Pointer.prototype instanceof COMPointer);
    const prototype: Record<string, unknown> = IPersistFile.Pointer.prototype as never;
    assert.equal(typeof prototype.Save, "function");
    assert.equal(typeof prototype.GetClassID, "function");
    assert.equal(Object.hasOwn(IPersistFile.Pointer.prototype, "GetClassID"), false);
  });

  it("does not shadow COMPointer members with raw IUnknown slots", () => {
    assert.equal(IUnknown.Pointer.prototype.QueryInterface, COMPointer.prototype.QueryInterface);
    assert.equal(IUnknown.Pointer.prototype.Release, COMPointer.prototype.Release);
  });

  it("validates definitions", () => {
    assert.throws(() => COMInterface({ name: "", iid: "{00000000-0000-0000-0000-000000000001}" }), TypeError);
    assert.throws(() => COMInterface({ name: "IBad", iid: "not a guid" }), SyntaxError);
    assert.throws(() => COMInterface({ name: "IDup", iid: "{00000000-0000-0000-0000-000000000001}", extends: IUnknown, methods: [STDMETHOD(HRESULT, "A"), STDMETHOD(HRESULT, "A")] }), TypeError);
    assert.throws(() => STDMETHOD(HRESULT, ""), TypeError);
  });

  it("copies the IID and freezes the definition", () => {
    const iid = "{11111111-2222-3333-4444-555555555555}";
    const IFoo = COMInterface({ name: "IFoo", iid, extends: IUnknown, methods: [STDMETHOD(HRESULT, "Foo", [ctypes.c_int32])] });
    assert.ok(Object.isFrozen(IFoo));
    assert.ok(Object.isFrozen(IFoo.methods));
    assert.equal(IFoo.iid.toString(), iid);
    assert.deepEqual(IFoo.methods[0].argtypes, [ctypes.c_int32]);
  });

  it("rejects NULL addresses", () => {
    assert.throws(() => COMPointer.fromAddress(0n, IUnknown), TypeError);
  });
});

describe("ABI layouts", () => {
  it("STATSTG matches the Windows layout", () => {
    assert.equal(ctypes.sizeof(STATSTG), ctypes.sizeof(ctypes.c_void_p) === 8 ? 80 : 72);
  });
});
