import assert from "node:assert/strict";
import { describe, it } from "node:test";
import * as ctypes from "node-ctypes";

import {
  APTTYPE_MAINSTA,
  APTTYPE_STA,
  CLSCTX_INPROC_SERVER,
  CLSID_ShellLink,
  COINIT_MULTITHREADED,
  COMError,
  COMPointer,
  CoCreateInstance,
  CoCreateInstanceEx,
  CoGetClassObject,
  CoInitializeEx,
  CoTaskMemAlloc,
  CoTaskMemFree,
  CoTaskMemRealloc,
  CoUninitialize,
  GUID,
  IClassFactory,
  IShellLinkW,
  IUnknown,
  REGDB_E_CLASSNOTREG,
  CO_E_CLASSSTRING,
  RPC_E_CHANGED_MODE,
  S_FALSE,
  S_OK,
  apartment,
  getApartmentType,
  getLivePointerCount,
} from "@node-ctypes/comtypes";

const skip = process.platform !== "win32" && "COM requires Windows";

function hresultOf(code: number) {
  return (error: unknown) => error instanceof COMError && error.hresult === code;
}

describe("apartment", { skip }, () => {
  it("initializes an STA by default and nests", () => {
    assert.equal(getApartmentType(), null);
    {
      using outer = apartment();
      assert.equal(outer.hresult, S_OK);
      const type = getApartmentType();
      assert.ok(type && (type.type === APTTYPE_STA || type.type === APTTYPE_MAINSTA));
      {
        using inner = apartment();
        assert.equal(inner.hresult, S_FALSE);
      }
      assert.notEqual(getApartmentType(), null);
    }
    assert.equal(getApartmentType(), null);
  });

  it("rejects a different concurrency model", () => {
    using _ = apartment();
    assert.throws(() => CoInitializeEx(COINIT_MULTITHREADED), hresultOf(RPC_E_CHANGED_MODE));
  });

  it("releases live pointers on the last CoUninitialize", async () => {
    CoInitializeEx();
    const before = getLivePointerCount();
    const link = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
    const unknown = link.QueryInterface(IUnknown);
    assert.equal(getLivePointerCount(), before + 2);
    CoUninitialize();
    assert.equal(getLivePointerCount(), 0);
    assert.ok(link.released && unknown.released);
    await assert.rejects(link.SetPath("x"), /released/);
    assert.throws(() => link.address, /released/);
    link[Symbol.dispose]();
  });
});

describe("object creation", { skip }, () => {
  it("CoCreateInstance defaults to IUnknown", () => {
    using _ = apartment();
    using unknown = CoCreateInstance(CLSID_ShellLink);
    assert.equal(unknown.interface, IUnknown);
    using link = unknown.QueryInterface(IShellLinkW);
    assert.ok(link.isSameObject(unknown));
  });

  it("CoCreateInstance accepts GUID strings and options", async () => {
    using _ = apartment();
    using link = CoCreateInstance(CLSID_ShellLink.toString(), IShellLinkW, { clsctx: CLSCTX_INPROC_SERVER });
    assert.equal(await link.SetPath("C:\\Windows"), S_OK);
  });

  it("reports unregistered classes", () => {
    using _ = apartment();
    assert.throws(() => CoCreateInstance(GUID.createNew(), IUnknown), hresultOf(REGDB_E_CLASSNOTREG));
  });

  it("CoCreateInstanceEx creates locally", async () => {
    using _ = apartment();
    using link = CoCreateInstanceEx(CLSID_ShellLink, IShellLinkW);
    assert.equal(await link.SetDescription("ex"), S_OK);
  });

  it("CoGetClassObject returns an IClassFactory", async () => {
    using _ = apartment();
    using factory = CoGetClassObject(CLSID_ShellLink, IClassFactory, { clsctx: CLSCTX_INPROC_SERVER });
    const out = new ctypes.c_void_p();
    await factory.CreateInstance(null, IShellLinkW.iid, ctypes.byref(out));
    using link = COMPointer.fromAddress(out.value!, IShellLinkW);
    assert.equal(await link.SetArguments("--flag"), S_OK);
  });
});

describe("GUID with COM", { skip }, () => {
  it("creates unique GUIDs", () => {
    const a = GUID.createNew();
    const b = GUID.createNew();
    assert.ok(!a.isNull());
    assert.ok(!a.equals(b));
    assert.ok(new GUID(a.toString()).equals(a));
  });

  it("maps ProgIDs and CLSIDs", () => {
    const clsid = GUID.fromProgID("Scripting.Dictionary");
    assert.equal(clsid.toString(), "{EE09B103-97E0-11CF-978F-00A02463E06F}");
    assert.equal(clsid.asProgID(), "Scripting.Dictionary");
    assert.throws(() => GUID.fromProgID("NodeComtypes.Does.Not.Exist"), hresultOf(CO_E_CLASSSTRING));
  });
});

describe("task memory", { skip }, () => {
  it("allocates, reallocates and frees", () => {
    const pointer = CoTaskMemAlloc(16);
    ctypes.memset(pointer, 0xab, 16);
    const grown = CoTaskMemRealloc(pointer, 64);
    assert.equal(ctypes.readValue(grown, ctypes.c_uint8, 15), 0xab);
    CoTaskMemFree(grown);
    CoTaskMemFree(null);
  });
});
