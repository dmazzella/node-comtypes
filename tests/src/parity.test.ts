import assert from "node:assert/strict";
import { describe, it } from "node:test";
import * as ctypes from "node-ctypes";

import {
  COMError,
  COMInterface,
  COMMETHOD,
  CoCreateInstance,
  Constants,
  CreateObject,
  DISP_E_EXCEPTION,
  DVASPECT_CONTENT,
  Dispatch,
  E_ACCESSDENIED,
  E_NOINTERFACE,
  GUID,
  HRESULT,
  IDispatch,
  IPersistPropertyBag,
  IPropertyBag,
  IServiceProvider,
  IShellLinkA,
  IShellLinkW,
  IUnknown,
  IViewObject,
  IViewObject2,
  S_FALSE,
  SIZEL,
  STDMETHOD,
  ShellLink,
  apartment,
  queryService,
} from "@node-ctypes/comtypes";
import { readMemory } from "@node-ctypes/comtypes/low-level";
import { COMObject, ReturnHRESULT, queryObject } from "@node-ctypes/comtypes/server";

const skip = process.platform !== "win32" && "COM requires Windows";
const CLSID_WebBrowser = new GUID("{8856F961-340A-11D0-A96B-00C04FD705A2}");

describe("Constants", { skip }, () => {
  it("exposes enum members, enums and module constants of the type library", async () => {
    using _ = apartment();
    using fso = await CreateObject("Scripting.FileSystemObject");
    const c = Constants(fso);
    assert.equal(c.TemporaryFolder, 2);
    assert.equal(c.SpecialFolderConst.TemporaryFolder, 2);
    assert.equal(c.ForAppending, 8);
    assert.ok("TristateTrue" in c);
    assert.throws(() => c.NoSuchConstant, /Scripting has no constant or enum NoSuchConstant/);
    assert.equal(c.then, undefined, "not thenable");
    using dynamic = await CreateObject("Scripting.Dictionary", { dynamic: true });
    assert.equal(Constants(dynamic), c, "cached per type library");
    using temp = await fso.GetSpecialFolder(c.TemporaryFolder);
    assert.ok(temp.Path.length > 0);
  });
});

describe("comtypes.persist", { skip }, () => {
  class Bag extends COMObject {
    static _com_interfaces_ = [IPropertyBag];
    readonly values = new Map<string, unknown>([
      ["ExtentX", 5000],
      ["ExtentY", 3000],
    ]);
    readonly reads: string[] = [];
    Read(name: string, initial: unknown) {
      this.reads.push(name);
      return this.values.has(name) ? this.values.get(name) : initial;
    }
    Write(name: string, value: unknown) {
      this.values.set(name, value);
    }
  }

  it("IPersistPropertyBag loads from and saves to a JavaScript IPropertyBag (WebBrowser control)", async () => {
    using _ = apartment();
    using browser = CoCreateInstance(CLSID_WebBrowser, IUnknown);
    using persist = browser.QueryInterface(IPersistPropertyBag);
    assert.ok((await persist.GetClassID()).equals(CLSID_WebBrowser));
    const bag = new Bag();
    await persist.Load(bag, null);
    assert.ok(bag.reads.includes("ExtentX") && bag.reads.includes("Height"), "the control reads its properties");
    bag.values.clear();
    await persist.Save(bag, 1, 1);
    assert.equal(bag.values.get("ExtentX"), 5000, "Write receives the VARIANT* value");
    assert.equal(bag.values.get("ExtentY"), 3000);
  });

  it("IPropertyBag calls round-trip through a JavaScript implementation", async () => {
    using _ = apartment();
    using bag = new Bag().QueryInterface(IPropertyBag);
    await bag.Write("Name", "value");
    assert.equal(await bag.Read("Name", undefined, null), "value");
    assert.equal(await bag.Read("Missing", 42, null), 42, "[in, out]: the initial value");
  });
});

describe("IViewObject2", { skip }, () => {
  it("returns the extent as a SIZEL", async () => {
    using _ = apartment();
    using browser = CoCreateInstance(CLSID_WebBrowser, IUnknown);
    using view = browser.QueryInterface(IViewObject2);
    const extent = await view.GetExtent(DVASPECT_CONTENT, -1, null);
    assert.ok(extent instanceof SIZEL);
    // Same result through a raw definition of the slot.
    const Raw = COMInterface({ name: "IViewObject2Raw", iid: IViewObject2.iid, extends: IViewObject, register: false, methods: [STDMETHOD(HRESULT, "GetExtent", [ctypes.c_uint32, ctypes.c_int32, ctypes.c_void_p, ctypes.c_void_p])] });
    using raw = browser.QueryInterface(Raw);
    const size = Buffer.alloc(8, 0x7f);
    await raw.GetExtent(DVASPECT_CONTENT, -1, null, size);
    assert.deepEqual([extent.cx, extent.cy], [size.readInt32LE(0), size.readInt32LE(4)]);
  });
});

describe("IServiceProvider", { skip }, () => {
  const SID_Test = new GUID("{0E8C5A20-0000-4000-8000-000000000001}");
  class Service extends COMObject {
    static _com_interfaces_ = [IDispatch];
    Answer() {
      return 42;
    }
  }
  class Provider extends COMObject {
    static _com_interfaces_ = [IServiceProvider];
    readonly service = new Service();
    QueryService(guidService: bigint, riid: bigint, ppv: bigint): number {
      ctypes.writeValue(ppv, ctypes.c_void_p, 0n);
      if (!new GUID(new Uint8Array(readMemory(guidService, 16))).equals(SID_Test)) return E_NOINTERFACE;
      const address = queryObject(this.service, new GUID(new Uint8Array(readMemory(riid, 16))));
      if (address === null) return E_NOINTERFACE;
      ctypes.writeValue(ppv, ctypes.c_void_p, address);
      return 0;
    }
  }

  it("queryService returns the service as the requested interface", async () => {
    using _ = apartment();
    const provider = new Provider();
    using service = queryService(provider, SID_Test, IDispatch);
    assert.equal(await Dispatch(service).Answer(), 42);
    assert.throws(
      () => queryService(provider, "{0E8C5A20-0000-4000-8000-000000000002}", IDispatch),
      (error: unknown) => error instanceof COMError && error.hresult === E_NOINTERFACE,
    );
    // The WebBrowser control provides services only once sited.
    using browser = CoCreateInstance(CLSID_WebBrowser, IUnknown);
    assert.ok(browser.tryQueryInterface(IServiceProvider));
  });
});

describe("comtypes.shelllink", { skip }, () => {
  it("ShellLink is a coclass with IShellLinkW and IShellLinkA", async () => {
    using _ = apartment();
    using link = await CreateObject(ShellLink);
    assert.equal(link.interface, IShellLinkW);
    using ansi = link.QueryInterface(IShellLinkA);
    await ansi.SetPath("C:\\Windows\\notepad.exe");
    const wide = ctypes.create_unicode_buffer(260);
    await link.GetPath(wide, 260, null, 0);
    assert.equal(wide.value.toLowerCase(), "c:\\windows\\notepad.exe");
    const narrow = ctypes.create_string_buffer(260);
    await ansi.GetPath(narrow, 260, null, 0);
    assert.equal(narrow.value.toString().toLowerCase(), "c:\\windows\\notepad.exe");
  });
});

describe("ReturnHRESULT", { skip }, () => {
  const IChecker = COMInterface({
    name: "IChecker",
    iid: "{0E8C5A20-0000-4000-8000-000000000010}",
    extends: IDispatch,
    methods: [COMMETHOD([], HRESULT, "Check", [["in"], ctypes.c_int32, "value"])],
  });
  class Checker extends COMObject {
    static _com_interfaces_ = [IChecker];
    static errors: unknown[] = [];
    static onError = (error: unknown) => void Checker.errors.push(error);
    Check(value: number) {
      if (value === 0) throw new ReturnHRESULT(S_FALSE);
      if (value < 0) throw new ReturnHRESULT(E_ACCESSDENIED, "negative values are not allowed");
      if (value > 100) throw new Error("too large");
    }
  }

  it("returns success codes and failures without onError", async () => {
    using _ = apartment();
    Checker.errors = [];
    using checker = new Checker().QueryInterface(IChecker);
    assert.equal(await checker.Check(1), 0);
    assert.equal(await checker.Check(0), S_FALSE);
    await assert.rejects(checker.Check(-1), (error: unknown) => error instanceof COMError && error.hresult === E_ACCESSDENIED && error.details?.description === "negative values are not allowed");
    assert.deepEqual(Checker.errors, []);
    await assert.rejects(checker.Check(101), COMError);
    assert.equal(Checker.errors.length, 1, "other exceptions still reach onError");
  });

  it("fills EXCEPINFO through IDispatch", async () => {
    using _ = apartment();
    Checker.errors = [];
    using dynamic = Dispatch(new Checker());
    assert.equal(await dynamic.Check(0), undefined);
    await assert.rejects(dynamic.Check(-1), (error: unknown) => error instanceof COMError && error.hresult === DISP_E_EXCEPTION && error.details?.scode === E_ACCESSDENIED && error.details.description === "negative values are not allowed");
    assert.deepEqual(Checker.errors, []);
  });
});
