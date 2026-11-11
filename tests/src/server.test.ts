import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";
import { describe, it } from "node:test";
import { setTimeout as delay } from "node:timers/promises";
import { fileURLToPath } from "node:url";
import * as ctypes from "node-ctypes";

import {
  BSTR,
  CLSCTX_INPROC_SERVER,
  CLSCTX_LOCAL_SERVER,
  COMError,
  COMInterface,
  COMMETHOD,
  COMPointer,
  CoCreateInstance,
  CreateObject,
  DISP_E_EXCEPTION,
  DISP_E_MEMBERNOTFOUND,
  DISP_E_UNKNOWNNAME,
  Dispatch,
  E_NOINTERFACE,
  E_NOTIMPL,
  GetEvents,
  GetObject,
  HRESULT,
  IDispatch,
  IEnumVARIANT,
  ISupportErrorInfo,
  IUnknown,
  PumpEvents,
  REGDB_E_CLASSNOTREG,
  SELF,
  STDMETHOD,
  ShowEvents,
  VARIANT,
  apartment,
  dispatchPointer,
  dispid,
  getLivePointerCount,
  named,
  startMessagePump,
  waitFor,
} from "@node-ctypes/comtypes";
import { COMObject, ConnectionPoint, VariantEnumerator, getLiveObjectCount, registerClassObject, registerServer, unregisterServer } from "@node-ctypes/comtypes/server";

const skip = process.platform !== "win32" && "COM requires Windows";
const { c_int32, c_uint32 } = ctypes;

const ICalc = COMInterface({
  name: "ICalc",
  iid: "{6B29FC40-CA47-1067-B31D-00DD010662DA}",
  extends: IDispatch,
  methods: [
    COMMETHOD([dispid(1)], HRESULT, "Add", [["in"], c_int32, "a"], [["in", "optional"], c_int32, "b", 1], [["out", "retval"], c_int32, "sum"]),
    COMMETHOD([dispid(2), "propget"], HRESULT, "Name", [["out", "retval"], BSTR, "name"]),
    COMMETHOD([dispid(2), "propput"], HRESULT, "Name", [["in"], BSTR, "name"]),
    COMMETHOD([dispid(3)], HRESULT, "Echo", [["in"], VARIANT, "value"], [["out", "retval"], VARIANT, "result"]),
    COMMETHOD([dispid(4)], HRESULT, "Fail", [["in"], BSTR, "message"]),
    COMMETHOD([dispid(5)], HRESULT, "Self", [["out", "retval"], SELF, "self"]),
    COMMETHOD([dispid(6)], HRESULT, "Twice", [["in", "out"], c_int32, "number"], [["in", "out"], BSTR, "text"], [["lcid"], c_uint32, "lcid"]),
    COMMETHOD([dispid(7)], HRESULT, "Missing"),
    STDMETHOD(HRESULT, "Raw", [c_int32]),
  ],
});

class Calc extends COMObject {
  static _com_interfaces_ = [ICalc];
  static _reg_clsid_ = "{6B29FC41-CA47-1067-B31D-00DD010662DA}";
  static _com_default_member_ = "Name";
  Name = "calc";
  released = 0;
  lastRaw = 0;
  Add(a: number, b: number) {
    return a + b;
  }
  Echo(value: unknown) {
    return value;
  }
  Fail(message: string): never {
    throw new Error(message);
  }
  Self() {
    return this;
  }
  Twice(number: number, text: string) {
    return [number * 2, text + text];
  }
  Raw(value: number) {
    this.lastRaw = value;
    return value < 0 ? E_NOTIMPL : 0;
  }
  *[Symbol.iterator]() {
    yield* [1, "two", 3];
  }
  _final_release_() {
    this.released++;
  }
}

describe("COMObject (vtable)", { skip }, () => {
  it("implements COMMETHODs: arguments, results, properties and SELF", async () => {
    using _ = apartment();
    const calc = new Calc();
    assert.equal(calc.refCount, 0);
    {
      using pointer = calc.QueryInterface(ICalc);
      assert.equal(calc.refCount, 1);
      assert.equal(await pointer.Add(2, 3), 5);
      assert.equal(await pointer.Add(41), 42, "optional argument with default");
      assert.equal(pointer.Name, "calc");
      pointer.Name = "renamed";
      assert.equal(calc.Name, "renamed");
      assert.deepEqual(await pointer.Echo([1, "x", null]), [1, "x", null]);
      assert.equal(((await pointer.Echo(new Date(2020, 1, 3))) as Date).getTime(), new Date(2020, 1, 3).getTime());
      assert.deepEqual(await pointer.Twice(4, "ab"), [8, "abab"]);
      using self = (await pointer.Self())!;
      assert.ok(self.isSameObject(pointer));
      assert.equal(await pointer.Raw(7), 0);
      assert.equal(calc.lastRaw, 7);
      await assert.rejects(pointer.Raw(-1), (error: unknown) => error instanceof COMError && error.hresult === E_NOTIMPL);
      await assert.rejects(pointer.Missing(), (error: unknown) => error instanceof COMError && error.hresult === E_NOTIMPL);
    }
    assert.equal(calc.refCount, 0);
    assert.equal(calc.released, 1, "_final_release_ runs at zero references");
  });

  it("reports exceptions through IErrorInfo", async () => {
    using _ = apartment();
    using pointer = new Calc().QueryInterface(ICalc);
    await assert.rejects(
      pointer.Fail("boom"),
      (error: unknown) => error instanceof COMError && error.hresult < 0 && error.method === "ICalc::Fail" && error.details?.description === "boom" && error.details.source === "Calc" && Boolean(error.details.guid?.equals(ICalc.iid)),
    );
    using support = pointer.QueryInterface(ISupportErrorInfo);
    assert.equal(await support.InterfaceSupportsErrorInfo(ICalc.iid), 0);
    assert.equal(await support.InterfaceSupportsErrorInfo("{00000000-0000-0000-0000-000000000001}"), 1);
  });

  it("answers QueryInterface for its interfaces and their bases only", () => {
    using _ = apartment();
    const calc = new Calc();
    using unknown = calc.QueryInterface(IUnknown);
    using dispatch = unknown.QueryInterface(IDispatch);
    using again = dispatch.QueryInterface(IUnknown);
    assert.equal(unknown.address, again.address, "IUnknown identity is stable");
    assert.equal(unknown.tryQueryInterface(IEnumVARIANT), null);
    assert.throws(
      () => calc.QueryInterface(IEnumVARIANT),
      (error: unknown) => error instanceof COMError && error.hresult === E_NOINTERFACE,
    );
    assert.ok(calc.supportsInterface(ISupportErrorInfo.iid));
    assert.deepEqual(
      calc.interfaces.map((iface) => iface.name),
      ["ICalc", "ISupportErrorInfo"],
    );
  });

  it("is passed to native code and released by it", async () => {
    using _ = apartment();
    const calc = new Calc();
    {
      // Created on the home thread: the dictionary holds the object itself (no proxy).
      using dict = await CreateObject("Scripting.Dictionary", { home: true });
      await dict.Add("calc", calc);
      assert.equal(calc.refCount, 1, "the dictionary holds a reference");
      using stored = dict.Item("calc");
      assert.equal(await stored.Add(1, 2), 3);
      using variant = new VARIANT(calc);
      assert.equal(calc.refCount, 3);
      await dict.RemoveAll();
    }
    assert.equal(calc.refCount, 0);
  });

  it("does not leak objects or pointers", async () => {
    using _ = apartment();
    const objects = getLiveObjectCount();
    const pointers = getLivePointerCount();
    for (let i = 0; i < 300; i++) {
      using pointer = new Calc().QueryInterface(ICalc);
      using self = (await pointer.Self())!;
      assert.equal(await self.Add(i, 1), i + 1);
      using dynamic = Dispatch(pointer);
      assert.equal(await dynamic.Add(i, 2), i + 2);
    }
    assert.equal(getLiveObjectCount(), objects);
    assert.equal(getLivePointerCount(), pointers);
  });

  it("supports thread-safe callbacks (free threading) called from the JavaScript thread", async () => {
    class FreeCalc extends Calc {
      static _threading_ = "free" as const;
    }
    using _ = apartment();
    using pointer = new FreeCalc().QueryInterface(ICalc);
    assert.equal(await pointer.Add(20, 22), 42);
    assert.equal(pointer.Name, "calc");
  });
});

describe("COMObject (IDispatch)", { skip }, () => {
  it("serves late-bound calls from dual interfaces and JavaScript members", async () => {
    using _ = apartment();
    const calc = new Calc();
    using dynamic = Dispatch(calc);
    assert.equal(await dynamic.Add(4, 5), 9);
    assert.equal(dynamic.Name, "calc");
    dynamic.Name = "late";
    assert.equal(calc.Name, "late");
    // Without type information the default member is called as a method.
    assert.equal(await dynamic(), "late", "DISPID_VALUE is the default member");
    assert.equal(dynamic.released, 0, "JavaScript fields are members too");
    assert.deepEqual([...dynamic], [1, "two", 3], "iterables answer _NewEnum");
    assert.equal(await dynamic.Add(named({ a: 10, b: 5 })), 15, "named arguments of COMMETHOD specs");
  });

  it("returns DISP_E_* errors and EXCEPINFO", async () => {
    using _ = apartment();
    using dynamic = Dispatch(new Calc());
    assert.throws(
      () => dynamic.Nope,
      (error: unknown) => error instanceof COMError && error.hresult === DISP_E_UNKNOWNNAME,
    );
    await assert.rejects(dynamic.Fail("late boom"), (error: unknown) => error instanceof COMError && error.hresult === DISP_E_EXCEPTION && error.details?.description === "late boom" && error.details.source === "Calc");
    class Plain extends COMObject {
      static _com_interfaces_ = [IDispatch];
      value = 1;
      method() {
        return "called";
      }
    }
    using plain = Dispatch(new Plain());
    assert.equal(plain.value, 1);
    assert.equal(await plain.method(), "called");
    await assert.rejects(plain(), (error: unknown) => error instanceof COMError && error.hresult === DISP_E_MEMBERNOTFOUND);
  });

  it("enumerates with VariantEnumerator", async () => {
    using _ = apartment();
    using enumerator = new VariantEnumerator(["a", 2, true]).QueryInterface(IEnumVARIANT);
    using item = new VARIANT();
    const fetched = new ctypes.c_uint32();
    assert.equal(await enumerator.Next(1, item, ctypes.byref(fetched)), 0);
    assert.equal(item.value, "a");
    item.clear();
    assert.equal(await enumerator.Skip(1), 0);
    using clone = (await enumerator.Clone())!;
    assert.equal(await clone.Next(1, item, ctypes.byref(fetched)), 0);
    assert.equal(item.value, true);
    item.clear();
    assert.equal(await clone.Next(1, item, ctypes.byref(fetched)), 1, "S_FALSE at the end");
    assert.equal(fetched.value, 0);
    await enumerator.Reset();
    assert.equal(await enumerator.Next(1, item, ctypes.byref(fetched)), 0);
    assert.equal(item.value, "a");
  });
});

const ITickEvents = COMInterface({
  name: "ITickEvents",
  iid: "{9E3F1A10-1111-4C4C-9A9A-0123456789AB}",
  extends: IDispatch,
  methods: [COMMETHOD([dispid(1)], HRESULT, "Tick", [["in"], VARIANT, "count"]), COMMETHOD([dispid(2)], HRESULT, "Done", [["in"], BSTR, "message"])],
});
const IVtableEvents = COMInterface({
  name: "IVtableEvents",
  iid: "{9E3F1A10-3333-4C4C-9A9A-0123456789AB}",
  extends: IUnknown,
  methods: [COMMETHOD([], HRESULT, "Changed", [["in"], BSTR, "what"], [["in"], c_int32, "count"])],
});
const ITicker = COMInterface({ name: "ITicker", iid: "{9E3F1A10-2222-4C4C-9A9A-0123456789AB}", extends: IDispatch });

class Ticker extends COMObject {
  static _com_interfaces_ = [ITicker];
  static _outgoing_interfaces_ = [ITickEvents, IVtableEvents];
  run(times: number) {
    for (let i = 1; i <= times; i++) this.fireEvent("Tick", i);
    this.fireEvent("Done", "finished");
  }
}

describe("events", { skip }, () => {
  it("connects JavaScript sinks to a JavaScript source", () => {
    using _ = apartment();
    const ticker = new Ticker();
    using source = ticker.QueryInterface(ITicker);
    const seen: unknown[] = [];
    {
      using connection = GetEvents(source, { Tick: (count: number) => seen.push(count), Done: (message: string) => seen.push(message) });
      assert.equal(connection.interface, ITickEvents, "default source interface through IProvideClassInfo2");
      assert.equal(ticker.connectedSinks(ITickEvents).length, 1);
      ticker.run(2);
    }
    assert.equal(ticker.connectedSinks(ITickEvents).length, 0, "disposal unadvises");
    ticker.run(5);
    assert.deepEqual(seen, [1, 2, "finished"]);
  });

  it("delivers vtable events and logs with ShowEvents", () => {
    using _ = apartment();
    const ticker = new Ticker();
    const changes: unknown[] = [];
    using connection = GetEvents(ticker, { Changed: (what: string, count: number) => changes.push([what, count]) }, { interface: IVtableEvents });
    ticker.fireEvent("Changed", "x", 3);
    assert.deepEqual(changes, [["x", 3]]);
    const lines: string[] = [];
    using shown = ShowEvents(ticker, { log: (line) => lines.push(line) });
    ticker.run(1);
    assert.deepEqual(lines, ["Event Tick(1)", "Event Done('finished')"]);
    void connection;
    void shown;
  });

  it("exposes connection points", () => {
    using _ = apartment();
    const ticker = new Ticker();
    const point = ticker.FindConnectionPoint(ITickEvents.iid);
    assert.ok(point instanceof ConnectionPoint);
    assert.equal(ticker.FindConnectionPoint(ITickEvents.iid), point);
    assert.throws(() => ticker.FindConnectionPoint(ICalc.iid), COMError);
    assert.throws(() => point.Unadvise(99), COMError);
  });

  it("receives WMI events through the message pump", async () => {
    using _ = apartment();
    using pump = startMessagePump();
    using wmi = await GetObject("winmgmts:");
    using sink = await CreateObject("WbemScripting.SWbemSink");
    const names: string[] = [];
    let completed = false;
    using connection = GetEvents(sink, {
      OnObjectReady(object: { Name: string }) {
        names.push(object.Name);
      },
      OnCompleted() {
        completed = true;
      },
    });
    assert.equal(connection.interface.name, "ISWbemSinkEvents");
    wmi.ExecQueryAsync(sink, `SELECT Name FROM Win32_Process WHERE ProcessId = ${process.pid}`);
    assert.equal(await waitFor(() => completed, { timeout: 20_000 }), true);
    assert.deepEqual(names, ["node.exe"]);
    void pump;
  });

  it("PumpEvents blocks for the given time", () => {
    using _ = apartment();
    const started = performance.now();
    PumpEvents(50);
    assert.ok(performance.now() - started >= 40);
  });
});

describe("class factories and servers", { skip }, () => {
  it("registers an in-process class object", async () => {
    using _ = apartment();
    {
      using registration = registerClassObject(Calc, { clsctx: CLSCTX_INPROC_SERVER });
      using pointer = CoCreateInstance(Calc._reg_clsid_, ICalc, { clsctx: CLSCTX_INPROC_SERVER });
      assert.equal(await pointer.Add(1, 2), 3);
      using dynamic = await CreateObject(Calc._reg_clsid_, { clsctx: CLSCTX_INPROC_SERVER });
      assert.equal(dynamic.Name, "calc");
      void registration;
    }
    assert.throws(
      () => CoCreateInstance(Calc._reg_clsid_, ICalc, { clsctx: CLSCTX_INPROC_SERVER }),
      (error: unknown) => error instanceof COMError && error.hresult === REGDB_E_CLASSNOTREG,
    );
  });

  it("runs a LocalServer32 in another Node.js process (per-user registration)", async () => {
    const clsid = `{${randomUUID().toUpperCase()}}`;
    class Remote extends COMObject {
      static _com_interfaces_ = [IDispatch];
      static _reg_clsid_ = clsid;
      static _reg_desc_ = "node-comtypes test LocalServer32";
    }
    registerServer(Remote, { script: fileURLToPath(new URL("../fixtures/localserver.ts", import.meta.url)), args: [clsid] });
    let pid = 0;
    try {
      using _ = apartment();
      using calculator = await CreateObject(clsid, { clsctx: CLSCTX_LOCAL_SERVER });
      pid = calculator.Pid;
      assert.notEqual(pid, process.pid);
      assert.equal(await calculator.Add(2, 40), 42);
      assert.deepEqual(await calculator.Echo([1, "x", true]), [1, "x", true]);
      assert.equal(calculator.Name, "calculator");
      await assert.rejects(calculator.Fail("remote"), (error: unknown) => error instanceof COMError && error.hresult === DISP_E_EXCEPTION && error.details?.description === "remote");
    } finally {
      unregisterServer(Remote);
    }
    const alive = () => {
      try {
        process.kill(pid, 0);
        return true;
      } catch {
        return false;
      }
    };
    for (let attempt = 0; attempt < 100 && alive(); attempt++) await delay(100);
    assert.equal(alive(), false, "the server exits once its objects are released");
  });

  it("exposes the dispatch pointer of a COMObject", () => {
    using _ = apartment();
    using dynamic = Dispatch(new Calc());
    assert.ok(dispatchPointer(dynamic) instanceof COMPointer);
  });
});
