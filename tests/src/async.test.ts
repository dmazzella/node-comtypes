import assert from "node:assert/strict";
import { describe, it } from "node:test";
import * as ctypes from "node-ctypes";

import {
  BSTR,
  CLSCTX_INPROC_SERVER,
  COINIT_MULTITHREADED,
  COMPointer,
  COMInterface,
  COMMETHOD,
  CreateObject,
  DISPMETHOD,
  Dispatch,
  GUID,
  GetObject,
  HRESULT,
  IDispatch,
  IUnknown,
  SAFEARRAYBOUND,
  VARIANT,
  VT_DISPATCH,
  apartment,
  destroySafeArray,
  dispatchPointer,
  dispid,
  finalize,
  getIDsOfNames,
  getLivePointerCount,
  invokeMethod,
  isDispatch,
  readSafeArray,
  run,
  stats,
  sync,
} from "@node-ctypes/comtypes";
import type { InterfaceDef } from "@node-ctypes/comtypes";
import { executorCall, ole32, oleaut32 } from "@node-ctypes/comtypes/low-level";
import { COMObject } from "@node-ctypes/comtypes/server";

const skip = process.platform !== "win32" && "COM requires Windows";

// msxml6.idl: IXMLDOMDocument (dual) createElement, IXMLDOMNode nodeName.
const IXMLDOMNode = COMInterface({
  name: "IXMLDOMNode",
  iid: "{2933BF80-7B36-11D2-B20E-00C04F983E60}",
  extends: IDispatch,
  methods: [COMMETHOD([dispid(2), "propget"], HRESULT, "nodeName", [["out", "retval"], BSTR, "name"])],
});

/** Executor calls settled while `action` runs. */
async function executorCalls(action: () => Promise<unknown>): Promise<number> {
  const before = stats();
  await action();
  const after = stats();
  return after.generation === before.generation ? after.completed - before.completed : after.completed;
}

describe("asynchronous methods", { skip }, () => {
  it("return promises; properties and the sync() view stay synchronous", async () => {
    using _ = apartment();
    using dict = (await CreateObject("Scripting.Dictionary")) as any;
    const added = dict.Add("a", 1);
    assert.ok(added instanceof Promise, "methods return promises");
    await added;
    assert.equal(dict.Count, 1, "properties are synchronous");
    assert.equal(dict.Item("a"), 1, "parameterized properties are synchronous");
    const view = sync(dict) as any;
    assert.equal(await view.Exists("a"), true, "the sync() view calls methods synchronously");
    assert.equal(sync(dict), view, "one view per pointer");
    assert.equal(view.address, dict.address);
  });

  it("the sync() view shares the reference of its pointer", async () => {
    using _ = apartment();
    const before = getLivePointerCount();
    const dict = await CreateObject("Scripting.Dictionary", { interface: IUnknown });
    const view = sync(dict);
    assert.equal(getLivePointerCount(), before + 1, "no reference of its own");
    view.Release();
    assert.equal(dict.released, true, "released through the view");
    assert.equal(getLivePointerCount(), before);
  });

  it("call agile objects of a single-threaded home on the executor, at the same address", async () => {
    using _ = apartment();
    using document = (await CreateObject("MSXML2.FreeThreadedDOMDocument.6.0")) as any;
    let element: any;
    const calls = await executorCalls(async () => {
      element = await document.createElement("item");
    });
    using created = element;
    assert.ok(calls > 0, "ran on the executor");
    // Interface results are usable on the home thread.
    assert.equal(sync(dispatchPointer(Dispatch(created))).QueryInterface(IXMLDOMNode).nodeName, "item");
  });

  it("call objects created on the home thread (home: true) there", async () => {
    using _ = apartment();
    using dict = (await CreateObject("Scripting.Dictionary", { home: true })) as any;
    const calls = await executorCalls(() => dict.Add("k", "v"));
    assert.equal(calls, 0, "no executor call");
    assert.equal(dict.Item("k"), "v");
  });

  it("create objects on the executor: a single-threaded home calls them through it", async () => {
    using _ = apartment();
    // wbemdisp.dll is in-process: created on the executor, the object lives in the MTA.
    using wmi = (await GetObject("winmgmts:")) as any;
    let set: any;
    const calls = await executorCalls(async () => {
      set = await wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE ProcessId = " + process.pid);
    });
    assert.ok(calls > 0, "ran on the executor");
    // The result is a home reference (a proxy): its properties are synchronous.
    using results = set;
    assert.equal(results.Count, 1);
  });

  it("share every pointer with the executor when the home is multithreaded", async () => {
    using _ = apartment(COINIT_MULTITHREADED);
    using dict = (await CreateObject("Scripting.Dictionary")) as any;
    const calls = await executorCalls(() => dict.Add("k", 1));
    assert.ok(calls > 0, "ran on the executor");
    assert.equal(dict.Count, 1);
  });

  it("call apartment-threaded JavaScript objects on the home thread", async () => {
    const IAdder = COMInterface({
      name: "IAdder",
      iid: "{7A3C1D40-5E2B-4F11-9C8D-1B2A3C4D5E6F}",
      extends: IUnknown,
      methods: [COMMETHOD([], HRESULT, "Add", [["in"], ctypes.c_int32, "a"], [["in"], ctypes.c_int32, "b"], [["out", "retval"], ctypes.c_int32, "sum"])],
    });
    class Adder extends COMObject {
      static _com_interfaces_ = [IAdder as InterfaceDef<any>];
      Add(a: number, b: number) {
        return a + b;
      }
    }
    using _ = apartment(COINIT_MULTITHREADED);
    using adder = new Adder().QueryInterface(IAdder);
    let sum = 0;
    const calls = await executorCalls(async () => {
      sum = await adder.Add(20, 22);
    });
    assert.equal(sum, 42);
    assert.equal(calls, 0, "the callbacks ran on the home thread");
  });

  it("keep calls whose arguments cannot travel on the home thread", async () => {
    using _ = apartment();
    using document = (await CreateObject("MSXML2.FreeThreadedDOMDocument.6.0")) as any;
    using root = await document.createElement("root");
    // An agile argument travels with the call.
    const agile = await executorCalls(() => document.appendChild(root).then((child: Disposable) => child[Symbol.dispose]()));
    assert.ok(agile > 0, "agile arguments travel");
    // A Scripting.Dictionary is bound to the home apartment: the call runs there (and fails there).
    using dict = (await CreateObject("Scripting.Dictionary")) as any;
    const bound = await executorCalls(() => document.save(dict).catch(() => undefined));
    assert.equal(bound, 0, "the call ran on the home thread");
  });

  it("release after the calls in progress", async () => {
    using _ = apartment(COINIT_MULTITHREADED);
    const dict = (await CreateObject("Scripting.Dictionary")) as any;
    const pending = await dict.Add("k", 1);
    assert.equal(dict.Release(), 0, "the release follows the call on the executor");
    await pending;
    assert.equal(dict.released, true);
  });

  it("start a new executor after finalize", async () => {
    using _ = apartment();
    using document = (await CreateObject("MSXML2.FreeThreadedDOMDocument.6.0")) as any;
    using first = await document.createElement("first");
    const generation = stats().generation;
    await finalize();
    assert.equal(stats().running, false);
    using second = await document.createElement("second");
    assert.equal(stats().generation, generation + 1);
    assert.equal(sync(dispatchPointer(Dispatch(second))).QueryInterface(IXMLDOMNode).nodeName, "second");
    void first;
  });

  it("call DISPMETHOD methods through IDispatch::Invoke asynchronously", async () => {
    using _ = apartment(COINIT_MULTITHREADED);
    using dict = (await CreateObject("Scripting.Dictionary")) as any;
    await dict.Add("k", 1);
    const [exists] = getIDsOfNames(sync(dispatchPointer(Dispatch(dict))) as any, ["Exists"]);
    const IDictionaryDisp = COMInterface({
      name: "IDictionaryDisp",
      iid: "{42C642C1-97E1-11CF-978F-00A02463E06F}",
      extends: IDispatch,
      register: false,
      dispMethods: [DISPMETHOD([dispid(exists)], VARIANT, "Exists", [["in"], BSTR, "key"])],
    });
    using disp = dict.QueryInterface(IDictionaryDisp) as any;
    const result = disp.Exists("k");
    assert.ok(result instanceof Promise);
    assert.equal(await result, true);
    assert.equal(sync(disp).Exists("missing"), false);
  });
});

describe("results of the executor, handed to a single-threaded home", { skip }, () => {
  it("pointers, VARIANTs and arrays of objects get references for the home apartment", async () => {
    const { handOver } = await import("../../dist/executor/apartment.js");
    using _ = apartment();
    const clsid = GUID.fromProgID("Scripting.Dictionary");
    const createArgs = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint32, ctypes.c_void_p, ctypes.c_void_p];
    // Dictionaries created on the executor live in the multithreaded apartment.
    const createOnExecutor = async () => {
      const out = new ctypes.c_void_p();
      const { value } = await executorCall({ fn: ole32().CoCreateInstance.address, restype: ctypes.c_int32, argtypes: createArgs, args: [clsid._buffer, null, CLSCTX_INPROC_SERVER, IDispatch.iid._buffer, ctypes.byref(out)] });
      assert.equal(value, 0);
      return out.value!;
    };
    const original = await createOnExecutor();
    const pointer = new ctypes.c_void_p(original);
    const variant = new VARIANT();
    variant.vt = VT_DISPATCH;
    ctypes.writeValue(variant._buffer, ctypes.c_void_p, await createOnExecutor(), 8);
    const bound = SAFEARRAYBOUND.create({ cElements: 2, lLbound: 0 });
    const psa = oleaut32().SafeArrayCreate(VT_DISPATCH, 1, bound._buffer)!;
    const data = new ctypes.c_void_p();
    oleaut32().SafeArrayAccessData(psa, ctypes.byref(data));
    for (let index = 0; index < 2; index++) ctypes.writeValue(data.value!, ctypes.c_void_p, await createOnExecutor(), index * ctypes.sizeof(ctypes.c_void_p));
    oleaut32().SafeArrayUnaccessData(psa);
    const array = new ctypes.c_void_p(psa);

    await handOver([
      { location: ctypes.addressof(pointer), kind: "pointer", iid: IDispatch.iid },
      { location: ctypes.addressof(variant._buffer), kind: "variant" },
      { location: ctypes.addressof(array), kind: "safearray" },
    ]);

    // Proxies for the home apartment, usable on the home thread.
    assert.notEqual(pointer.value, original);
    using fromPointer = Dispatch(COMPointer.fromAddress(pointer.value!, IDispatch)) as any;
    await fromPointer.Add("k", 1);
    assert.equal(fromPointer.Count, 1);
    using fromVariant = variant.value as COMPointer;
    variant.clear();
    using owned = Dispatch(fromVariant) as any;
    await owned.Add("v", 2);
    assert.equal(owned.Item("v"), 2);
    const elements = readSafeArray(psa, VT_DISPATCH) as COMPointer[];
    destroySafeArray(psa);
    assert.equal(elements.length, 2);
    for (const element of elements) {
      using owned = element;
      using item = Dispatch(owned) as any;
      await item.Add("a", 3);
      assert.equal(item.Count, 1);
    }
  });
});

describe("asynchronous client", { skip }, () => {
  const kernel32 = process.platform === "win32" ? new ctypes.WinDLL("kernel32.dll") : null;
  const SLEEP_ARGS: ctypes.FFIType[] = [ctypes.c_uint32];
  /** Keep the executor busy for `ms` milliseconds. */
  const busy = (ms: number) => executorCall({ fn: kernel32!.func("Sleep", null, [ctypes.c_uint32]).address, restype: null, argtypes: SLEEP_ARGS, args: [ms] });

  it("create objects on the executor; home: true creates them on the home thread", async () => {
    using _ = apartment();
    const before = stats();
    using remote = (await CreateObject("Scripting.Dictionary")) as any;
    const after = stats();
    assert.ok(after.generation !== before.generation || after.completed > before.completed, "created on the executor");
    using local = (await CreateObject("Scripting.Dictionary", { home: true })) as any;
    // Both are typed (IDictionary) and work the same way.
    for (const dict of [remote, local]) {
      assert.equal(dict.interface.name, "IDictionary");
      await dict.Add("k", 1);
      assert.equal(dict.Count, 1);
    }
  });

  it("keep dual interfaces typed through IDispatch when their vtable cannot be marshaled", async () => {
    using _ = apartment();
    // Shell.Application lives in COM's host apartment: IShellDispatch6 has no proxy/stub.
    using shell = (await CreateObject("Shell.Application")) as any;
    assert.equal(isDispatch(shell), false);
    assert.equal(shell.interface.name, "IShellDispatch6");
    const pending = shell.NameSpace(process.env.SystemRoot);
    assert.ok(pending instanceof Promise, "methods return promises");
    using folder = await pending;
    assert.ok(typeof folder.Title === "string" && folder.Title.length > 0, "properties are synchronous");
  });

  it("call dynamic methods asynchronously; properties and sync() views stay synchronous", async () => {
    using _ = apartment();
    using dict = await CreateObject("Scripting.Dictionary", { dynamic: true });
    const added = dict.Add("a", 1);
    assert.ok(added instanceof Promise);
    await added;
    assert.equal(dict.Count, 1);
    assert.equal(dict.Item("a"), 1, "parameterized properties are synchronous");
    assert.equal(dict("a"), 1, "a property default member is synchronous");
    const view = sync(dict);
    assert.equal(sync(dict), view);
    assert.equal(view.Exists("a"), true, "sync() views call methods synchronously");
    view.Add("b", 2);
    assert.equal(dict.Count, 2);
    assert.equal(await invokeMethod(dict, "Exists", "b"), true);
  });

  it("enumerate collections with for await, also stopping early", async () => {
    using _ = apartment();
    using dict = await CreateObject("Scripting.Dictionary", { dynamic: true });
    for (let index = 0; index < 70; index++) await dict.Add(`k${index}`, index);
    const keys: string[] = [];
    for await (const key of dict) keys.push(key);
    assert.equal(keys.length, 70, "more than one batch");
    assert.deepEqual(keys.slice(0, 3), ["k0", "k1", "k2"]);
    assert.deepEqual(keys, [...dict], "the same items as for...of");
    const before = getLivePointerCount();
    for await (const key of dict) if (key === "k1") break;
    assert.equal(getLivePointerCount(), before, "the enumerator is released");
    // Typed dispatch pointers too, with objects handed back to the home apartment.
    // (Shell FolderItems repeats one item when Next is asked for more than one.)
    using shell = (await CreateObject("Shell.Application")) as any;
    using folder = await shell.NameSpace(process.env.SystemRoot);
    using items = await folder.Items();
    const names = (iterable: Iterable<any>, limit: number) => {
      const result: string[] = [];
      for (const item of iterable) {
        using owned = item;
        result.push(owned.Name);
        if (result.length === limit) break;
      }
      return result;
    };
    const expected = names(items, 40);
    const fetched: string[] = [];
    for await (const item of items) {
      using owned = item;
      fetched.push(owned.Name);
      if (fetched.length === expected.length) break;
    }
    assert.ok(expected.length > 32, "more than one batch");
    assert.deepEqual(fetched, expected);
  });

  it("run() cancels operations still queued on the executor", async () => {
    using _ = apartment();
    using dict = (await CreateObject("Scripting.Dictionary")) as any;
    const blocker = busy(200);
    const controller = new AbortController();
    const reason = new Error("cancelled");
    const added = run(dict.Add("late", 1), { signal: controller.signal });
    controller.abort(reason);
    await assert.rejects(added, (error) => error === reason);
    await blocker;
    // The queued call was skipped.
    await busy(0);
    assert.equal(dict.Count, 0);
    // An already aborted signal rejects at once (and skips the queued call); run() without a signal just awaits.
    const second = busy(100);
    await assert.rejects(run(dict.Add("x", 1), { signal: AbortSignal.abort(reason) }), (error) => error === reason);
    await second;
    await run(dict.Add("y", 1));
    assert.equal(dict.Count, 1);
  });
});
