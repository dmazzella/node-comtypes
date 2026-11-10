import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import { once } from "node:events";
import { describe, it } from "node:test";
import { threadId, Worker } from "node:worker_threads";

import {
  COMError,
  COMInterface,
  CreateObject,
  DISPMETHOD,
  DISP_E_MEMBERNOTFOUND,
  Dispatch,
  GetEvents,
  GetModule,
  GlobalInterfaceTable,
  IDispatch,
  VARIANT,
  VT_BOOL,
  VT_BSTR,
  VT_I4,
  VT_RECORD,
  apartment,
  byRef,
  dispid,
  getLivePointerCount,
  marshalInterface,
  record,
  recordInfo,
  releaseMarshaled,
  safeArray,
  startMessagePump,
  systemMessage,
} from "@node-ctypes/comtypes";
import type { ByRef, CoClass, RecordValue } from "@node-ctypes/comtypes";
import { COMObject, DispatchError, getLiveObjectCount } from "@node-ctypes/comtypes/server";
import type { InvokeArguments } from "@node-ctypes/comtypes/server";

const skip = process.platform !== "win32" && "COM requires Windows";
const DRAWING = "C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\System.Drawing.tlb";
const DISPATCH_METHOD = 1;

/** Doubles (or negates) every ByRef argument; answers only method calls. */
class Doubler extends COMObject {
  static _com_interfaces_ = [IDispatch];
  _invoke_(_dispid: number, flags: number, args: InvokeArguments): unknown {
    if (!(flags & DISPATCH_METHOD)) throw new DispatchError(DISP_E_MEMBERNOTFOUND);
    let count = 0;
    for (const reference of args.references) {
      if (!reference) continue;
      count++;
      const value = reference.value;
      reference.value = typeof value === "string" ? value + value : typeof value === "boolean" ? !value : (value as number) * 2;
    }
    return count;
  }
  Twice() {}
}

describe("ByRef arguments", { skip }, () => {
  it("round-trip through IDispatch::Invoke, as VARIANTs or typed", () => {
    using _ = apartment();
    using doubler = Dispatch(new Doubler());
    const variant = byRef(21);
    const text = byRef("ab");
    const typedNumber = byRef(5, VT_I4);
    const typedText = byRef("x", VT_BSTR);
    const typedBool = byRef(false, VT_BOOL);
    assert.equal(doubler.Twice(variant, text, 7, typedNumber, typedText, typedBool), 5);
    assert.deepEqual([variant.value, text.value, typedNumber.value, typedText.value, typedBool.value], [42, "abab", 10, "xx", true]);
    assert.throws(() => byRef(1, VT_I4 | 0x4000), RangeError);
    assert.throws(() => new VARIANT(byRef(1)), TypeError);
  });

  it("let event handlers answer through ByRef parameters", () => {
    using _ = apartment();
    const IEvents = COMInterface({ name: "IClosingEvents", iid: "{C0000000-0000-0000-0000-0000000000E1}", extends: IDispatch, dispMethods: [DISPMETHOD([dispid(1)], null, "BeforeClose", [["in"], VARIANT, "cancel"])] });
    const ISource = COMInterface({ name: "IClosingSource", iid: "{C0000000-0000-0000-0000-0000000000E2}", extends: IDispatch });
    class Window extends COMObject {
      static _com_interfaces_ = [ISource];
      static _outgoing_interfaces_ = [IEvents];
      close(): boolean {
        const cancel = byRef(false);
        this.fireEvent("BeforeClose", cancel);
        return cancel.value;
      }
    }
    const window = new Window();
    assert.equal(window.close(), false);
    using connection = GetEvents(window, { BeforeClose: (cancel: ByRef<boolean>) => (cancel.value = true) });
    assert.equal(window.close(), true);
    void connection;
  });
});

describe("VT_RECORD", { skip: skip || (!existsSync(DRAWING) && "System.Drawing.tlb (.NET Framework 4) is not installed") }, () => {
  it("converts records with IRecordInfo", () => {
    using _ = apartment();
    const Drawing = GetModule(DRAWING);
    const point = record(Drawing.Point as object, { x: 3, y: -4 });
    using variant = new VARIANT(point);
    assert.equal(variant.vt, VT_RECORD);
    const value = variant.value as RecordValue;
    assert.deepEqual({ ...value }, { x: 3, y: -4 });
    assert.equal(value[recordInfo].name, "Point");
    assert.ok(value[recordInfo].guid.equals("{CD6CB0A8-D6EF-33E8-888E-FE8C78CA568F}"));
    assert.deepEqual(value[recordInfo].fields, ["x", "y"]);
    using copy = new VARIANT(value);
    assert.deepEqual({ ...(copy.value as object) }, { x: 3, y: -4 }, "values read back are records again");
    const Rectangle = Drawing.Rectangle as new () => { x: number; width: number };
    const rectangle = new Rectangle();
    rectangle.x = 1;
    rectangle.width = 10;
    using fromInstance = new VARIANT(rectangle);
    assert.deepEqual({ ...(fromInstance.value as object) }, { x: 1, y: 0, width: 10, height: 0 });
  });

  it("stores records in SAFEARRAYs and passes them to automation objects", () => {
    using _ = apartment();
    const Drawing = GetModule(DRAWING);
    const points = [record(Drawing.Point as object, { x: 1, y: 2 }), record(Drawing.Point as object, { y: 7 })];
    using records = new VARIANT(safeArray(points, { vt: VT_RECORD }));
    assert.equal(records.vt, 0x2000 | VT_RECORD);
    assert.deepEqual(
      (records.value as object[]).map((point) => ({ ...point })),
      [
        { x: 1, y: 2 },
        { x: 0, y: 7 },
      ],
    );
    using variants = new VARIANT(points);
    assert.equal((variants.value as RecordValue[])[1].y, 7);
    using dict = CreateObject("Scripting.Dictionary");
    dict.Add("p", points[0]);
    assert.deepEqual({ ...dict.Item("p") }, { x: 1, y: 2 });
    assert.throws(() => record({} as object), TypeError);
  });
});

describe("cross-thread references", { skip }, () => {
  class Calculator extends COMObject {
    static _com_interfaces_ = [IDispatch];
    calls = 0;
    get ThreadId() {
      return threadId;
    }
    Add(a: number, b: number) {
      this.calls++;
      return a + b;
    }
    Echo(value: unknown) {
      return value;
    }
  }
  const fixture = new URL("../fixtures/worker.ts", import.meta.url);

  it("marshals a JavaScript object to STA and MTA worker threads", async () => {
    using _ = apartment();
    using pump = startMessagePump({ interval: 2 });
    const calculator = new Calculator();
    for (const mta of [false, true]) {
      const worker = new Worker(fixture, { workerData: { token: marshalInterface(calculator, IDispatch), mta } });
      const workerId = worker.threadId;
      // Listen for exit before awaiting the message: a worker may exit right after posting.
      const exited = once(worker, "exit");
      const [result] = await once(worker, "message");
      await exited;
      assert.deepEqual(result, { apartment: mta ? "MTA" : "STA", sum: 42, owner: threadId, worker: workerId, echo: ["a", 1] });
    }
    assert.equal(calculator.calls, 2, "calls run on the owning (main) thread");
    assert.equal(calculator.refCount, 0, "the proxies were released");
    void pump;
  });

  it("shares a reference through the Global Interface Table", async () => {
    using _ = apartment();
    using pump = startMessagePump({ interval: 2 });
    const calculator = new Calculator();
    using table = new GlobalInterfaceTable();
    const cookie = table.register(calculator, IDispatch);
    const workers = [0, 1, 2].map(() => new Worker(fixture, { workerData: { cookie } }));
    const exited = Promise.all(workers.map((worker) => once(worker, "exit")));
    const results = await Promise.all(workers.map((worker) => once(worker, "message")));
    await exited;
    assert.deepEqual(
      results.map(([result]) => result.sum),
      [42, 42, 42],
    );
    using local = table.getDispatch(cookie);
    assert.equal(local.Add(1, 1), 2);
    table.revoke(cookie);
    assert.equal(calculator.calls, 4);
    void pump;
  });

  it("releases unused tokens", () => {
    using _ = apartment();
    const objects = getLiveObjectCount();
    const calculator = new Calculator();
    const token = marshalInterface(calculator, IDispatch);
    assert.equal(calculator.refCount > 0, true);
    releaseMarshaled(token);
    assert.equal(calculator.refCount, 0);
    assert.equal(getLiveObjectCount(), objects);
    assert.throws(() => releaseMarshaled({ kind: "x" } as never), TypeError);
    void getLivePointerCount;
  });
});

describe("Visual Basic error messages", { skip }, () => {
  it("describe FACILITY_CONTROL HRESULTs", () => {
    using _ = apartment();
    const text = systemMessage(0x800a01c9 | 0);
    // vbscript.dll is a Feature on Demand: the text is localized, or null when it is not installed.
    if (text === null) return;
    assert.ok(text.length > 0);
    const Scripting = GetModule("scrrun.dll");
    using dict = CreateObject(Scripting.Dictionary as CoClass) as unknown as Disposable & { Add(key: string, value: number): void };
    dict.Add("a", 1);
    assert.throws(
      () => dict.Add("a", 2),
      (error: unknown) => error instanceof COMError && error.message.includes(text),
    );
  });
});
