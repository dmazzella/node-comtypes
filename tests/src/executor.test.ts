import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { describe, it } from "node:test";
import { fileURLToPath } from "node:url";
import * as ctypes from "node-ctypes";

import { CreateObject, GetModule, apartment, finalize, stats } from "@node-ctypes/comtypes";
import { allocBSTR, executorCall, executorGeneration, oleaut32, vtableEntry } from "@node-ctypes/comtypes/low-level";

const skip = process.platform !== "win32" && "COM requires Windows";
const kernel32 = process.platform === "win32" ? new ctypes.WinDLL("kernel32.dll") : null;
const user32 = process.platform === "win32" ? new ctypes.WinDLL("user32.dll") : null;
const NO_ARGS: ctypes.FFIType[] = [];
const SLEEP_ARGS: ctypes.FFIType[] = [ctypes.c_uint32];
const POINTER_PAIR: ctypes.FFIType[] = [ctypes.c_void_p, ctypes.c_void_p];
const threadId = () => executorCall({ fn: kernel32!.func("GetCurrentThreadId", ctypes.c_uint32, []).address, restype: ctypes.c_uint32, argtypes: NO_ARGS, args: [] });
const sleep = (ms: number, signal?: AbortSignal) => executorCall({ fn: kernel32!.func("Sleep", null, [ctypes.c_uint32]).address, restype: null, argtypes: SLEEP_ARGS, args: [ms] }, { signal });
const fixture = (name: string) => fileURLToPath(new URL(`../fixtures/${name}`, import.meta.url));

describe("COM executor", { skip }, () => {
  it("runs calls on one dedicated thread, in order, each settled once", async () => {
    const mainThread = kernel32!.func("GetCurrentThreadId", ctypes.c_uint32, [])();
    const settled: number[] = [];
    const calls = Array.from({ length: 500 }, (_, index) => threadId().then((result) => (settled.push(index), result.value as number)));
    const threads = new Set(await Promise.all(calls));
    assert.equal(threads.size, 1, "a single executor thread");
    assert.notEqual([...threads][0], mainThread, "not the JavaScript thread");
    assert.deepEqual(
      settled,
      calls.map((_, index) => index),
      "FIFO, exactly once",
    );
    assert.equal(stats().pending, 0);
    assert.equal(stats().running, true);
  });

  it("runs in the multithreaded apartment", async () => {
    const getApartmentType = ole32Function("CoGetApartmentType", [ctypes.c_void_p, ctypes.c_void_p]);
    const type = new ctypes.c_int32();
    const qualifier = new ctypes.c_int32();
    const { value } = await executorCall({ fn: getApartmentType, restype: ctypes.c_int32, argtypes: POINTER_PAIR, args: [ctypes.byref(type), ctypes.byref(qualifier)] });
    assert.equal(value, 0);
    assert.equal(type.value, 1, "APTTYPE_MTA");
  });

  it("keeps the event loop free during a long native call", async () => {
    let ticks = 0;
    const timer = setInterval(() => ticks++, 10);
    try {
      await sleep(300);
    } finally {
      clearInterval(timer);
    }
    assert.ok(ticks >= 10, `timers ran meanwhile (${ticks} ticks)`);
  });

  it("passes structures by value and memory by address", async () => {
    class RECT extends ctypes.Structure {
      static _fields_ = [
        ["left", ctypes.c_int32],
        ["top", ctypes.c_int32],
        ["right", ctypes.c_int32],
        ["bottom", ctypes.c_int32],
      ];
    }
    class POINT extends ctypes.Structure {
      static _fields_ = [
        ["x", ctypes.c_int32],
        ["y", ctypes.c_int32],
      ];
    }
    const PtInRect = user32!.func("PtInRect", ctypes.c_int32, [ctypes.c_void_p, POINT]);
    const argtypes = [ctypes.c_void_p, POINT];
    const rect = RECT.create({ left: 0, top: 0, right: 10, bottom: 10 });
    const call = (x: number, y: number) => executorCall({ fn: PtInRect.address, restype: ctypes.c_int32, argtypes, args: [rect, POINT.create({ x, y })] });
    assert.notEqual((await call(5, 5)).value, 0);
    assert.equal((await call(50, 5)).value, 0);
  });

  it("reads IErrorInfo on the executor thread after a failure", async () => {
    using _ = apartment();
    const MSXML = (await GetModule("msxml6.dll")) as Record<string, any>;
    using document = (await CreateObject("MSXML2.DOMDocument.6.0", { interface: MSXML.IXMLDOMDocument })) as any;
    // The DOM document is free-threaded: callable from the executor's apartment.
    const argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p];
    const slot = MSXML.IXMLDOMDocument.methodIndex.get("createElement");
    const createElement = (name: bigint, out: InstanceType<typeof ctypes.c_void_p>) =>
      executorCall({ fn: vtableEntry(document.address, slot), restype: ctypes.c_int32, argtypes, args: [document.address, name, ctypes.byref(out)], errorInfo: { object: document.address, iid: MSXML.IXMLDOMDocument.iid } });
    const invalid = allocBSTR("1invalid");
    const valid = allocBSTR("valid");
    const element = new ctypes.c_void_p();
    try {
      const failure = await createElement(invalid, element);
      assert.ok((failure.value as number) < 0);
      assert.match(failure.details?.description ?? "", /"1"|1/);
      // The error object was taken: a later success reports nothing.
      const success = await createElement(valid, element);
      assert.equal(success.value, 0);
      assert.equal(success.details, null);
      assert.ok(element.value);
      await executorCall({ fn: vtableEntry(element.value!, 2), restype: ctypes.c_uint32, argtypes: [ctypes.c_void_p], args: [element.value] });
      // GetErrorInfo runs only right after the failed call: with another call
      // already submitted behind it, the failure has no details.
      const failing = createElement(invalid, element);
      const behind = threadId();
      const burst = await failing;
      assert.ok((burst.value as number) < 0);
      assert.equal(burst.details, null);
      await behind;
      // Awaited alone again: the details are back.
      const again = await createElement(invalid, element);
      assert.match(again.details?.description ?? "", /"1"|1/);
    } finally {
      oleaut32().SysFreeString(invalid);
      oleaut32().SysFreeString(valid);
    }
  });

  it("cancels a queued call with an AbortSignal", async () => {
    const controller = new AbortController();
    const running = sleep(200);
    const queued = sleep(1, controller.signal);
    const reason = new Error("stop");
    controller.abort(reason);
    await assert.rejects(queued, (error) => error === reason);
    await running;
  });

  it("drains accepted calls on finalize, then starts a new executor", async () => {
    await threadId();
    const before = executorGeneration();
    let settled = 0;
    const accepted = Array.from({ length: 5 }, () => sleep(20).then(() => settled++));
    await finalize();
    assert.equal(settled, 5, "accepted calls settled before finalize() resolved");
    assert.equal(stats().running, false);
    await Promise.all(accepted);
    assert.equal(typeof (await threadId()).value, "number");
    assert.equal(executorGeneration(), before + 1);
    assert.equal(stats().running, true);
  });

  it("does not use libuv workers", () => {
    const child = spawnSync(process.execPath, [fixture("executor-pool.ts")], { env: { ...process.env, UV_THREADPOOL_SIZE: "1" }, encoding: "utf8", timeout: 60_000 });
    assert.equal(child.status, 0, child.stderr);
    const result = JSON.parse(child.stdout);
    assert.equal(result.cryptoDone, false, "the executor answered while pbkdf2 held the only libuv worker");
    assert.equal(typeof result.executorThread, "number");
  });

  it("lets the process exit when idle", () => {
    const started = performance.now();
    const child = spawnSync(process.execPath, [fixture("executor-exit.ts")], { encoding: "utf8", timeout: 30_000 });
    assert.equal(child.status, 0, child.stderr);
    assert.match(child.stdout, /done/);
    assert.ok(performance.now() - started < 20_000, "no hang at exit");
  });
});

function ole32Function(name: string, argtypes: ctypes.FFIType[]): bigint {
  return new ctypes.WinDLL("ole32.dll").func(name, ctypes.c_int32, argtypes).address;
}
