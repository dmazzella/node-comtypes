import assert from "node:assert/strict";
import { describe, it } from "node:test";
import * as ctypes from "node-ctypes";

import { CreateObject, apartment, finalize, getLivePointerCount, stats } from "@node-ctypes/comtypes";
import { executorCall } from "@node-ctypes/comtypes/low-level";

const skip = process.platform !== "win32" && "COM requires Windows";
const kernel32 = process.platform === "win32" ? new ctypes.WinDLL("kernel32.dll") : null;
const SLEEP_ARGS: ctypes.FFIType[] = [ctypes.c_uint32];
const sleep = (ms: number) => executorCall({ fn: kernel32!.func("Sleep", null, [ctypes.c_uint32]).address, restype: null, argtypes: SLEEP_ARGS, args: [ms] });

describe("executor backlog", { skip }, () => {
  it("settles thousands of queued method calls exactly once, in order", async () => {
    using _ = apartment();
    using dict = (await CreateObject("Scripting.Dictionary")) as any;
    // The first call obtains the executor's own reference: later calls are queued at once.
    assert.equal(await dict.Exists("k0"), false);
    const pointers = getLivePointerCount();
    const settled: number[] = [];
    const calls = Array.from({ length: 5000 }, (_, index) => dict.Add(`k${index}`, index).then(() => settled.push(index)));
    assert.ok(stats().pending > 0, "queued on the executor");
    await Promise.all(calls);
    assert.equal(settled.length, 5000);
    assert.ok(
      settled.every((value, index) => value === index),
      "FIFO",
    );
    assert.equal(dict.Count, 5000);
    assert.equal(dict.Item("k4999"), 4999);
    assert.equal(getLivePointerCount(), pointers, "no pointer leaked");
    assert.equal(stats().pending, 0);
  });

  it("keeps the event loop free while calls queue behind a long one", async () => {
    using _ = apartment();
    using dict = (await CreateObject("Scripting.Dictionary")) as any;
    let ticks = 0;
    const timer = setInterval(() => ticks++, 5);
    try {
      const long = sleep(300);
      const queued = Array.from({ length: 200 }, (_, index) => dict.Add(`q${index}`, index));
      // Properties are read synchronously meanwhile (on the home thread, through the proxy).
      assert.equal(dict.Count, 0);
      await Promise.all([long, ...queued]);
    } finally {
      clearInterval(timer);
    }
    // Windows timers fire about every 15.6 ms.
    assert.ok(ticks >= 10, `timers ran meanwhile (${ticks} ticks)`);
    assert.equal(dict.Count, 200);
  });

  it("finalize drains a backlog, then a new executor serves the same objects", async () => {
    using _ = apartment();
    using dict = (await CreateObject("Scripting.Dictionary")) as any;
    // The first call obtains the executor's own reference: later calls are queued at once.
    await dict.Add("first", 0);
    const generation = stats().generation;
    let settled = 0;
    const calls = Array.from({ length: 1000 }, (_, index) => dict.Add(`f${index}`, index).then(() => settled++));
    await finalize();
    assert.equal(settled, 1000, "accepted calls settled before finalize() resolved");
    await Promise.all(calls);
    await dict.Add("after", 1);
    assert.equal(stats().generation, generation + 1);
    assert.equal(dict.Count, 1002);
  });
});
