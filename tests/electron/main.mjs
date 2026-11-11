// Manual check in Electron's main process, whose JavaScript thread is
// Chromium's single-threaded apartment (UI thread). Run after `npm run build`:
//
//   electron tests/electron/main.mjs
//
// (unset ELECTRON_RUN_AS_NODE if set). Prints a JSON report and exits with 0
// when every check passed.
import { app } from "electron";
import * as ctypes from "node-ctypes";
import { CreateObject, GetEvents, apartment, getApartmentType, isDispatch, run, stats } from "../../dist/index.js";
import { executorCall } from "../../dist/low-level/index.js";

const checks = {};
const check = async (name, body) => {
  try {
    checks[name] = { ok: true, value: await body() };
  } catch (error) {
    checks[name] = { ok: false, error: String(error?.stack ?? error) };
  }
};

app.whenReady().then(async () => {
  // Chromium initialized COM on this thread: apartment() only balances it.
  using _ = apartment();
  await check("home apartment", () => {
    const type = getApartmentType()?.name;
    if (type !== "STA" && type !== "MAINSTA") throw new Error(`expected a single-threaded apartment, got ${type}`);
    return type;
  });

  await check("UI thread free during a 500 ms executor call", async () => {
    const kernel32 = new ctypes.WinDLL("kernel32.dll");
    const sleep = kernel32.func("Sleep", null, [ctypes.c_uint32]);
    let ticks = 0;
    const timer = setInterval(() => ticks++, 10);
    await executorCall({ fn: sleep.address, restype: null, argtypes: [ctypes.c_uint32], args: [500] });
    clearInterval(timer);
    if (ticks < 15) throw new Error(`the event loop was blocked (${ticks} ticks)`);
    return { ticks };
  });

  await check("objects created on the executor: async methods, sync properties", async () => {
    using dict = await CreateObject("Scripting.Dictionary");
    const pending = dict.Add("a", 1);
    if (!(pending instanceof Promise)) throw new Error("methods must return promises");
    await pending;
    await dict.Add("b", 2);
    return { interface: dict.interface.name, count: dict.Count, item: dict.Item("b") };
  });

  await check("typed Shell objects and for await", async () => {
    using shell = await CreateObject("Shell.Application");
    if (isDispatch(shell)) throw new Error("Shell.Application should stay typed");
    using folder = await shell.NameSpace(process.env.SystemRoot);
    using items = await folder.Items();
    const names = [];
    for await (const item of items) {
      using owned = item;
      names.push(owned.Name);
      if (names.length === 3) break;
    }
    return { interface: shell.interface.name, title: folder.Title, names };
  });

  await check("WMI events delivered to a JavaScript sink (Chromium pumps messages)", async () => {
    using locator = await CreateObject("WbemScripting.SWbemLocator", { dynamic: true });
    using services = await locator.ConnectServer();
    using sink = await CreateObject("WbemScripting.SWbemSink", { dynamic: true });
    const names = [];
    const completed = new Promise((resolve, reject) => {
      const timeout = setTimeout(() => reject(new Error("no OnCompleted event")), 20_000);
      checks.connection = GetEvents(sink, {
        OnObjectReady(object) {
          names.push(object.Name);
        },
        OnCompleted() {
          clearTimeout(timeout);
          resolve();
        },
      });
    });
    await services.ExecQueryAsync(sink, `SELECT Name FROM Win32_Process WHERE ProcessId = ${process.pid}`);
    await completed;
    checks.connection[Symbol.dispose]();
    delete checks.connection;
    return { names };
  });

  await check("run() rejects on abort", async () => {
    using dict = await CreateObject("Scripting.Dictionary");
    const reason = new Error("aborted");
    try {
      await run(dict.Add("x", 1), { signal: AbortSignal.abort(reason) });
    } catch (error) {
      if (error === reason) return "rejected with the reason";
      throw error;
    }
    throw new Error("run() did not reject");
  });

  const ok = Object.values(checks).every((entry) => entry.ok);
  console.log(JSON.stringify({ electron: process.versions.electron, node: process.versions.node, arch: process.arch, ok, checks, stats: stats() }, null, 2));
  app.exit(ok ? 0 : 1);
});
