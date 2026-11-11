// Child process for tests/src/executor.test.ts: the libuv pool is reduced to
// one thread and kept busy by pbkdf2; a call on the COM executor must still
// complete, because the executor does not use libuv workers.
import { pbkdf2 } from "node:crypto";
import * as ctypes from "node-ctypes";
import { executorCall } from "@node-ctypes/comtypes/low-level";

const kernel32 = new ctypes.WinDLL("kernel32.dll");
const GetCurrentThreadId = kernel32.func("GetCurrentThreadId", ctypes.c_uint32, []);

let cryptoDone = false;
pbkdf2("password", "salt", 2_000_000, 64, "sha512", () => {
  cryptoDone = true;
});
const { value } = await executorCall({ fn: GetCurrentThreadId.address, restype: ctypes.c_uint32, argtypes: [], args: [] });
process.stdout.write(JSON.stringify({ executorThread: value, cryptoDone, threadpool: process.env.UV_THREADPOOL_SIZE }));
