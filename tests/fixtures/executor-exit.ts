// Child process for tests/src/executor.test.ts: one call on the COM executor,
// then the process must exit on its own (the idle executor does not keep the
// event loop alive; beforeExit finalizes it).
import * as ctypes from "node-ctypes";
import { executorCall } from "@node-ctypes/comtypes/low-level";

const kernel32 = new ctypes.WinDLL("kernel32.dll");
const GetTickCount = kernel32.func("GetTickCount", ctypes.c_uint32, []);

await executorCall({ fn: GetTickCount.address, restype: ctypes.c_uint32, argtypes: [], args: [] });
process.stdout.write("done\n");
