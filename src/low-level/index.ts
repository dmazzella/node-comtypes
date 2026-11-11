export { ole32, oleaut32, allocBSTR, readBSTR, takeBSTR, takeCoTaskMemString, readMemory, writeMemory, isWindows, POINTER_SIZE } from "./native.ts";
export type { Ole32, OleAut32 } from "./native.ts";
export { vtableEntry, functionCache, rawQueryInterface, rawAddRef, rawRelease } from "./vtable.ts";
export { executorCall, executorGeneration } from "../executor/executor.ts";
export type { CallOptions, NativeCall, NativeResult } from "../executor/executor.ts";
export { COSERVERINFO, MULTI_QI } from "./types.ts";
export { HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, readString, subkeys, writeString, deleteTree } from "./registry.ts";
