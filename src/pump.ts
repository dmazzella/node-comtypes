import * as ctypes from "node-ctypes";
import { setTimeout as delay } from "node:timers/promises";
import { checkHRESULT } from "./errors.ts";
import { ole32 } from "./low-level/native.ts";

// The Node.js event loop does not dispatch Win32 messages. Single-threaded
// apartments receive cross-apartment and cross-process calls, and many event
// sources fire events, through window messages: they need a message pump.

const PM_REMOVE = 0x1;
const COWAIT_DISPATCH_CALLS = 0x8;
const COWAIT_DISPATCH_WINDOW_MESSAGES = 0x10;
const RPC_S_CALLPENDING = 0x80010115 | 0;

let user32Functions: ReturnType<typeof bindUser32> | null = null;
let kernel32Functions: ReturnType<typeof bindKernel32> | null = null;

function bindUser32(library: ctypes.WinDLL) {
  return {
    PeekMessageW: library.func("PeekMessageW", ctypes.c_int32, [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint32, ctypes.c_uint32, ctypes.c_uint32]),
    TranslateMessage: library.func("TranslateMessage", ctypes.c_int32, [ctypes.c_void_p]),
    DispatchMessageW: library.func("DispatchMessageW", ctypes.c_void_p, [ctypes.c_void_p]),
  };
}

function bindKernel32(library: ctypes.WinDLL) {
  return {
    CreateEventW: library.func("CreateEventW", ctypes.c_void_p, [ctypes.c_void_p, ctypes.c_int32, ctypes.c_int32, ctypes.c_void_p]),
    CloseHandle: library.func("CloseHandle", ctypes.c_int32, [ctypes.c_void_p]),
  };
}

function user32() {
  return (user32Functions ??= bindUser32(new ctypes.WinDLL("user32.dll")));
}

function kernel32() {
  return (kernel32Functions ??= bindKernel32(new ctypes.WinDLL("kernel32.dll")));
}

let message: Buffer | null = null;

/**
 * Dispatch the messages waiting in the thread's queue, without blocking.
 * Returns how many were dispatched (at most `limit`).
 */
export function pumpMessages(limit = Number.POSITIVE_INFINITY): number {
  const api = user32();
  const msg = (message ??= Buffer.alloc(ctypes.sizeof(ctypes.wintypes.MSG)));
  let count = 0;
  while (count < limit && api.PeekMessageW(msg, null, 0, 0, PM_REMOVE)) {
    api.TranslateMessage(msg);
    api.DispatchMessageW(msg);
    count++;
  }
  return count;
}

/**
 * Block for `milliseconds`, dispatching COM calls and window messages (the
 * event loop does not run meanwhile). Python equivalent: `comtypes.client.PumpEvents`.
 */
export function PumpEvents(milliseconds: number): void {
  if (!Number.isFinite(milliseconds) || milliseconds < 0) throw new RangeError("PumpEvents requires a non-negative timeout");
  const event = kernel32().CreateEventW(null, 1, 0, null);
  if (!event) throw new Error("CreateEventW failed");
  try {
    const handles = Buffer.alloc(ctypes.sizeof(ctypes.c_void_p));
    ctypes.writeValue(handles, ctypes.c_void_p, event);
    const index = new ctypes.c_uint32();
    const hr = ole32().CoWaitForMultipleHandles(COWAIT_DISPATCH_CALLS | COWAIT_DISPATCH_WINDOW_MESSAGES, Math.floor(milliseconds), 1, handles, ctypes.byref(index));
    if (hr !== RPC_S_CALLPENDING) checkHRESULT(hr, "CoWaitForMultipleHandles");
  } finally {
    kernel32().CloseHandle(event);
  }
}

export interface MessagePump extends Disposable {
  stop(): void;
  readonly running: boolean;
}

/**
 * Pump messages periodically from the event loop (`setInterval`, unref'd so
 * it does not keep the process alive). The JavaScript thread stays free between
 * ticks: use this for events and incoming calls while the program runs.
 */
export function startMessagePump(options: { readonly interval?: number } = {}): MessagePump {
  const interval = options.interval ?? 10;
  const timer = setInterval(() => pumpMessages(), interval);
  timer.unref();
  let running = true;
  const stop = () => {
    if (!running) return;
    running = false;
    clearInterval(timer);
  };
  return {
    stop,
    get running() {
      return running;
    },
    [Symbol.dispose]: stop,
  };
}

/**
 * Pump messages until `condition()` holds or `timeout` elapses; resolves to
 * the last value of the condition. The event loop runs between checks.
 */
export async function waitFor(condition: () => boolean, options: { readonly timeout?: number; readonly interval?: number } = {}): Promise<boolean> {
  const deadline = performance.now() + (options.timeout ?? 10_000);
  for (;;) {
    pumpMessages();
    if (condition()) return true;
    if (performance.now() >= deadline) return false;
    await delay(options.interval ?? 5);
  }
}
