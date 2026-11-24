import * as ctypes from "node-ctypes";
import { checkHRESULT } from "./errors.ts";
import { ole32 } from "./low-level/native.ts";

// The Node.js event loop does not dispatch Win32 messages. Single-threaded
// apartments receive cross-apartment and cross-process calls, and many event
// sources fire events, through window messages: they need a message pump.

const PM_REMOVE = 0x1;
const QS_ALLINPUT = 0x04ff;
const MWMO_INPUTAVAILABLE = 0x0004;
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
    MsgWaitForMultipleObjectsEx: library.func("MsgWaitForMultipleObjectsEx", ctypes.c_uint32, [ctypes.c_uint32, ctypes.c_void_p, ctypes.c_uint32, ctypes.c_uint32, ctypes.c_uint32]),
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

let msgSize = 0;
const msgPool: Buffer[] = [];

/**
 * Dispatch the messages waiting in the thread's queue, without blocking.
 * Returns how many were dispatched (at most `limit`). Re-entrant safe: nested
 * `pumpMessages()` calls from inside `DispatchMessageW` use their own MSG buffer.
 */
export function pumpMessages(limit = Number.POSITIVE_INFINITY): number {
  const api = user32();
  const msg = msgPool.pop() ?? Buffer.alloc((msgSize ||= ctypes.sizeof(ctypes.wintypes.MSG)));
  let count = 0;
  try {
    while (count < limit && api.PeekMessageW(msg, null, 0, 0, PM_REMOVE)) {
      api.TranslateMessage(msg);
      api.DispatchMessageW(msg);
      count++;
    }
  } finally {
    if (msgPool.length < 4) msgPool.push(msg);
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

/**
 * Drain all queued Win32 messages, or if the queue is empty sleep in the
 * Windows kernel via `MsgWaitForMultipleObjectsEx` (~0% CPU) until any input
 * or COM STA message arrives or `timeoutMs` elapses, then drain immediately.
 */
export function waitAndPumpMessages(timeoutMs = 8, limit = Number.POSITIVE_INFINITY): number {
  let count = pumpMessages(limit);
  if (count > 0) return count;
  const wait = user32().MsgWaitForMultipleObjectsEx(0, null, Math.max(0, Math.floor(timeoutMs)) >>> 0, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
  if (wait === 0) count += pumpMessages(limit);
  return count;
}

export interface MessagePump extends Disposable {
  stop(): void;
  readonly running: boolean;
}

export interface StartMessagePumpOptions {
  /**
   * Polling interval in ms when `mode === "interval"` (default `10`), or kernel
   * idle yield in ms when `mode === "wait"` (default `8`).
   */
  readonly interval?: number;
  /**
   * - `"interval"` (default): `setInterval` pump (non-blocking, ~15.6 ms timer granularity on Windows).
   * - `"wait"`: `MsgWaitForMultipleObjectsEx` + `setImmediate` pump — wakes in < 0.1 ms on incoming
   *   Win32/COM/UI messages, sleeps at ~0% CPU when idle, and yields every `interval` ms to libuv.
   */
  readonly mode?: "interval" | "wait";
  /** Whether the pump keeps the Node.js event loop alive (default `false` for `"interval"`, `true` for `"wait"`). */
  readonly ref?: boolean;
}

/**
 * Pump Win32 messages from the Node.js event loop.
 *
 * - Default (`mode: "interval"`): periodic unref'd `setInterval` — leaves the
 *   JS thread completely idle between ticks.
 * - Low-latency (`mode: "wait"`): `MsgWaitForMultipleObjectsEx` + `setImmediate`
 *   — sub-millisecond wakeup for interactive UI / STA COM callbacks.
 */
export function startMessagePump(options: StartMessagePumpOptions = {}): MessagePump {
  const mode = options.mode ?? "interval";
  let running = true;

  if (mode === "wait") {
    const idleYieldMs = options.interval ?? 8;
    const keepRef = options.ref ?? true;
    let handle: NodeJS.Immediate | null = null;
    const step = () => {
      handle = null;
      if (!running) return;
      waitAndPumpMessages(idleYieldMs);
      if (running) {
        handle = setImmediate(step);
        if (!keepRef) handle.unref();
      }
    };
    handle = setImmediate(step);
    if (!keepRef) handle.unref();
    const stop = () => {
      if (!running) return;
      running = false;
      if (handle !== null) {
        clearImmediate(handle);
        handle = null;
      }
    };
    return {
      stop,
      get running() {
        return running;
      },
      [Symbol.dispose]: stop,
    };
  }

  const interval = options.interval ?? 10;
  const timer = setInterval(() => pumpMessages(), interval);
  if (!options.ref) timer.unref();
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
 * the last value of the condition. Uses `waitAndPumpMessages` so STA COM events
 * wake immediately (< 0.1 ms) while still yielding to the Node.js event loop.
 */
export async function waitFor(condition: () => boolean, options: { readonly timeout?: number; readonly interval?: number } = {}): Promise<boolean> {
  const deadline = performance.now() + (options.timeout ?? 10_000);
  const sliceMs = options.interval ?? 8;
  for (;;) {
    pumpMessages();
    if (condition()) return true;
    const remaining = deadline - performance.now();
    if (remaining <= 0) return false;
    waitAndPumpMessages(Math.min(sliceMs, remaining));
    if (condition()) return true;
    if (performance.now() >= deadline) return false;
    await new Promise<void>((resolve) => setImmediate(resolve));
  }
}
