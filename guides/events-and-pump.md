---
title: Events and message pump
---

# Events and message pump

## Subscribing to COM events (`GetEvents` / `ShowEvents`)

`GetEvents(source, handler, { interface? })` (Python: `comtypes.client.GetEvents`) locates the source's outgoing event interface (from `options.interface`, `IProvideClassInfo2`, the coclass's default `[source]` interface, or its type library), builds a sink `COMObject`, and calls `IConnectionPoint::Advise`. Disposing the returned `EventConnection` (or calling `.disconnect()`) calls `Unadvise`.

```typescript
import { GetEvents, ShowEvents, startMessagePump } from "@node-ctypes/comtypes";

declare const source: object; // COMPointer, dynamic object, or COMObject

using pump = startMessagePump({ mode: "wait" }); // low-latency STA message pump
using connection = GetEvents(source, {
  OnObjectReady(object: unknown) {
    console.log("ready", object);
  },
  OnCompleted(result: unknown) {
    console.log("completed", result);
  },
});
using log = ShowEvents(source); // prints every incoming event
```

- **Dispatching rules**: each event invokes `handler[Name](...args)`, falling back to `handler[Interface_Name](...args)` and then `handler.onEvent(name, args)`. Unhandled events are ignored.
- **Borrowed arguments**: `IDispatch` arguments arrive as dynamic objects borrowed for the duration of the callback; `ByRef` event parameters (such as `Cancel`) arrive as `ByRef` wrappers — assign `cancel.value = true` to reply to the source.
- **Both event styles**: dispinterface and custom vtable event interfaces are supported.
- **Teardown safety**: any `EventConnection` still active when the thread's apartment tears down (`CoUninitialize`) is cleanly unadvised and released automatically so sink objects do not leak.

See [examples/events.ts](../examples/events.ts) for a complete runnable WMI event subscription example.

> **Note on Active Directory / ADSI**: ADSI interfaces (`IDirectorySearch`, `IADs`) do not implement `IConnectionPointContainer`, so `GetEvents` cannot attach to them directly. For AD change notifications via `GetEvents`, subscribe through the WMI LDAP provider (`winmgmts:root\directory\LDAP`) with `SWbemSink`.

## Firing events from JavaScript (`COMObject`)

A `COMObject` becomes an event source by declaring `static _outgoing_interfaces_` (which automatically implements `IConnectionPointContainer` and `IProvideClassInfo2`) and calling `this.fireEvent("Name", ...args)`:

- For dispinterfaces, `fireEvent` invokes each connected sink through `IDispatch::Invoke` using the member's DISPID.
- For vtable event interfaces, `fireEvent` calls the matching vtable slot directly.

See [examples/server.ts](../examples/server.ts) for a full example.

## Win32 message pump

The Node.js event loop does not dispatch Win32 window messages. Out-of-process COM servers, STA cross-apartment marshaling and many in-process event sources (such as WMI `SWbemSink`) deliver callbacks through the STA's hidden window message queue, so you need a pump while waiting for events or serving calls.

| Function                                    | Behavior                                                                                                                                                                                                             |
| ------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `startMessagePump({ mode, interval, ref })` | Runs in the background while keeping the event loop free (unref'd by default). `"interval"` (default) polls via `PeekMessage`; `"wait"` uses short `waitAndPumpMessages()` slices for near-zero event wakeup latency |
| `waitAndPumpMessages(timeoutMs?, limit?)`   | Blocks efficiently in `MsgWaitForMultipleObjectsEx` until a Win32 message arrives (or `timeoutMs` elapses), then drains queued messages without CPU spinning                                                         |
| `await waitFor(condition, { timeout })`     | Pumps (using `waitAndPumpMessages`) until `condition()` returns truthy or rejects on timeout                                                                                                                         |
| `pumpMessages(limit?)`                      | Drains currently queued Win32 messages without blocking (re-entrant safe)                                                                                                                                            |
| `PumpEvents(ms)`                            | Blocks for `ms` in `CoWaitForMultipleHandles`, dispatching COM calls and messages (Python: `PumpEvents`)                                                                                                             |

```typescript
import { startMessagePump, waitAndPumpMessages, waitFor } from "@node-ctypes/comtypes";

// Low-latency background pump while awaiting async work:
using pump = startMessagePump({ mode: "wait" });
await waitFor(() => received >= 3, { timeout: 5_000 });

// Or drive a single slice explicitly in a custom loop:
const dispatched = waitAndPumpMessages(25);
```

In Electron's main process, Chromium already pumps the Win32 message loop, so `startMessagePump()` is unnecessary there.
