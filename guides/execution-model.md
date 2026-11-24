---
title: Execution model
---

# Execution model

**Calls are asynchronous; live access is synchronous.** Methods return promises and run on the **COM executor**, a native thread owned by `@node-ctypes/comtypes` (a `node-ctypes` `Executor`, not a libuv worker) initialized in the multithreaded apartment (MTA), executing one call at a time in FIFO order. The JavaScript event loop stays responsive while slow objects (Office applications, remote DCOM servers, WMI, ADSI paging, or unresponsive servers) are running.

| Asynchronous (`Promise`)                                            | Synchronous                                                                                                |
| ------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------- |
| `COMMETHOD` / `STDMETHOD` methods, `DISPMETHOD`s, dynamic methods   | Properties: get, put, putref, parameterized (`dict.Item("a")`, `_get_X` / `_set_X`)                        |
| `CreateObject`, `GetObject`, `GetActiveObject`                      | `QueryInterface`, `AddRef`, `Release`, `using`, `GetBestInterface`, `queryService`                         |
| `for await` over collections (`IEnumVARIANT::Next` on the executor) | `for...of`, `VARIANT` conversions, `Constants`, `getModuleSync`, `GetEvents`, generated modules            |
| `GetModule` (in event-loop slices on the JS thread)                 | Server side (`COMObject` implementations, `fireEvent`, registration); [`sync()`](#synchronous-calls) views |
| `executorCall` (low-level)                                          |                                                                                                            |

```typescript
import { apartment, CreateObject, run } from "@node-ctypes/comtypes";

using _ = apartment();
using excel = await CreateObject("Excel.Application", { dynamic: true });
using books = excel.Workbooks; // property: synchronous
using book = await books.Add(); // method: on the COM executor
for await (const sheet of book.Worksheets) {
  using owned = sheet; // fetched on the executor
  console.log(owned.Name);
}
await run(book.SaveAs("C:\\Temp\\a.xlsx"), { signal: AbortSignal.timeout(10_000) });
```

## Objects and apartments

Objects created via `CreateObject`, `GetObject` or `GetActiveObject` are created on the COM executor so their methods never block the JavaScript thread. The JavaScript thread holds its own reference valid in the home apartment:

- **Multithreaded home** (`apartment(COINIT_MULTITHREADED)`) and **agile objects** (`IAgileObject`, free-threaded marshaler) share one reference directly with the executor — zero marshaling overhead.
- **Single-threaded home** (`apartment()`): objects are handed over through the Global Interface Table so properties run on the home thread through a proxy while methods run on the executor. This requires interfaces that can be marshaled (`IDispatch`, `[oleautomation]`/dual interfaces with a registered type library, or interfaces with a proxy/stub). Dual interfaces whose vtable cannot be marshaled automatically fall back to `IDispatch` with the same members and DISPIDs.
- **Home-thread pointers** (`CoCreateInstance`, `QueryInterface` of an STA object, or JavaScript `COMObject`s with `_threading_ = "apartment"`) run on the home thread while keeping the same `Promise` return shape. Calls passing such pointers as arguments run on the home thread too.
- **Interface results** (out pointers, `VARIANT`s, `SAFEARRAY`s of COM objects) are handed over to the home apartment before the promise settles.

### Creating at home

Pass `{ home: true }` to `CreateObject`, `GetObject` or `GetActiveObject` to create or bind the object directly on the JavaScript thread (no marshaling, no executor). In a single-threaded home, use `{ home: true }` for objects whose interfaces cannot be marshaled (such as ADSI `IDirectorySearch`), or initialize the home thread as MTA (`apartment(COINIT_MULTITHREADED)`). Classes registered on the current thread with `registerClassObject` are created at home automatically.

## Synchronous calls

`sync(pointer)` (or `sync(dynamicObject)`) returns a view of the same object whose methods run synchronously on the JavaScript thread, sharing its underlying reference (releasing either releases both).

Use `sync()` in:

- event handlers and `COMObject` implementations;
- code that must not interleave with other JavaScript turns;
- tight loops over fast in-memory calls where executor round-trip latency adds up (e.g. reading columns from an already-fetched ADSI row in [examples/active-directory.ts](../examples/active-directory.ts), while keeping network paging `GetNextRow` asynchronous).

Each asynchronous call pays one executor round trip (~20 µs with conversions vs ~2–4 µs synchronously).

```typescript
import { CreateObject, sync } from "@node-ctypes/comtypes";

using dict = await CreateObject("Scripting.Dictionary");
sync(dict).Add("a", 1); // returns S_OK synchronously, not a Promise
```

Generated modules type both `XMethods` (promises) for pointers and `XSyncMethods` for `sync()` views.

## Cancellation and executor lifecycle

### Cancellation with `run()`

`run(promise, { signal })` awaits a COM operation with an `AbortSignal`:

```typescript
import { run } from "@node-ctypes/comtypes";

await run(slowObject.Compute(), { signal: AbortSignal.timeout(5_000) });
```

On abort, `run()` rejects with `signal.reason` and cancels any calls of that operation still queued on the executor. A native call that has already started runs to completion because COM does not interrupt in-flight calls.

### Executor lifetime (`stats()` and `finalize()`)

- The executor thread starts lazily on the first asynchronous call, keeps the event loop alive only while calls are pending, and stops automatically on normal `beforeExit`.
- `await finalize()` drains all accepted calls, releases internal executor resources (including the executor's GIT reference), calls `CoUninitialize` on the executor thread, and stops it. A subsequent async call transparently starts a fresh executor.
- `stats()` returns `{ running, pending, completed, generation }`.

## Electron and mirror mode

In Electron, the main process's JavaScript thread is Chromium's UI thread, which is already initialized as an STA (`CoInitializeEx(COINIT_MULTITHREADED)` fails there with `RPC_E_CHANGED_MODE`). Calling `apartment()` preserves that STA:

- Methods run on the COM executor, keeping the UI responsive.
- Chromium already pumps Win32 messages, so events arrive without calling `startMessagePump()`.
- Keep synchronous property reads out of tight loops on slow out-of-process objects (they go through a COM proxy).
- Electron disables V8 external ArrayBuffers (`NODE_CTYPES_NO_EXTERNAL_BUFFERS=1`, mirror mode). `@node-ctypes/comtypes` synchronizes mirrored buffers automatically across `GUID`, `VARIANT`, `IDispatch`, structures, `SAFEARRAY`s and `CoCreateInstanceEx`, and runs its full test suite under `npm run test:mirror`.

[tests/electron/main.mjs](../tests/electron/main.mjs) verifies this end-to-end inside Electron (`electron tests/electron/main.mjs` after a build).
