---
title: Apartments and lifetime
---

# Apartments and lifetime

## Apartments

`CoInitializeEx()` defaults to `COINIT_APARTMENTTHREADED` (STA), matching Python `comtypes`; Shell and most automation objects require it. `apartment()` wraps `CoInitializeEx` and `CoUninitialize` in a `Disposable`:

```typescript
import { apartment, COINIT_MULTITHREADED, getApartmentType } from "@node-ctypes/comtypes";

{
  using _ = apartment(); // STA (COINIT_APARTMENTTHREADED)
  console.log(getApartmentType()?.name); // "STA"
}

{
  using _ = apartment(COINIT_MULTITHREADED); // MTA
  console.log(getApartmentType()?.name); // "MTA"
}
```

The JavaScript thread's apartment is the **home** apartment; asynchronous method calls run on the dedicated COM executor thread (see [Execution model](execution-model.md)). The Node.js event loop does not pump Win32 messages by default: calls into in-process STA objects work directly, while incoming cross-apartment calls and events delivered through window messages need a message pump (see [Events and message pump](events-and-pump.md)).

## Pointers (`COMPointer`)

`COMPointer` (Python: `POINTER(IUnknown)`) owns **one** COM reference:

| Operation                                    | Effect                                                                                  |
| -------------------------------------------- | --------------------------------------------------------------------------------------- |
| `CoCreateInstance`, `QueryInterface`         | Returns a new owned pointer                                                             |
| `tryQueryInterface(iface)`                   | Returns an owned pointer, or `null` on `E_NOINTERFACE`                                  |
| `Release()` / `using` / `[Symbol.dispose]()` | Releases the owned reference; the pointer becomes unusable                              |
| `detach()`                                   | Hands the owned reference to the caller without calling `Release()`                     |
| `COMPointer.fromAddress(p, iface)`           | Adopts an existing reference (e.g. an out parameter); pass `{ addRef: true }` to borrow |
| `AddRef()` / `callMethod("Release")`         | Raw reference counting, independent of `COMPointer` ownership                           |
| `isSameObject(other)`                        | COM identity comparison through `IUnknown` (fast-pathed when addresses match)           |

Because JavaScript has no deterministic destructor semantics outside `using`, always release pointers with `using` or `.Release()`. A `FinalizationRegistry` acts as a safety net for forgotten pointers and never runs after the last `CoUninitialize` on the thread — `CoUninitialize` releases every pointer still alive on that thread and invalidates them first.

Use `getLivePointerCount()` (and `getLiveObjectCount()` for JS `COMObject` servers) in tests to detect reference leaks.

## Threads and marshaling

Each Node.js `worker_threads` worker has its own COM apartment state. A raw COM pointer belongs to its creating apartment: pass it to another thread marshaled.

```typescript
import { Worker } from "node:worker_threads";
import { COMPointer, GlobalInterfaceTable, IDispatch, apartment, marshalInterface, startMessagePump, unmarshalDispatch } from "@node-ctypes/comtypes";

declare const object: COMPointer; // or a dispatch object, or a COMObject

// Owner thread (STA: incoming calls arrive as window messages, so pump).
using pump = startMessagePump({ mode: "wait" });
const token = marshalInterface(object, IDispatch);
new Worker("./worker.ts", { workerData: { token } });

// Inside worker.ts:
// using _ = apartment();
// using proxy = unmarshalDispatch(workerData.token);
// await proxy.Add(1, 2);
```

### Single-use stream tokens vs Global Interface Table

- **`marshalInterface(object, iface?)`** (`CoMarshalInterThreadInterfaceInStream`) returns a single-use `MarshaledInterface` token that can be cloned across `postMessage`. Consume it exactly once with `unmarshalInterface(token, IFoo)` or `unmarshalDispatch(token)`, or discard an unused token with `releaseMarshaled(token)`.
- **`GlobalInterfaceTable`** (`IGlobalInterfaceTable`) registers an interface once and lets any thread obtain proxies as many times as needed with its numeric cookie:

```typescript
using table = new GlobalInterfaceTable();
const cookie = table.register(object, IDispatch);
// From any thread:
using proxy = table.getDispatch(cookie); // or table.get(cookie, IFoo)
table.revoke(cookie);
```

`IDispatch` always crosses apartments. Custom vtable interfaces need a registered proxy/stub or type library (see [COM servers](servers.md#registration)).
