---
title: COM servers
---

# COM servers

## Implementing COM objects in JavaScript (`COMObject`)

`COMObject` (Python: `comtypes.COMObject`) builds native vtables whose slots invoke JavaScript methods:

```typescript
import { BSTR, COMInterface, COMMETHOD, CreateObject, Dispatch, HRESULT, IDispatch, VARIANT, dispid } from "@node-ctypes/comtypes";
import { COMObject } from "@node-ctypes/comtypes/server";

const ICounter = COMInterface({
  name: "ICounter",
  iid: "{...}",
  extends: IDispatch, // dual: vtable and late binding
  methods: [
    COMMETHOD([dispid(1)], HRESULT, "Increment", [["in", "optional"], VARIANT, "step"], [["out", "retval"], VARIANT, "value"]),
    COMMETHOD([dispid(2), "propget"], HRESULT, "Label", [["out", "retval"], BSTR, "label"]),
    COMMETHOD([dispid(2), "propput"], HRESULT, "Label", [["in"], BSTR, "label"]),
  ],
});

class Counter extends COMObject {
  static _com_interfaces_ = [ICounter];
  Label = "counter"; // propget / propput fall back to the JavaScript property
  value = 0;

  Increment(step = 1) {
    return (this.value += step); // written to the [out, retval] parameter
  }
}

const counter = new Counter();
using pointer = counter.QueryInterface(ICounter); // owned ICounter pointer (AddRef)
await pointer.Increment(5); // called on the JS thread (apartment-threaded object)
using dynamic = Dispatch(counter); // late binding through the generated IDispatch
using dict = await CreateObject("Scripting.Dictionary");
await dict.Add("counter", counter); // pass to native code like any COM object
```

### Method lookup, marshaling and lifetime

- **Method dispatch**: vtable methods resolve by name (`Increment`, `_get_Label`, `_set_Label`, `_setref_Label`); property accessors fall back to a JS property of the same name. Unimplemented methods return `E_NOTIMPL`.
- **Arguments and outputs**:
  - `COMMETHOD` inputs are decoded automatically; return a single value for one `[out]` parameter or an array for multiple `[out]` parameters (`[in, out]` parameters arrive as inputs and are written back from the return value). A `[in] VARIANT*` arrives decoded; other `T*` pointers arrive as raw addresses.
  - `STDMETHOD` slots receive raw `node-ctypes` arguments (pointers as `bigint`) and may return an `HRESULT`.
  - Interface arguments (`COMPointer` / dynamic objects) are borrowed for the duration of the call and released afterwards — call `QueryInterface()` or `AddRef()` to retain them.
- **Late binding (`IDispatch`)**: any `COMObject` implementing an interface derived from `IDispatch` exposes all declared interface members (with their DISPIDs and named parameter positions) plus public JavaScript members (excluding `_`-prefixed names). Set `static _com_default_member_` to answer `DISPID_VALUE`, implement `[Symbol.iterator]()` to answer `DISPID_NEWENUM` via `VariantEnumerator`, or override `_invoke_` / `_ids_of_names_` for custom dispatch. Interfaces loaded from a type library also answer `IDispatch::GetTypeInfo`.
- **Lifetime**: native reference counting pins the `COMObject` in memory while `refCount > 0`. Override `_final_release_()` to run cleanup when the count drops to zero. Use `getLiveObjectCount()` to check for server leaks.
- **Threading**: callbacks always execute on the JavaScript thread.
  - `static _threading_ = "apartment"` (default): calls on the object, and calls passing it as an argument, execute on the home thread instead of the COM executor.
  - `static _threading_ = "free"`: uses thread-safe callbacks so calls from the executor or other native threads marshal back to the JS thread. Note that a `"free"` object deadlocks if another thread calls into it while the JS thread is blocked inside a synchronous COM call waiting for that thread.

## COM aggregation and delegation

`COMObject` supports standard COM aggregation and selective interface delegation (useful for composing native COM/WinRT classes such as WinUI 3 `Application` or custom outer wrappers):

```typescript
import { IUnknown } from "@node-ctypes/comtypes";
import { COMObject } from "@node-ctypes/comtypes/server";

class Outer extends COMObject {
  static _com_interfaces_ = [IUnknown];
}

const outer = new Outer();
const outerIUnknown = outer.addressOf(IUnknown); // pass as controlling outer to CoCreateInstance / RoActivateInstance

// Forward unimplemented QueryInterface requests to an aggregated inner object:
outer.setAggregateInner(innerPointer, { own: true });

// Or forward a specific interface (by InterfaceDef or GUID string) to a delegate:
outer.delegateInterface(ISomeInterface, delegatePointer, { own: true });
outer.undelegateInterface(ISomeInterface);

// Releases owned inner and delegated references (also runs automatically at refCount === 0):
outer.releaseAggregate();
```

| Method                                       | Purpose                                                                                            |
| -------------------------------------------- | -------------------------------------------------------------------------------------------------- |
| `addressOf(iface?)`                          | Returns the native vtable pointer for `iface` (default: primary interface / `IUnknown`)            |
| `setAggregateInner(inner, { own? })`         | Sets the non-delegating inner `IUnknown` used as fallback for `QueryInterface`                     |
| `delegateInterface(iface, target, { own? })` | Forwards `QueryInterface` for `iface` directly to `target` (queried if needed)                     |
| `undelegateInterface(iface)`                 | Removes a previously registered interface delegate (releasing it if owned)                         |
| `releaseAggregate()`                         | Releases owned inner and delegate references (also called automatically when `refCount` reaches 0) |

## Errors in server implementations

- Throwing an `Error` or `COMError` returns a failure `HRESULT` (`error.hresult`, or `E_FAIL`) and populates `IErrorInfo` (`description`, `source` from `_reg_progid_` or class name, and interface `GUID`). `ISupportErrorInfo` is implemented automatically.
- Through `IDispatch`, exceptions populate `EXCEPINFO` (`DISP_E_EXCEPTION`); throw `new DispatchError(hresult)` to return a dispatch-level HRESULT (`DISP_E_*`) directly.
- Define `static onError(error, object)` to observe unhandled exceptions in server callbacks.
- `throw new ReturnHRESULT(hresult, text?)` (Python: `comtypes.ReturnHRESULT`) returns any `HRESULT` — including success codes such as `S_FALSE` — without invoking `onError`.
- Call `setErrorInfo({ description, source, guid, helpFile, helpContext })` to set thread error info manually (Python: `comtypes.ReportError`).

## Registering and serving LocalServer32 servers

```typescript
import { CLSCTX_INPROC_SERVER, IDispatch } from "@node-ctypes/comtypes";
import { COMObject, registerClassObject, UseCommandLine } from "@node-ctypes/comtypes/server";

class Calculator extends COMObject {
  static _com_interfaces_ = [IDispatch];
  static _reg_clsid_ = "{...}";
  static _reg_progid_ = "My.Calculator.1";
  static _reg_novers_progid_ = "My.Calculator";
  static _reg_desc_ = "Calculator written in JavaScript";

  Add(a: number, b: number) {
    return a + b;
  }
}

// In-process only (no registry writes):
using registration = registerClassObject(Calculator, { clsctx: CLSCTX_INPROC_SERVER });

// LocalServer32 (out-of-process): end the server script with:
await UseCommandLine([Calculator]);
// node calculator.ts /regserver     registers in HKCU\Software\Classes
// node calculator.ts /unregserver   removes the registration
// node calculator.ts -Embedding     launched automatically by COM when a client activates the class
```

See [examples/localserver.ts](../examples/localserver.ts) for a complete server callable from PowerShell or any COM client.

`UseCommandLine` (Python: `comtypes.server.register.UseCommandLine`) dispatches to `registerServer`, `unregisterServer` and `runLocalServer` (which initializes an STA, registers class factories and pumps messages via `waitAndPumpMessages` until idle). Command-line switches accept `/` or `-` case-insensitively (use `-regserver` in Git Bash so paths are not rewritten).

### Registration

`registerServer` writes per-user entries under `HKCU\Software\Classes`:

- `LocalServer32`: the Node.js executable, `process.execArgv`, the script path and any extra arguments;
- `ProgID`, `VersionIndependentProgID` (`_reg_novers_progid_` with `CurVer`), and `_reg_desc_`;
- `TypeLib` and `Version` when the class has a type library (`_reg_typelib_` `[libid, major, minor]`, or inferred from `GetModule` / generated interfaces). The `.tlb` file (`_typelib_path_`, or the loaded file) is also registered unless `{ typelib: false }`.

> **Why registering the type library matters:** Windows marshals `IDispatch` automatically, but **custom `[oleautomation]` and dual vtable interfaces** rely on a registered type library (the standard type-library marshaler) to cross processes and apartments. Without it, `QueryInterface` for custom interfaces fails with `E_NOINTERFACE` on the client. Classes with `_reg_clsid_` and a type library also implement `IProvideClassInfo2` automatically so clients get their best typed interface.
