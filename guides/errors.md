---
title: Errors and limits
---

# Errors and limits

## `COMError` and HRESULT helpers

Every failing COM call rejects (for async methods) or throws (for synchronous calls and properties) a `COMError`:

- `error.hresult` — signed 32-bit `HRESULT` (e.g. `E_NOINTERFACE`, `DISP_E_EXCEPTION`, `REGDB_E_CLASSNOTREG`)
- `error.text` — human-readable Win32 system text from `FormatMessageW` (also available directly via `systemMessage(hresult)`)
- `error.method` — failing call label (e.g. `IShellLinkW::SetPath`, `IDispatch::Invoke(Add)`)
- `error.details` — structured error details when provided by the server:
  - **Vtable calls**: `IErrorInfo` (`description`, `source`, `helpFile`, `helpContext`, `guid`), queried only when the object confirms support for the failing interface via `ISupportErrorInfo` (matching Python `comtypes`).
  - **`IDispatch::Invoke`**: for `DISP_E_EXCEPTION`, the `EXCEPINFO` fields and `scode` (servers such as the Scripting runtime often populate only `scode`, whose text is resolved into the message); for `DISP_E_TYPEMISMATCH` and `DISP_E_PARAMNOTFOUND`, `argumentIndex`.

Visual Basic run-time errors (`FACILITY_CONTROL`, `0x800Axxxx`, e.g. from `Scripting.Dictionary`) resolve their descriptions through `vbscript.dll` when that Windows Feature on Demand is installed.

Useful helpers exported from `@node-ctypes/comtypes`: `checkHRESULT(hr, method?)`, `SUCCEEDED(hr)`, `FAILED(hr)`, `hresultHex(hr)`, `hresultName(hr)`, `systemMessage(hr)`.

## Compatibility and limits

- **Single FIFO executor thread**: asynchronous calls run sequentially on one MTA executor thread. One slow call delays subsequent queued calls (without blocking the JavaScript event loop). Each async call pays one executor round trip; use [`sync()`](execution-model.md#synchronous-calls) in tight loops of fast local calls.
- **Win32 message pump**: Node.js does not pump Win32 messages automatically; start a pump ([Events and message pump](events-and-pump.md#win32-message-pump)) when waiting for out-of-process events or serving STA calls.
- **ByRef / raw VARIANT pointers across apartments**: COM objects inside `ByRef` arguments or `VARIANT*` buffers passed by address are not automatically handed over across apartments. In an STA home, pass them to objects created with `{ home: true }` or call via `sync()`.
- **Non-thread-safe hand-rolled vtables**: raw hand-rolled COM vtables (outside `COMObject`) that are not thread-safe must not be passed to executor calls from an MTA home.
- **Typed vs dynamic members**: objects returned without `{ dynamic: true }` expose only the members declared in their type library. Dynamic properties added at runtime (such as WMI class properties or `IDispatchEx`) require `{ dynamic: true }`. Accessing an unknown member on a typed pointer evaluates to `undefined` (whereas Python raises `AttributeError`), and typed pointers are not directly callable for `DISPID_VALUE`.
- **Properties with only optional arguments**: on dynamic objects, a property whose getter has only optional arguments is read as a value; use `getProperty(obj, "Name", arg)` to pass arguments.
- **`named()`**: resolved at runtime only; static TypeScript signatures remain positional.
- **C-style `size_is` arrays**: have no automatic `COMMETHOD` codec; declare them with `STDMETHOD` or `POINTER(T)` and manage the buffer explicitly.
- **`VT_RECORD`**: requires a record type that carries a GUID in its type library (a Windows `IRecordInfo` requirement).
- **Custom interface marshaling**: custom vtable interfaces cross apartments/processes only with a registered type library or proxy/stub; `IDispatch` always crosses apartments.
- **InprocServer32**: in-process DLL servers cannot directly host a Node.js runtime, so `_reg_threading_` and `_reg_clsctx_` have no equivalent; register out-of-process servers via `LocalServer32` (`UseCommandLine`).
- **`_threading_ = "free"` deadlock**: a `"free"` `COMObject` deadlocks if another thread invokes it while the JS thread is blocked inside a synchronous COM call waiting on that thread.
- **Objects without `ISupportErrorInfo`**: components such as the Scripting runtime and Shell do not implement `ISupportErrorInfo`, so their vtable failures carry no `error.details`.
