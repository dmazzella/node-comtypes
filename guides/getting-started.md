---
title: Getting started
---

# Getting started

## Installation

```sh
npm install @node-ctypes/comtypes
```

`@node-ctypes/comtypes` requires **Node.js 24 or later**, ESM, and has no native addon of its own — only [node-ctypes](https://www.npmjs.com/package/node-ctypes). The [Quick start](../README.md#quick-start) shows a first program.

## Supported platforms

| Operating system | x64 | ARM64 | Runtime support                                                                       |
| ---------------- | --- | ----- | ------------------------------------------------------------------------------------- |
| Windows          | ✅  | ✅    | Full COM client, automation, type libraries, events and servers (`ole32`, `oleaut32`) |
| macOS            | ✅  | ✅    | Imports cleanly; `GUID`, `HRESULT`, interface definitions and unit tests work         |
| Linux            | ✅  | ✅    | Imports cleanly; `GUID`, `HRESULT`, interface definitions and unit tests work         |

COM runtime calls are Windows-only. Keeping non-Windows imports working lets shared TypeScript libraries, declarations and unit tests run on every OS.

## First steps

Always initialize a COM apartment on the current thread before creating or calling COM objects, and release owned pointers deterministically with `using` (or `Release()`):

```typescript
import { create_unicode_buffer } from "node-ctypes";
import { apartment, CreateObject, IPersistFile, ShellLink } from "@node-ctypes/comtypes";

using _ = apartment(); // CoInitializeEx(COINIT_APARTMENTTHREADED) ... CoUninitialize()

// Vtable interface: the ShellLink coclass gives an IShellLinkW pointer.
using link = await CreateObject(ShellLink);
await link.SetPath("C:\\Windows\\notepad.exe");
await link.SetDescription("Notepad");
using file = link.QueryInterface(IPersistFile);
await file.Save("C:\\Temp\\notepad.lnk", 1);

const text = create_unicode_buffer(260);
await link.GetPath(text, 260, null, 0);
console.log(text.value);
```

Automation objects (such as `Scripting.Dictionary`, Office applications, or WMI) return their best typed interface from their type library automatically:

```typescript
import { apartment, CreateObject } from "@node-ctypes/comtypes";

using _ = apartment();

using dict = await CreateObject("Scripting.Dictionary"); // IDictionary
await dict.Add("a", 1); // methods return promises (run on the COM executor)
console.log(dict.Count, dict.Item("a")); // properties are synchronous
```

Failures reject (or throw, for synchronous members) with a `COMError`. Raw methods (`STDMETHOD`) resolve to the HRESULT (`S_OK`, `S_FALSE`); methods declared with `COMMETHOD` resolve to their output parameters.

## Entry points

The root entry point `@node-ctypes/comtypes` re-exports everything except `low-level`. Focused subpath exports are available when you want narrower imports:

| Subpath                            | Exports                                                                                                                                                                                                                                                                                                      |
| ---------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `@node-ctypes/comtypes`            | Core (`GUID`, `HRESULT`, `COMError`, `COMInterface`, `COMMETHOD`, `STDMETHOD`, `DISPMETHOD`, `DISPPROPERTY`, `COMPointer`, `apartment`, `CoCreateInstance(Ex)`, message pump, threads, `sync`, `run`, `stats`, `finalize`) plus all non-low-level subpaths                                                   |
| `@node-ctypes/comtypes/client`     | `CreateObject`, `GetObject`, `GetActiveObject`, `sync`, `GetBestInterface` / `wrap`, `Constants`, `Dispatch`, dynamic helpers (`getProperty`, `putProperty`, `putRefProperty`, `invokeMethod`, `enumerate`, `dispatchPointer`, `isDispatch`, `release`, `memberInfo`, `typeName`), `GetEvents`, `ShowEvents` |
| `@node-ctypes/comtypes/automation` | `VARIANT`, `typed`, `safeArray`, `byRef` / `ByRef`, `named`, `missing`, `ErrorValue`, `record` / `RecordInfo`, SAFEARRAY helpers, `DECIMAL`, dates, `DISPPARAMS`, `EXCEPINFO`, `invoke` / `getIDsOfNames`, `setDefaultLCID`, type info helpers, `VT_*` / `DISPID_*` constants                                |
| `@node-ctypes/comtypes/typelib`    | `GetModule`, `getModuleSync`, `loadTypeLib`, `typeLibOf`, `readTypeLib`, `buildModule`, `generateModule`, `writeModule`, `createTypeLib`, `registerTypeLib`, `unregisterTypeLib`, `isCoClass` / `CoClass`                                                                                                    |
| `@node-ctypes/comtypes/server`     | `COMObject`, `VariantEnumerator`, `ConnectionPoint`, `ClassFactory`, `DispatchError`, `ReturnHRESULT`, `registerClassObject`, `registerServer`, `unregisterServer`, `runLocalServer`, `UseCommandLine`, `REGCLS_*`, `setErrorInfo`, `getLiveObjectCount`                                                     |
| `@node-ctypes/comtypes/interfaces` | Built-in COM interfaces (`IUnknown`, `IDispatch`, `IEnumVARIANT`, `ITypeInfo`, `ITypeLib`, `IShellLinkW`, `IStream`, `IConnectionPoint`, `IPersistFile`, ...) and their structures/constants                                                                                                                 |
| `@node-ctypes/comtypes/low-level`  | Lazy `ole32`/`oleaut32` bindings, `executorCall`, raw vtable helpers (`rawQueryInterface`, `rawAddRef`, `rawRelease`), BSTR/task-memory/mirror helpers, registry helpers                                                                                                                                     |

## Building from source

From a clone of this repository (Node.js 24+):

```sh
npm install
npm run build
npm test
```

See [Building, testing and internals](contributing.md) for mirror-mode tests, type tests, Electron checks and docs generation.

## Next steps

- [Apartments and lifetime](apartments-and-lifetime.md): STA/MTA, pointer ownership, worker threads
- [Execution model](execution-model.md): asynchronous COM executor, `sync()`, cancellation, Electron
- [Automation and VARIANTs](automation.md): `CreateObject`, dynamic objects, `VARIANT` and `SAFEARRAY`
- [Defining interfaces](interfaces.md): `COMMETHOD`, dispinterfaces, `STDMETHOD`
- [Type libraries](typelib.md): `GetModule` and `node-comtypes generate`
