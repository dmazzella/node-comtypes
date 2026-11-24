# node-comtypes (`@node-ctypes/comtypes`)

[![npm version](https://img.shields.io/npm/v/@node-ctypes/comtypes.svg)](https://www.npmjs.com/package/@node-ctypes/comtypes)
[![npm downloads](https://img.shields.io/npm/dm/@node-ctypes/comtypes.svg)](https://www.npmjs.com/package/@node-ctypes/comtypes)
[![CI](https://github.com/dmazzella/node-comtypes/actions/workflows/ci.yml/badge.svg)](https://github.com/dmazzella/node-comtypes/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Typed Windows COM client and server support for Node.js**, powered by
[node-ctypes](https://www.npmjs.com/package/node-ctypes) and modelled on Python
[comtypes](https://github.com/enthought/comtypes).

- **Node.js 24+**, ESM, TypeScript declarations, zero native addon of its own.
- **Full COM stack**: vtable interfaces (`COMMETHOD` / `STDMETHOD`), automation (`IDispatch`, `VARIANT`, `SAFEARRAY`, `VT_RECORD`), type libraries and code generation, JavaScript `COMObject` servers with aggregation/delegation, events, and `LocalServer32`.
- **Non-blocking by default**: methods return promises and run on a dedicated COM executor thread so slow objects never freeze the event loop or an Electron UI; properties stay synchronous, and `sync()` provides direct synchronous views.

---

## Quick start

```sh
npm install @node-ctypes/comtypes
```

```typescript
import { create_unicode_buffer } from "node-ctypes";
import { apartment, CreateObject, IPersistFile, ShellLink } from "@node-ctypes/comtypes";

using _ = apartment(); // CoInitializeEx(COINIT_APARTMENTTHREADED) ... CoUninitialize()

// Vtable interface: the ShellLink coclass returns an IShellLinkW pointer.
using link = await CreateObject(ShellLink);
await link.SetPath("C:\\Windows\\notepad.exe");
await link.SetDescription("Notepad");
using file = link.QueryInterface(IPersistFile);
await file.Save("C:\\Temp\\notepad.lnk", 1);
const text = create_unicode_buffer(260);
await link.GetPath(text, 260, null, 0);
console.log(text.value);

// Automation: typed automatically from the object's type library.
using dict = await CreateObject("Scripting.Dictionary"); // IDictionary
await dict.Add("a", 1); // methods return promises
console.log(dict.Count, dict.Item("a")); // properties are synchronous
```

Failures reject (or throw, for synchronous members) with a `COMError`. To build from source, see [Getting started](guides/getting-started.md#building-from-source).

## Supported platforms

| Operating system | x64 | ARM64 | Notes                                                                                  |
| ---------------- | --- | ----- | -------------------------------------------------------------------------------------- |
| Windows          | ✅  | ✅    | Full COM runtime (`ole32.dll`, `oleaut32.dll`), tested on Node.js 24 and 26            |
| macOS / Linux    | ✅  | ✅    | Imports cleanly; `GUID`, `HRESULT`, interface definitions and unit tests work anywhere |

Node.js 24 or later. Also verified in Electron and under `NODE_CTYPES_NO_EXTERNAL_BUFFERS=1` (mirror mode).

---

## Guides

- [Getting started](guides/getting-started.md) — installation, supported platforms, entry points, building from source
- [Apartments and lifetime](guides/apartments-and-lifetime.md) — STA/MTA apartments, `COMPointer` ownership, worker threads and marshaling
- [Execution model](guides/execution-model.md) — the COM executor, `sync()`, cancellation, Electron and mirror mode
- [Automation and VARIANTs](guides/automation.md) — `CreateObject`, `GetObject`, `Constants`, dynamic objects, `VARIANT` and `SAFEARRAY` conversions
- [Defining interfaces](guides/interfaces.md) — `COMMETHOD`, dispinterfaces (`DISPMETHOD` / `DISPPROPERTY`), `STDMETHOD`, built-in interfaces
- [Type libraries](guides/typelib.md) — `GetModule`, code generation (`node-comtypes generate`), creating and registering type libraries
- [Events and message pump](guides/events-and-pump.md) — `GetEvents`, `ShowEvents`, `COMObject` event sources, `waitAndPumpMessages` and `startMessagePump`
- [COM servers](guides/servers.md) — `COMObject`, late binding, COM aggregation and delegation, `LocalServer32` registration
- [Errors and limits](guides/errors.md) — `COMError`, `IErrorInfo`, `EXCEPINFO`, Visual Basic errors, compatibility and limits
- [Python comtypes correspondence](guides/python-comtypes.md) — Python `comtypes` API mapped to `node-comtypes`
- [Building, testing and internals](guides/contributing.md) — scripts, mirror-mode testing, CI and release workflow

Detailed internal flows (in Italian): [FLOWS.md](FLOWS.md).

---

## Examples

Runnable TypeScript examples live in `examples/`; see the [index](examples/README.md).

- [shortcut.ts](examples/shortcut.ts) — shell links through `IShellLinkW` and `IPersistFile`
- [automation.ts](examples/automation.ts) — `Scripting.Dictionary`, `FileSystemObject` and WMI
- [stream.ts](examples/stream.ts) — `IStream` and `COMMETHOD` outputs
- [server.ts](examples/server.ts) — JavaScript `COMObject` with events, called through vtable, `IDispatch` and native code
- [events.ts](examples/events.ts) — WMI event subscription with `GetEvents` and the low-latency `waitAndPumpMessages` pump
- [localserver.ts](examples/localserver.ts) — out-of-process `LocalServer32` server with `UseCommandLine`
- [active-directory.ts](examples/active-directory.ts) — Active Directory search via ADSI `IDirectorySearch` and `activeds.tlb`
- [typelib.ts](examples/typelib.ts) — `GetModule` and typed coclasses

---

## License

[MIT](LICENSE). See [NOTICE](NOTICE) for the Python `comtypes` attribution.
