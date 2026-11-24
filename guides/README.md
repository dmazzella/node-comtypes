---
title: Guides
children:
  - ./getting-started.md
  - ./apartments-and-lifetime.md
  - ./execution-model.md
  - ./automation.md
  - ./interfaces.md
  - ./typelib.md
  - ./events-and-pump.md
  - ./servers.md
  - ./errors.md
  - ./python-comtypes.md
  - ./contributing.md
---

# Guides

- [Getting started](getting-started.md) — installation, supported platforms, entry points, building from source
- [Apartments and lifetime](apartments-and-lifetime.md) — STA/MTA apartments, `COMPointer` ownership, worker threads and marshaling
- [Execution model](execution-model.md) — the COM executor, `sync()`, cancellation, Electron and mirror mode
- [Automation and VARIANTs](automation.md) — `CreateObject`, `GetObject`, `Constants`, dynamic objects, `VARIANT` and `SAFEARRAY` conversions
- [Defining interfaces](interfaces.md) — `COMMETHOD`, dispinterfaces (`DISPMETHOD` / `DISPPROPERTY`), `STDMETHOD`, built-in interfaces
- [Type libraries](typelib.md) — `GetModule`, code generation (`node-comtypes generate`), creating and registering type libraries
- [Events and message pump](events-and-pump.md) — `GetEvents`, `ShowEvents`, `COMObject` event sources, `waitAndPumpMessages` and `startMessagePump`
- [COM servers](servers.md) — `COMObject`, late binding, COM aggregation and delegation, `LocalServer32` registration
- [Errors and limits](errors.md) — `COMError`, `IErrorInfo`, `EXCEPINFO`, Visual Basic errors, compatibility and limits
- [Python comtypes correspondence](python-comtypes.md) — Python `comtypes` API mapped to `node-comtypes`
- [Building, testing and internals](contributing.md) — scripts, mirror-mode testing, CI and release workflow

The [examples](../examples/README.md) are runnable programs built on these APIs.
