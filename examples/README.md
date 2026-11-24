---
title: Examples
---

# Examples

Runnable TypeScript examples built on `@node-ctypes/comtypes`. Build the package once (`npm run build`), then execute any example directly with Node.js 24+:

```sh
npm run build
node examples/automation.ts
```

| Example                                    | What it demonstrates                                                                                                                     |
| ------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------- |
| [shortcut.ts](shortcut.ts)                 | Creating and reading a Windows `.lnk` shell shortcut through `IShellLinkW` and `IPersistFile`                                            |
| [automation.ts](automation.ts)             | `Scripting.Dictionary`, `Scripting.FileSystemObject` (typed from their type library, `named()`, `putProperty`) and WMI (`dynamic: true`) |
| [stream.ts](stream.ts)                     | In-memory `IStream` (`CreateStreamOnHGlobal`) and `COMMETHOD` output parameters                                                          |
| [server.ts](server.ts)                     | JavaScript `COMObject` with events (`_outgoing_interfaces_`, `fireEvent`), called through vtable, `IDispatch` and native code            |
| [events.ts](events.ts)                     | WMI process-event subscription via `GetEvents` and the low-latency `waitAndPumpMessages` / `startMessagePump({ mode: "wait" })` pump     |
| [localserver.ts](localserver.ts)           | Out-of-process `LocalServer32` COM server registered with `UseCommandLine` (`/regserver`, `/unregserver`, `-Embedding`)                  |
| [active-directory.ts](active-directory.ts) | Active Directory paged search via ADSI `IDirectorySearch` with `activeds.tlb` structs/enums, async paging and fast `sync()` row reads    |
| [typelib.ts](typelib.ts)                   | Loading type libraries dynamically with `GetModule` and instantiating typed coclasses                                                    |

For conceptual walkthroughs and full API guides, see [Guides](../guides/README.md).
