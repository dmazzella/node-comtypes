---
title: Type libraries
---

# Type libraries

## Loading at runtime (`GetModule`)

`GetModule(source)` loads a COM type library and builds every type in memory — enums (frozen objects), records (`Structure` subclasses), vtable and dual interfaces (`COMMETHOD`), dispinterfaces (`DISPMETHOD` / `DISPPROPERTY`), coclasses (`CoClass`), aliases and module constants:

```typescript
import { CreateObject, GetModule } from "@node-ctypes/comtypes";

const Scripting = (await GetModule("scrrun.dll")) as any;
using fso = await CreateObject(Scripting.FileSystemObject); // default interface IFileSystem3, via vtable
using temp = await fso.GetSpecialFolder(Scripting.SpecialFolderConst.TemporaryFolder); // IFolder
console.log(temp.Path);
```

- Built-in interfaces (`IUnknown`, `IDispatch`, `IEnumVARIANT`, ...) are reused automatically; mutual references between interfaces are resolved lazily.
- Results are cached per `(LIBID, major, minor, lcid)`.
- **Non-blocking sliced build**: `await GetModule()` reads and builds large libraries (such as `mshtml.tlb`) in slices of a few milliseconds, yielding to the event loop between slices so timers, I/O and Electron UI stay responsive. Concurrent calls for the same library share a single build.
- `getModuleSync(source)` performs the same work synchronously when code cannot `await` (sharing the same cache).
- `CreateObject`, `GetObject` and `GetActiveObject` pre-load the type library of the object's best interface asynchronously before returning it, loading registered libraries in-process rather than walking `ITypeInfo` call-by-call over a proxy.

### Supported sources

`GetModule` and `loadTypeLib` accept:

- a file path or resource spec (`"scrrun.dll"`, `"C:\\path\\foo.tlb"`, `"file.dll\\2"`);
- a registered LIBID (`"{...}"` or `{ libid, major, minor, lcid }`, defaulting to the highest registered version);
- a ProgID (`"Scripting.Dictionary"`, resolved through the registry without instantiating the object);
- an existing `ITypeLib` pointer.

Lower-level building blocks in `@node-ctypes/comtypes/typelib`: `loadTypeLib`, `typeLibOf(object)`, `readTypeLib` (produces a JSON-serializable `TypeLibModel`) and `buildModule(model)`.

## Generated modules

For full static TypeScript types and autocompletion, generate a typed module ahead of time using the CLI:

```sh
npx node-comtypes generate scrrun.dll --out src/com                # writes src/com/Scripting.ts
npx node-comtypes generate "Excel.Application" --out src/com --js  # writes Excel.js + Excel.d.ts (needs typescript)
npx node-comtypes info msxml6.dll                                  # lists types in a library
npx -p @node-ctypes/comtypes node-comtypes info scrrun.dll         # without installing locally
```

Then import the generated module directly:

```typescript
import { CreateObject, sync } from "@node-ctypes/comtypes";
import * as Scripting from "./com/Scripting.ts";

using dict = await CreateObject(Scripting.Dictionary); // InterfacePointer<typeof Scripting.IDictionary>
dict.CompareMode = Scripting.CompareMethod.TextCompare;
const count: number = dict.Count; // read-only: assigning is a compile-time error
const exists: boolean = await dict.Exists("a");
const existsNow: boolean = sync(dict).Exists("a");
```

A generated module embeds the library's `TypeLibModel`, builds it with `buildModule` at import time, and declares static types for:

- frozen enum objects with literal member values and union types;
- record field interfaces;
- `XMethods` (async promise signatures) and `XSyncMethods` (synchronous `sync()` view signatures) for every interface and dispinterface, including read-only properties, `SELF` and cross-interface references;
- typed coclasses and module constants.

### Cache directory (`comtypes.gen` equivalent)

Pass `{ cacheDir }` to `GetModule(source, { cacheDir })` — or set the `NODE_COMTYPES_CACHE` environment variable — to persist generated `<Library>.ts` files on first load. Programmatic generation uses `generateModule(model)` and `writeModule(model, directory)`.

## Creating and registering type libraries

You can emit binary `.tlb` files programmatically without MIDL using `createTypeLib(file)` (Python: `comtypes.typeinfo.CreateTypeLib`), backed by `ICreateTypeLib2` and `ICreateTypeInfo`:

- Call `SaveAllChanges()` on the returned `ICreateTypeLib2` to write the `.tlb` file.
- Register it with `registerTypeLib(file, { scope })` (`"user"` by default; `"machine"` requires administrator privileges), which returns its identity `{ libid, major, minor, lcid }`.
- Unregister it later with `unregisterTypeLib(identity, { scope })`.

Registering a type library enables the Windows type-library marshaler for `[oleautomation]` and dual interfaces across apartments and processes (see [COM servers](servers.md#registration)).
