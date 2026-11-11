# node-comtypes (`@node-ctypes/comtypes`)

Typed Windows COM support for Node.js powered by [node-ctypes](https://www.npmjs.com/package/node-ctypes), modelled on Python [comtypes](https://github.com/enthought/comtypes).

- **Node.js 24+**, ESM, TypeScript declarations.
- Windows only at runtime (`ole32.dll`, `oleaut32.dll`). The package imports on Linux and macOS, and GUID/HRESULT/interface definitions work there, so shared code and unit tests run everywhere.
- Independent package: no native addon of its own, only `node-ctypes`.
- Client and server side: vtable interfaces with IDL flags, automation (`IDispatch`, `VARIANT`, `SAFEARRAY`), type libraries and code generation, COM objects written in JavaScript, events, LocalServer32 servers.
- **Asynchronous calls**: methods return promises and run on a dedicated COM thread, so slow objects never block the event loop (or an Electron UI); properties stay synchronous. See [Execution model](#execution-model).

```sh
npm install @node-ctypes/comtypes
```

To build from source, run from this directory:

```sh
npm install
npm run build
npm test
npm pack
```

## Quick start

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

// Automation: typed from the type library of the object.
using dict = await CreateObject("Scripting.Dictionary"); // IDictionary
await dict.Add("a", 1); // methods return promises
console.log(dict.Count, dict.Item("a")); // properties are synchronous
```

Failures reject (or throw, for synchronous members) with a `COMError`. Raw methods (`STDMETHOD`) resolve to the HRESULT (`S_OK`, `S_FALSE`); methods declared with `COMMETHOD` resolve to their output parameters.

## Apartments and lifetime

### Apartments

`CoInitializeEx()` defaults to `COINIT_APARTMENTTHREADED`, like comtypes; Shell and most automation objects require it. `apartment()` initializes and uninitializes with `using`. This is the apartment of the JavaScript thread, the **home** apartment; method calls run on the COM executor (see [Execution model](#execution-model)). The Node.js event loop does not pump Win32 messages: calls into in-process STA objects work directly, while incoming calls and events delivered through window messages need a pump (see [Message pump](#message-pump)).

### Pointers

`COMPointer` (Python: `POINTER(IUnknown)`) owns exactly one reference:

| Operation                                    | Effect                                                                             |
| -------------------------------------------- | ---------------------------------------------------------------------------------- |
| `CoCreateInstance`, `QueryInterface`         | New owned pointer                                                                  |
| `tryQueryInterface(iface)`                   | Owned pointer, or `null` on `E_NOINTERFACE`                                        |
| `Release()` / `using` / `[Symbol.dispose]()` | Releases the owned reference; the pointer becomes unusable                         |
| `detach()`                                   | Gives the reference to the caller without releasing it                             |
| `COMPointer.fromAddress(p, iface)`           | Adopts an existing reference (e.g. an out parameter); `{ addRef: true }` to borrow |
| `AddRef()` / `callMethod("Release")`         | Raw reference counting, not tied to ownership                                      |
| `isSameObject(other)`                        | COM identity through `IUnknown`                                                    |

JavaScript has no reference counting, so release pointers deterministically with `using` or `Release()`. A `FinalizationRegistry` releases forgotten pointers as a safety net. It never runs after the last `CoUninitialize` of the thread, which releases every pointer still alive and invalidates it. `getLivePointerCount()` helps detect leaks.

### Threads

Each worker thread has its own apartment state, and a COM pointer belongs to its apartment: pass it to another thread marshaled.

```typescript
import { Worker } from "node:worker_threads";
import { COMPointer, GlobalInterfaceTable, IDispatch, marshalInterface, startMessagePump } from "@node-ctypes/comtypes";

declare const object: COMPointer; // or a dispatch object, or a COMObject

// Owner thread (an STA: incoming calls arrive as window messages, so pump).
using pump = startMessagePump();
new Worker("./worker.ts", { workerData: { token: marshalInterface(object, IDispatch) } });
// worker.ts: using _ = apartment(); using proxy = unmarshalDispatch(workerData.token); await proxy.Add(1, 2);

// Or register once and get a proxy in any thread, any number of times:
using table = new GlobalInterfaceTable();
const cookie = table.register(object, IDispatch); // table.get(cookie, IFoo) / table.getDispatch(cookie)
```

`marshalInterface` (`CoMarshalInterThreadInterfaceInStream`) returns a single-use token that can be cloned with `postMessage`. `unmarshalInterface(token, IFoo)` and `unmarshalDispatch(token)` consume it; `releaseMarshaled(token)` discards it. `IDispatch` always crosses apartments; custom vtable interfaces need a registered proxy/stub or type library (see [Servers](#servers)).

## Execution model

**Calls are asynchronous; live access is synchronous.** Methods return promises and run on the **COM executor**, a native thread owned by node-comtypes (a node-ctypes `Executor`, not a libuv worker) in the multithreaded apartment, one call at a time in FIFO order. The event loop stays free while a call runs, however slow the object (an Office application, a remote server, WMI, an application that does not respond).

| Asynchronous (`Promise`)                                               | Synchronous                                                                                                   |
| ---------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------- |
| COMMETHOD / STDMETHOD methods, DISPMETHODs, methods of dynamic objects | properties: get, put, putref, parameterized (`dict.Item("a")`, `_get_X` / `_set_X`)                           |
| `CreateObject`, `GetObject`, `GetActiveObject`                         | `QueryInterface`, `AddRef`, `Release`, `using`, `GetBestInterface`, `queryService`                            |
| `for await` over collections (`IEnumVARIANT::Next` on the executor)    | `for...of`, `VARIANT` conversions, `Constants`, `getModuleSync`, `GetEvents`, generated modules               |
| `GetModule` (in slices on the JavaScript thread, see below)            |                                                                                                               |
| `executorCall` (low-level)                                             | the server side: `COMObject` implementations, `fireEvent`, registration; [`sync()`](#synchronous-calls) views |

```typescript
import { apartment, CreateObject, run } from "@node-ctypes/comtypes";

using _ = apartment();
using excel = await CreateObject("Excel.Application", { dynamic: true });
using books = excel.Workbooks; // property: synchronous
using book = await books.Add(); // method: on the executor
for await (const sheet of book.Worksheets) {
  using owned = sheet; // fetched on the executor
  console.log(owned.Name);
}
await run(book.SaveAs("C:\\Temp\\a.xlsx"), { signal: AbortSignal.timeout(10_000) });
```

**Objects and apartments.** Objects are created on the executor, so their calls never run on the JavaScript thread. The JavaScript thread keeps its own reference, valid in its home apartment:

- a multithreaded home (`apartment(COINIT_MULTITHREADED)`) and agile objects (`IAgileObject`, the free-threaded marshaler) share one reference with the executor;
- in a single-threaded home, objects are passed through the Global Interface Table: properties are read through a proxy, methods on the executor. This needs interfaces that can be marshaled: `IDispatch`, `[oleautomation]` and dual interfaces with a registered type library, interfaces with a proxy/stub. Dual interfaces that cannot be marshaled are called through `IDispatch` (same members, with their DISPIDs);
- pointers obtained on the JavaScript thread (`CoCreateInstance`, `QueryInterface` of an STA object, JavaScript `COMObject`s with `_threading_ = "apartment"`) are called on the JavaScript thread, synchronously, while the promise keeps the same shape. Calls passing them as arguments run there too;
- interface results (out pointers, `VARIANT`s, `SAFEARRAY`s of objects) are handed over to the home apartment before the promise settles.

`{ home: true }` creates or binds an object on the JavaScript thread instead (direct calls, no marshaling, but on that thread): in a single-threaded home, use it for objects whose interfaces cannot be marshaled, such as ADSI `IDirectorySearch`, or choose a multithreaded home. Classes this thread registered with `registerClassObject` are created at home automatically.

**Cancellation.** `run(promise, { signal })` awaits a COM operation with an `AbortSignal`: on abort it rejects with `signal.reason` and skips the operation's calls still queued on the executor. A call already running completes (COM does not interrupt it).

**Lifetime.** The executor starts at the first asynchronous call, keeps the process alive only while calls are pending, and stops at exit. `await finalize()` drains the accepted calls and stops it (a later call starts a new one); `stats()` returns `{ running, pending, completed, generation }`.

### Synchronous calls

`sync(pointer)` (or `sync(dynamicObject)`) returns a view of the same object whose methods run synchronously on the JavaScript thread, sharing its reference (releasing either releases both): for event handlers, `COMObject` implementations, or code that must not interleave with other JavaScript.

Each asynchronous call is a round trip to the executor thread: a typed method costs about 20 µs with its conversions, against 2–4 µs synchronously. That is negligible for most code; `sync()` is there for loops of many fast calls where it adds up.

```typescript
import { CreateObject, sync } from "@node-ctypes/comtypes";

using dict = await CreateObject("Scripting.Dictionary");
sync(dict).Add("a", 1); // S_OK, not a Promise
```

Generated modules type both: `XMethods` (promises) for pointers, `XSyncMethods` for `sync()` views.

### Electron

The main process's JavaScript thread is Chromium's UI thread, already a single-threaded apartment (`CoInitializeEx(COINIT_MULTITHREADED)` there fails with `RPC_E_CHANGED_MODE`); `apartment()` keeps it. Methods run on the executor, so the UI does not freeze; Chromium pumps window messages, so events arrive without `startMessagePump`. Synchronous property reads go through a proxy to the object's apartment: keep them out of tight loops on slow objects. [tests/electron/main.mjs](./tests/electron/main.mjs) checks this in Electron (`electron tests/electron/main.mjs` after a build).

## Automation

`CreateObject`, `GetObject` and `GetActiveObject` return, as in Python, the object's **best interface** (`GetBestInterface`): the default interface of its coclass (`IProvideClassInfo`), or the interface `IDispatch::GetTypeInfo` describes, built from its type library with `GetModule` (dual interfaces are called through the vtable). An object without type information becomes a dynamic object, or stays an `IUnknown` pointer when it has no `IDispatch`.

```typescript
import { apartment, CreateObject, named, putProperty } from "@node-ctypes/comtypes";

using _ = apartment();

using dict = await CreateObject("Scripting.Dictionary"); // IDictionary, typed from scrrun.dll
dict.CompareMode = 1; // property put (only while the dictionary is empty)
await dict.Add("a", 1); // method
console.log(dict.Count, dict.Item("a")); // property get, parameterized property
putProperty(dict, "Item", "a", 2); // parameterized property put (late bound, by name)
for (const key of dict) console.log(key); // _NewEnum / IEnumVARIANT; for await: on the executor

using fso = await CreateObject("Scripting.FileSystemObject"); // IFileSystem3
using file = await fso.CreateTextFile("C:\\Temp\\a.txt", named({ Overwrite: true })); // named arguments; ITextStream
```

- **Static types**: the result is typed at run time, but its TypeScript type is `AutomationObject` (members of type `any`). For static types, create a coclass of a [generated module](#generated-modules).
- **Results**: `IDispatch` values returned by typed calls (a `VARIANT` holding an object, `Application.ActiveSheet`) get their best interface too (Python: `wrap_outparam`). `GetBestInterface(pointer)` (also `wrap`) does it explicitly and returns a new reference.
- **Iteration and helpers**: typed pointers deriving from `IDispatch` iterate their `_NewEnum` collection, and the late-bound helpers (`getProperty`, `putProperty`, `invokeMethod`, ...) accept them.
- **Options**: `{ interface: IFoo }` returns that interface; `{ dynamic: true }` a [dynamic object](#dynamic-objects); `{ home: true }` creates or binds it on the JavaScript thread ([Execution model](#execution-model)); `CreateObject` also takes `clsctx` and `machine` (a remote computer, through `CoCreateInstanceEx`). `GetActiveObject(progid)` attaches to an object of the Running Object Table, `GetObject(displayName)` binds a moniker (`"winmgmts:"`, a file).

### Constants

`Constants(obj)` (Python: `comtypes.client.Constants`) returns the constants of the object's type library: enum members and module constants by name, and enums by name. Unknown names throw, so typos do not become `undefined`.

```typescript
import { Constants, CreateObject } from "@node-ctypes/comtypes";

using fso = await CreateObject("Scripting.FileSystemObject");
const c = Constants(fso);
using temp = await fso.GetSpecialFolder(c.TemporaryFolder); // or c.SpecialFolderConst.TemporaryFolder
```

### Dynamic objects

`{ dynamic: true }` returns a dynamic object driven by `GetIDsOfNames` + `Invoke` (Python: `dynamic=True`, `comtypes.client.dynamic.Dispatch`). Use it for members that are not in the type library, such as WMI class properties:

```typescript
import { CreateObject, GetObject } from "@node-ctypes/comtypes";

using wmi = await GetObject("winmgmts:", { dynamic: true });
for await (const os of await wmi.ExecQuery("SELECT Caption FROM Win32_OperatingSystem")) console.log(os.Caption);

using dict = await CreateObject("Scripting.Dictionary", { dynamic: true });
await dict.Add("a", 1);
console.log(dict("a")); // default member (DISPID_VALUE), a property: synchronous
```

| Expression                                 | Invoke                                                                   |
| ------------------------------------------ | ------------------------------------------------------------------------ |
| `obj.Name`                                 | `DISPATCH_PROPERTYGET`, or a method caller (see below)                   |
| `obj.Name(...args)`                        | `DISPATCH_METHOD \| DISPATCH_PROPERTYGET` (a `Promise` for methods)      |
| `obj.Name = value`                         | `DISPATCH_PROPERTYPUT` (`PROPERTYPUTREF` for COM objects when supported) |
| `obj(...args)`                             | default member (`DISPID_VALUE`)                                          |
| `for (const x of obj)`, `for await`        | `_NewEnum` (`DISPID_NEWENUM`) and `IEnumVARIANT`                         |
| `getProperty(obj, name, ...args)`          | `DISPATCH_PROPERTYGET` with arguments                                    |
| `putProperty(obj, name, ...args, value)`   | `DISPATCH_PROPERTYPUT` with arguments                                    |
| `putRefProperty(...)`, `invokeMethod(...)` | `DISPATCH_PROPERTYPUTREF`, `DISPATCH_METHOD` only                        |

Whether `obj.Name` reads a value or returns a callable comes from the type information (`IDispatch::GetTypeInfo`, cached per type):

- a property whose get has no required arguments is read: `dict.Count`, Excel `range.Value`;
- methods and properties with required arguments (`dict.Item`) return a function: methods resolve a `Promise` (executor), parameterized properties return their value (synchronous);
- without type information (e.g. WMI objects), the property is read, and a function is returned only when the object rejects the get (`DISP_E_MEMBERNOTFOUND`, `DISP_E_BADPARAMCOUNT`, ...), as comtypes does.

A property whose get has only **optional** arguments is read as a value, so `shell.Environment("PROCESS")` would call the default member of the result: use `getProperty(shell, "Environment", "PROCESS")`.

Results that are objects are dynamic objects again, also inside arrays. Each dynamic object owns its `IDispatch` reference: use `using` or `release(obj)`, otherwise the finalizer releases it later. `Dispatch(pointer)` wraps an existing pointer or `COMObject`, and `dispatchPointer(obj)` returns the underlying `IDispatch` pointer.

### Arguments

- `undefined` is an omitted optional argument (`missing`: `VT_ERROR` with `DISP_E_PARAMNOTFOUND`).
- Named arguments go in a final `named({ Name: value })`.
- `byRef(value)` passes a variable by reference (VB `ByRef`) as `VT_BYREF | VT_VARIANT`, `byRef(value, VT_BSTR)` as `VT_BYREF | VT_BSTR`; its `value` holds what the callee left after the call (`const n = byRef(1); await obj.Increment(n); n.value`).
- `typed(vt, value)` and `safeArray(values, { vt, dimensions })` choose the VARIANT type.

### VARIANT conversions

| JavaScript                                                                                                                      | VARIANT                                                          | Back to JavaScript                                                       |
| ------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------- | ------------------------------------------------------------------------ |
| `undefined` / `null`                                                                                                            | `VT_EMPTY` / `VT_NULL`                                           | `undefined` / `null`                                                     |
| `boolean`                                                                                                                       | `VT_BOOL`                                                        | `boolean`                                                                |
| `number`                                                                                                                        | `VT_I4` for 32-bit integers, else `VT_R8`                        | `number` for integer types up to 32 bits, reals, `VT_CY`, `VT_DECIMAL`   |
| `bigint`                                                                                                                        | `VT_I8` (`VT_UI8` above 2^63 - 1)                                | `bigint` for `VT_I8` / `VT_UI8`                                          |
| `string`                                                                                                                        | `VT_BSTR`                                                        | `string`                                                                 |
| `Date`                                                                                                                          | `VT_DATE`, local wall-clock time                                 | `Date`                                                                   |
| `COMPointer`, dynamic object, `COMObject`                                                                                       | `VT_DISPATCH` (`VT_UNKNOWN` if not an `IDispatch`)               | `COMPointer`; best interface from typed calls, dynamic object if dynamic |
| `Array`                                                                                                                         | `VT_ARRAY \| VT_VARIANT` (nested arrays: arrays of arrays)       | `Array`; N-D arrays as nested arrays                                     |
| `Buffer`, typed arrays                                                                                                          | `VT_ARRAY` of `VT_UI1`, `VT_I4`, `VT_R8`, ...                    | `Array`                                                                  |
| `ErrorValue`, `missing`                                                                                                         | `VT_ERROR`                                                       | `ErrorValue`                                                             |
| record value (`record(type, values)`, a value read from a `VT_RECORD`, an instance of a record class built from a type library) | `VT_RECORD` through `IRecordInfo` (the record type needs a GUID) | a plain object of fields, keeping its `IRecordInfo` under `recordInfo`   |
| `typed(vt, value)`                                                                                                              | the given type: `typed(VT_UI1, 7)`, `typed(VT_CY, "1.25")`       |                                                                          |
| `safeArray(values, { vt, dimensions })`                                                                                         | `VT_ARRAY` of `vt`; `dimensions: 2` for `values[i][j]`           |                                                                          |

`VT_BYREF` values are dereferenced when read. A `VARIANT` (`new VARIANT(value)`, `.value`, `.vt`, `.clear()`, `.copy()`, `.changeType(vt, flags, lcid)`, `.exact` for the exact decimal text of `VT_DECIMAL` / `VT_CY`) owns its BSTRs, interfaces and arrays: dispose it. SAFEARRAYs use the Windows column-major layout; `readSafeArray` / `createSafeArray` are exported for vtable interfaces.

## Defining interfaces

The `methods` array is the vtable order after the base interface, as `_methods_` in Python. Signatures are inferred from the specs, inherited methods included.

### COMMETHOD (IDL flags)

`COMMETHOD(idlflags, restype, name, ...params)` (Python: `comtypes.COMMETHOD`) converts arguments and results. Each parameter is `[flags, type, name?, default?]`:

```typescript
import { c_int32, c_uint32 } from "node-ctypes";
import { COMInterface, COMMETHOD, HRESULT, BSTR, VARIANT, SELF, IDispatch } from "@node-ctypes/comtypes";

export const IThing = COMInterface({
  name: "IThing",
  iid: "{...}",
  extends: IDispatch,
  methods: [
    COMMETHOD(["propget"], HRESULT, "Name", [["out", "retval"], BSTR, "pVal"]), // thing.Name
    COMMETHOD(["propput"], HRESULT, "Name", [["in"], BSTR, "value"]), // thing.Name = "x"
    COMMETHOD(["propget"], HRESULT, "Item", [["in"], VARIANT, "index"], [["out", "retval"], SELF, "pItem"]), // thing.Item(1): IThing pointer
    COMMETHOD([], HRESULT, "Find", [["in"], BSTR, "what"], [["in", "optional"], c_int32, "start", 0], [["out"], c_int32, "index"], [["out", "retval"], VARIANT, "match"]), // [index, match]
    COMMETHOD([], HRESULT, "Reset", [["lcid"], c_uint32, "lcid"]), // thing.Reset(): lcid filled
  ],
});
```

| Parameter flags         | Behaviour                                                                                                                                              |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `in` (or none)          | a JavaScript argument                                                                                                                                  |
| `out`, `out` + `retval` | allocated and returned: one output is the result, several are a tuple in declaration order; without outputs the HRESULT (`S_OK`/`S_FALSE`) is returned |
| `in` + `out`            | passed and returned                                                                                                                                    |
| `optional`, 4th element | may be omitted (pass `undefined`); the default value is used, `missing` for a `VARIANT`                                                                |
| `lcid`                  | filled with the default LCID, not an argument                                                                                                          |

A final `named({ name: value })` passes arguments by name at run time (not in the static signature). Missing required arguments, extra arguments and unknown names throw a `TypeError` before the call.

| Parameter type (the **value** type; `out` adds the pointer)        | `in`                                                                                                                                                              | `out`                                        |
| ------------------------------------------------------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------- |
| `BSTR`                                                             | `string \| null`, allocated and freed                                                                                                                             | `string`, freed                              |
| `LPOLESTR`                                                         | `string \| null`                                                                                                                                                  | `string \| null`, freed with `CoTaskMemFree` |
| `VARIANT`                                                          | any value (see the conversions), **by value**; a `VARIANT` instance is passed as is                                                                               | converted value, cleared                     |
| an interface (`IFoo` means `IFoo*`), `SELF`                        | `COMPointer`, dynamic object or `COMObject`; queried for the interface when needed                                                                                | owned `InterfacePointer` or `null`           |
| `SAFEARRAY(vt)`                                                    | arrays, typed arrays, `typed(...)`                                                                                                                                | `unknown[]`, destroyed                       |
| `REFGUID` / `REFIID` / `REFCLSID`                                  | `GUID` or GUID string                                                                                                                                             | `GUID`                                       |
| `GUID` and other `Structure`s                                      | instance or plain object, by value                                                                                                                                | the structure                                |
| simple types (`c_int32`, `c_double`, `wintypes.VARIANT_BOOL`, ...) | as node-ctypes                                                                                                                                                    | `.value`                                     |
| `POINTER(T)`                                                       | raw pointers (`bigint`, `Buffer`, node-ctypes arrays, structures) pass through; for `VARIANT`, simple and structure types, values go through a temporary `T` (\*) | as `T` (Python style `["out"], POINTER(T)`)  |
| `DATE`                                                             | `Date` or an OLE Automation date                                                                                                                                  | `Date`                                       |
| `CURRENCY`                                                         | number, exact decimal string or bigint (whole units)                                                                                                              | number                                       |
| `lazy(() => IFoo)`                                                 | as an interface defined later (mutual references)                                                                                                                 | owned `InterfacePointer<IFoo>` or `null`     |

(\*) For `POINTER(VARIANT)`, a COM pointer is a value (`VT_DISPATCH`/`VT_UNKNOWN`), not the address to pass, and `undefined` is a `VT_EMPTY` VARIANT.

Property accessors are named as in comtypes (`_get_Name`, `_set_Name`, `_setref_Name`) and also exposed as JavaScript properties:

- a `propget` without required arguments is a property, writable when a `propput`/`propputref` exists (`propputref` is chosen for COM objects);
- with arguments it is a method named after the property (`thing.Item(1)`), and the put stays available as `_set_Item(1, value)`.

`SELF` refers to the interface being defined (e.g. `IEnumVARIANT::Clone`, `IStream::Clone`) and is typed as such.

### Dispinterfaces (DISPMETHOD, DISPPROPERTY)

Members of a pure dispinterface have no vtable slot: they are called through `IDispatch::Invoke` with their DISPID. Declare them in `dispMethods` (Python: `_disp_methods_`):

```typescript
import { c_int32 } from "node-ctypes";
import { BSTR, COMInterface, CURRENCY, DISPMETHOD, DISPPROPERTY, IDispatch, VARIANT, dispid } from "@node-ctypes/comtypes";

const Font = COMInterface({
  name: "Font",
  iid: "{BEF6E003-A874-101A-8BBA-00AA00300CAB}",
  extends: IDispatch,
  dispMethods: [
    DISPPROPERTY([dispid(0)], BSTR, "Name"), // font.Name / font.Name = "Arial"
    DISPPROPERTY([dispid(2)], CURRENCY, "Size"),
    DISPPROPERTY([dispid(9), "readonly"], c_int32, "Charset"),
    DISPMETHOD([dispid(10)], VARIANT, "Describe", [["in"], VARIANT, "what"], [["in", "optional"], c_int32, "depth", 1]),
  ],
});
```

`DISPMETHOD(idlflags, valueType, name, ...params)` takes `in` parameters (optional, defaults, a final `named({...})`); arguments are converted to their declared type, the result is converted from the VARIANT, and interface results are queried for the declared interface. `propget`/`propput` DISPMETHODs and `DISPPROPERTY` become JavaScript properties as for COMMETHOD.

### STDMETHOD (raw)

`STDMETHOD(restype, name, argtypes)` passes node-ctypes arguments as they are, for methods whose buffers the caller manages (e.g. `IShellLinkW::GetPath` with `create_unicode_buffer`):

```typescript
import { c_int32, c_uint32, c_void_p } from "node-ctypes";
import { COMInterface, STDMETHOD, HRESULT, IUnknown } from "@node-ctypes/comtypes";

export const IShellLinkWPart = COMInterface({
  name: "IShellLinkW",
  iid: "{000214F9-0000-0000-C000-000000000046}",
  extends: IUnknown,
  methods: [STDMETHOD(HRESULT, "GetPath", [c_void_p, c_int32, c_void_p, c_uint32]) /* ... */],
  register: false, // do not replace the built-in definition for this IID
});
// link.GetPath(buffer, 260, null, 0): Promise<number>; sync(link).GetPath(...): number
```

Use `c_int32`/`c_uint32` for `LONG`/`DWORD`/`ULONG`: the `wintypes` aliases map the 32-bit `long` types to `bigint`. In-memory structures such as `GUID` can be passed to `POINTER(T)` parameters. `callMethod(name, ...args)` calls any slot by name. Definitions are registered by IID (the most recent wins) unless `register: false`.

### Built-in interfaces

`@node-ctypes/comtypes/interfaces` (also exported by `@node-ctypes/comtypes`) defines:

- COM core: `IUnknown`, `IClassFactory`, `IErrorInfo`, `ICreateErrorInfo`, `ISupportErrorInfo`, `IServiceProvider` (with `queryService(object, sid, IFoo)`);
- automation: `IDispatch`, `IEnumVARIANT`, `ITypeInfo`, `ITypeLib`, `ITypeComp`, `IProvideClassInfo`, `IProvideClassInfo2`, `ICreateTypeLib`, `ICreateTypeLib2`, `ICreateTypeInfo`;
- connection points: `IConnectionPoint`, `IConnectionPointContainer`, `IEnumConnections`, `IEnumConnectionPoints`;
- persistence and storage: `IPersist`, `IPersistFile`, `IPropertyBag`, `IPersistPropertyBag`, `IPropertyBag2`, `IPersistPropertyBag2`, `IErrorLog`, `ISequentialStream`, `IStream`;
- views: `IViewObject`, `IViewObject2` (`SIZEL`, `DVASPECT_*`);
- shell: `IShellLinkW`, `IShellLinkA`, the `ShellLink` coclass (`CLSID_ShellLink`).

Most use COMMETHOD. Raw methods (STDMETHOD) remain where the caller passes buffers, arrays or structures: `IUnknown`, `IDispatch`, `IClassFactory`, `ITypeInfo`, `ITypeLib`, `ITypeComp`, `IServiceProvider`, `IShellLinkW`/`IShellLinkA`, and some methods of `IEnumVARIANT` (`Next`), the connection point enumerators, `IPropertyBag2`, `IViewObject` and `ICreateTypeInfo`.

## COM objects in JavaScript

`COMObject` (Python: `comtypes.COMObject`) builds native vtables whose slots call JavaScript:

```typescript
import { COMInterface, COMMETHOD, CreateObject, Dispatch, HRESULT, BSTR, VARIANT, IDispatch, dispid } from "@node-ctypes/comtypes";
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
  Label = "counter"; // propget/propput fall back to the JavaScript property
  value = 0;
  Increment(step = 1) {
    return (this.value += step); // the [out, retval] value
  }
}

const counter = new Counter();
using pointer = counter.QueryInterface(ICounter); // owned ICounter pointer (AddRef)
await pointer.Increment(5); // called on the JavaScript thread: an apartment-threaded object
using dynamic = Dispatch(counter); // late binding through the generated IDispatch
using dict = await CreateObject("Scripting.Dictionary");
await dict.Add("counter", counter); // pass it to native code like any COM object
```

- **Implementations** are found by vtable name (`Increment`, `_get_Label`, `_set_Label`); accessors fall back to a property of the same name. Missing implementations return `E_NOTIMPL`.
- **Arguments**: COMMETHOD arguments are converted (`in` values as arguments, `out` values returned: one value, or an array; `[in, out]` values are received and returned). A `[in] VARIANT*` parameter arrives as its value; other `T*` parameters as addresses. STDMETHOD slots receive raw arguments (pointers as `bigint`) and may return an HRESULT.
- **Interface arguments** (pointers and dispatch objects) are lent for the duration of the call and released afterwards: `QueryInterface` or `AddRef` them to keep them.
- **Late binding** for interfaces deriving from `IDispatch`: the members of the declared interfaces (their `dispid`s, and their parameter names for named arguments) and the public JavaScript members (methods, accessors and fields, not `_`-prefixed); `_com_default_member_` answers `DISPID_VALUE`; iterable objects answer `_NewEnum` with a `VariantEnumerator`. Override `_invoke_` / `_ids_of_names_` to customize. Interfaces built from a type library (`GetModule`, generated modules) also provide their `ITypeInfo` through `GetTypeInfo`.
- **Lifetime**: native reference counting; while COM holds references the object is kept alive; `_final_release_()` runs when the count drops to zero. `getLiveObjectCount()` helps detect leaks.
- **Threading**: callbacks run on the JavaScript thread. With `_threading_ = "apartment"` (the default, for STA objects) calls on the object, and calls passing it as an argument, run on the JavaScript thread instead of the executor. `"free"` uses thread-safe callbacks: calls from the executor and other threads are queued to the JavaScript thread and wait for it, so they deadlock while that thread is blocked in a synchronous COM call that waits for them (`sync()`, a call at home).

### Errors in implementations

- An exception becomes a failure HRESULT (`COMError.hresult`, otherwise `E_FAIL`) with `IErrorInfo`: description, source (`_reg_progid_` or the class name) and interface GUID. `ISupportErrorInfo` is implemented automatically.
- Through `IDispatch`, exceptions fill `EXCEPINFO` (`DISP_E_EXCEPTION`); a `DispatchError` is returned as its HRESULT (`DISP_E_*`).
- `static onError(error, object)` observes the exceptions.
- `throw new ReturnHRESULT(hresult, text?)` (Python: `comtypes.ReturnHRESULT`) returns any HRESULT, also a success code such as `S_FALSE`, without calling `onError`.

## Events

```typescript
import { GetEvents, ShowEvents, startMessagePump } from "@node-ctypes/comtypes";

declare const source: object; // an event source: COM pointer, dispatch object or COMObject

using pump = startMessagePump(); // out-of-process and many in-process sources need window messages
using connection = GetEvents(source, {
  OnObjectReady(object: unknown) {
    console.log("ready", object);
  }, // event names from the type information
  OnCompleted(result: unknown) {
    console.log("completed", result);
  },
});
using log = ShowEvents(source); // prints every event
```

`GetEvents(source, handler, { interface })` (Python: `comtypes.client.GetEvents`) finds the event interface (the option, `IProvideClassInfo2`, the coclass default source, or the type library) and its names, creates a sink (a `COMObject`) and calls `Advise`; disposing the connection calls `Unadvise`.

- Events call `handler.Name(...)`, `handler.Interface_Name(...)` or `handler.onEvent(name, args)`; unhandled events are ignored.
- Dispatch arguments are dynamic objects lent for the call; by-reference parameters arrive as `ByRef` (`BeforeClose(cancel) { cancel.value = true; }`).
- Dispinterface and vtable event interfaces are supported.

A `COMObject` becomes an event source with `static _outgoing_interfaces_` (connection points, `IConnectionPointContainer`, `IProvideClassInfo2`) and raises events with `this.fireEvent("Name", ...args)`: through `IDispatch::Invoke` for dispinterfaces (the `dispid` of the spec), through the vtable otherwise.

### Message pump

| Function                                | Behaviour                                                                                                    |
| --------------------------------------- | ------------------------------------------------------------------------------------------------------------ |
| `startMessagePump({ interval })`        | `PeekMessage`/`DispatchMessage` every few ms from the event loop (unref'd); the JavaScript thread stays free |
| `await waitFor(condition, { timeout })` | pump until the condition holds                                                                               |
| `pumpMessages()`                        | dispatch what is queued, without blocking                                                                    |
| `PumpEvents(ms)`                        | block for `ms` in `CoWaitForMultipleHandles`, dispatching calls and messages (Python: `PumpEvents`)          |

## Servers

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

// In-process: creatable by CoCreateInstance(CLSCTX_INPROC_SERVER) in this process, no registry.
using registration = registerClassObject(Calculator, { clsctx: CLSCTX_INPROC_SERVER });

// LocalServer32 (another process): the script ends with
await UseCommandLine([Calculator]);
// node calculator.ts /regserver     registers it for the current user (HKCU\Software\Classes)
// node calculator.ts /unregserver   removes the registration
// node calculator.ts -Embedding     serves it: COM starts the script this way when a client creates the object
```

Clients then use `CreateObject("My.Calculator")`. See [`examples/localserver.ts`](./examples/localserver.ts), which PowerShell can call.

`UseCommandLine` (Python: `comtypes.server.register.UseCommandLine`) calls `registerServer`, `unregisterServer` and `runLocalServer` (STA, class factories, message pump until idle), which can also be used directly. Switches accept `/` or `-`, in any case (Git Bash rewrites `/regserver` as a path: use `-regserver` there).

### Registration

`registerServer` writes, per user:

- `LocalServer32`: the Node.js executable, its options (`execArgv`), the script and its arguments;
- `ProgID` and `VersionIndependentProgID` (`_reg_novers_progid_`, with `CurVer`), and the description (`_reg_desc_`);
- `TypeLib` and `Version`, when the class has a type library: `_reg_typelib_` (`[libid, major, minor]`), or the library its interfaces were built from (`GetModule`, generated modules). The library file (`_typelib_path_`, or the file `GetModule` loaded) is registered too, unless `typelib: false`.

Registering the type library matters across processes. `IDispatch` is marshaled by the system, but **custom vtable interfaces** (`[oleautomation]` or dual) get their proxy/stub from the registered type library (the type library marshaler); without it, `QueryInterface` for them fails with `E_NOINTERFACE` in the client. Classes with a `_reg_clsid_` and a type library also implement `IProvideClassInfo2`, so clients get their default interface typed.

## Type libraries

```typescript
import { CreateObject, GetModule } from "@node-ctypes/comtypes";

const Scripting = (await GetModule("scrrun.dll")) as any; // Python: comtypes.client.GetModule; untyped: generate a module for static types
using fso = await CreateObject(Scripting.FileSystemObject); // the coclass default interface (IFileSystem3), through the vtable
using temp = await fso.GetSpecialFolder(Scripting.SpecialFolderConst.TemporaryFolder); // IFolder
console.log(temp.Path);
```

`GetModule(source)` loads a type library and builds in memory every type: enums (frozen objects), records (`Structure` classes), interfaces (COMMETHOD, dual interfaces through the vtable), dispinterfaces (DISPMETHOD / DISPPROPERTY), coclasses (`CoClass`: `CreateObject(coclass)` returns the default interface), aliases and module constants. Interfaces defined by node-comtypes (`IUnknown`, `IDispatch`, `IEnumVARIANT`, ...) are reused; references between interfaces are resolved lazily. Results are cached per LIBID, version and LCID.

Reading and building a library is mostly JavaScript work on the calling thread (a large library such as mshtml takes a few seconds). `await GetModule()` does it in slices of a few milliseconds and yields to the event loop between them, so timers, I/O and an Electron UI keep running; concurrent calls for the same library share one build. `getModuleSync(source)` does the same work at once, for code that cannot wait (the cache is shared). `CreateObject`, `GetObject` and `GetActiveObject` load the type library of the object's best interface this way before returning it. The library of an object in another apartment or process is read from the same registered library loaded in-process, not call by call through its proxy.

Sources: a file (`"scrrun.dll"`, a path, `"file.dll\\2"`), a LIBID (`"{...}"` or `{ libid, major, minor, lcid }`, highest registered version by default), a ProgID (`"Scripting.Dictionary"`, through the registry, without creating the object) or an `ITypeLib` pointer. Lower level: `loadTypeLib`, `typeLibOf(object)`, `readTypeLib` (a JSON-serializable model) and `buildModule(model)`.

### Generated modules

```sh
npx node-comtypes generate scrrun.dll --out src/com        # src/com/Scripting.ts (with the package installed)
npx node-comtypes generate "Excel.Application" --out src/com --js   # Excel.js + Excel.d.ts (needs typescript)
npx node-comtypes info msxml6.dll                           # list the types
npx -p @node-ctypes/comtypes node-comtypes info scrrun.dll  # without installing it
```

```typescript
import { CreateObject } from "@node-ctypes/comtypes";
import * as Scripting from "./com/Scripting.ts";

using dict = await CreateObject(Scripting.Dictionary); // InterfacePointer<IDictionary>
dict.CompareMode = Scripting.CompareMethod.TextCompare;
const count: number = dict.Count; // read-only: assigning it is a type error
const exists: boolean = await dict.Exists("a"); // methods: Promise<...>; sync(dict).Exists("a"): boolean
```

A generated module embeds the library model and builds it with `buildModule` at import time (what `GetModule` does), and adds static types:

- enums as frozen objects with literal member types, and a union type of their values;
- record field interfaces;
- one `…Methods` interface per interface and dispinterface (methods returning promises, properties, read-only properties, `SELF` and cross references), and `…SyncMethods` for `sync()` views;
- typed coclasses and constants.

Python's `comtypes.gen` cache corresponds to `GetModule(source, { cacheDir })` (or `NODE_COMTYPES_CACHE`), which also writes `<Library>.ts`. The generator is `generateModule(model)` / `writeModule(model, directory)`.

### Writing and registering type libraries

`createTypeLib(file)` (Python: `comtypes.typeinfo.CreateTypeLib`) starts a type library without MIDL, through `ICreateTypeLib2` and `ICreateTypeInfo`; `SaveAllChanges()` writes it. `registerTypeLib(file, { scope })` registers it, per user by default (`"machine"` requires administrator rights), and returns its identity (`{ libid, major, minor, lcid }`), which `unregisterTypeLib(identity, { scope })` takes.

## Errors

`COMError` carries `hresult` (signed), `text` (`FormatMessageW`), `method` (`IShellLinkW::SetPath`, `IDispatch::Invoke(Add)`) and `details`:

- vtable calls: `IErrorInfo` (description, source, help file and context, GUID), read only when the object reports `ISupportErrorInfo` support for the failing interface, as comtypes does;
- `IDispatch::Invoke`: for `DISP_E_EXCEPTION`, the `EXCEPINFO` with its `scode` (some servers, such as the Scripting runtime, set only the scode, which is shown in the message); for `DISP_E_TYPEMISMATCH` / `DISP_E_PARAMNOTFOUND`, the `argumentIndex`.

Visual Basic run-time errors (`FACILITY_CONTROL`, `0x800Axxxx`, e.g. from the Scripting runtime) get their text from vbscript.dll when it is installed (a Windows Feature on Demand).

## Entry points

- **`@node-ctypes/comtypes`**: everything except `low-level`, plus the core:
  - `GUID`, HRESULT codes and helpers, `COMError`, `checkHRESULT`;
  - `COMInterface`, `COMMETHOD`, `STDMETHOD`, `DISPMETHOD`, `DISPPROPERTY` and the parameter types (`BSTR`, `LPOLESTR`, `REFGUID`, `SAFEARRAY`, `SELF`, `DATE`, `CURRENCY`, `lazy`, `dispid`, `helpstring`);
  - `COMPointer`, `CoInitializeEx`, `apartment`, `CoCreateInstance(Ex)`, `CoGetClassObject`, `CLSCTX_*` / `COINIT_*`;
  - message pump: `startMessagePump`, `waitFor`, `pumpMessages`, `PumpEvents`;
  - threads: `marshalInterface`, `unmarshalInterface`, `unmarshalDispatch`, `releaseMarshaled`, `GlobalInterfaceTable`;
  - execution: `sync`, `run`, `stats`, `finalize`.
- **`@node-ctypes/comtypes/client`**: `CreateObject`, `GetObject`, `sync`, `GetActiveObject`, `GetBestInterface` / `wrap`, `Constants`, `Dispatch` and the dynamic object helpers (`getProperty`, `putProperty`, `putRefProperty`, `invokeMethod`, `enumerate`, `dispatchPointer`, `isDispatch`, `release`, `memberInfo`, `typeName`), `GetEvents`, `ShowEvents`.
- **`@node-ctypes/comtypes/automation`**: `VARIANT`, `typed`, `safeArray`, `byRef` / `ByRef`, `named`, `missing`, `ErrorValue`, `record` / `RecordInfo`, SAFEARRAY functions, `DECIMAL`, dates, `DISPPARAMS`, `EXCEPINFO`, `invoke` / `getIDsOfNames`, `setDefaultLCID`, type information helpers, `VT_*` / `DISPID_*` and the other automation constants.
- **`@node-ctypes/comtypes/typelib`**: `GetModule`, `getModuleSync`, `loadTypeLib` and the registered library functions, `typeLibOf`, `readTypeLib` and the model types, `buildModule`, `generateModule`, `writeModule`, `createTypeLib`, `registerTypeLib`, `unregisterTypeLib`, `isCoClass` / `CoClass`.
- **`@node-ctypes/comtypes/server`**: `COMObject`, `VariantEnumerator`, `ConnectionPoint`, `ClassFactory`, `DispatchError`, `ReturnHRESULT`, `registerClassObject`, `registerServer`, `unregisterServer`, `runLocalServer`, `UseCommandLine`, `REGCLS_*`, `setErrorInfo`, `getLiveObjectCount`.
- **`@node-ctypes/comtypes/interfaces`**: the [built-in interfaces](#built-in-interfaces) with their structures and constants.
- **`@node-ctypes/comtypes/low-level`**: lazy `ole32`/`oleaut32` bindings, `executorCall` (a raw call on the COM executor, with `IErrorInfo`), raw vtable access (`rawQueryInterface`, `rawAddRef`, `rawRelease`), BSTR and task-memory helpers, registry helpers.

## Python comtypes correspondence

| comtypes                                                     | node-comtypes                                                     |
| ------------------------------------------------------------ | ----------------------------------------------------------------- |
| **Interfaces and pointers**                                  |                                                                   |
| `GUID("{...}")`, `GUID.create_new()`                         | `new GUID("{...}")`, `GUID.createNew()`                           |
| `GUID.from_progid()`, `.as_progid()`                         | `GUID.fromProgID()`, `.asProgID()`                                |
| `class IFoo(IUnknown): _iid_, _methods_`                     | `COMInterface({ name, iid, extends, methods })`                   |
| `STDMETHOD(HRESULT, "Name", [args])`                         | `STDMETHOD(HRESULT, "Name", [args])`                              |
| `COMMETHOD([flags], HRESULT, "Name", (["in"], T, "x"), ...)` | `COMMETHOD([flags], HRESULT, "Name", [["in"], T, "x"], ...)`      |
| `(["out"], POINTER(BSTR), "p")`                              | `[["out"], BSTR, "p"]` (`POINTER(T)` also accepted)               |
| `POINTER(IFoo)` parameter / `POINTER(POINTER(IFoo))` out     | `IFoo` (`SELF` for the interface itself)                          |
| `_midlSAFEARRAY(BSTR)`                                       | `SAFEARRAY(VT_BSTR)`                                              |
| `dispid(n)`, `helpstring(text)`                              | `dispid(n)`, `helpstring(text)`                                   |
| `DISPMETHOD`, `DISPPROPERTY`, `_disp_methods_`               | `DISPMETHOD`, `DISPPROPERTY`, `dispMethods`                       |
| `p.Name`, `p._get_Name()`, `p.Name = v`                      | same                                                              |
| out parameters as a tuple                                    | out parameters as an array                                        |
| `CoInitializeEx()`, `CoCreateInstance(clsid, interface)`     | `CoInitializeEx()` / `apartment()`, `CoCreateInstance(...)`       |
| `p.QueryInterface(IFoo)`                                     | `p.QueryInterface(IFoo)`                                          |
| `p == q`                                                     | `p.isSameObject(q)`                                               |
| `del p` / refcount                                           | `p.Release()` / `using`                                           |
| `COMError(hresult, text, details)`                           | `COMError` with `hresult`, `text`, `details`                      |
| **Client**                                                   |                                                                   |
| `client.CreateObject(progid)`                                | `await CreateObject(progid)` (best interface)                     |
| `obj.Method(...)` (blocking)                                 | `await obj.Method(...)`; blocking: `sync(obj).Method(...)`        |
| `client.CreateObject(progid, dynamic=True)`                  | `CreateObject(progid, { dynamic: true })`                         |
| `client.GetActiveObject`, `client.GetObject`                 | `GetActiveObject`, `GetObject`                                    |
| `client.GetBestInterface`, `client.wrap`                     | `GetBestInterface`, `wrap`                                        |
| `client.dynamic.Dispatch(p)`                                 | `Dispatch(p)`                                                     |
| `client.Constants(obj)`                                      | `Constants(obj)`                                                  |
| `obj.Item[key] = value`                                      | `putProperty(obj, "Item", key, value)`                            |
| keyword arguments                                            | `named({ Name: value })`                                          |
| byref arguments (win32com: returned values)                  | `byRef(value)`, read `.value` after the call                      |
| `automation.VARIANT`, `.value`                               | `VARIANT`, `.value`                                               |
| `comtypes.automation.Missing`                                | `missing`                                                         |
| `client.GetEvents`, `client.ShowEvents`                      | `GetEvents`, `ShowEvents`                                         |
| `client.PumpEvents`                                          | `PumpEvents` (+ `startMessagePump`)                               |
| `CoMarshalInterThreadInterfaceInStream`, GIT                 | `marshalInterface` / `unmarshalInterface`, `GlobalInterfaceTable` |
| **Type libraries**                                           |                                                                   |
| `client.GetModule(path)`, `comtypes.gen`                     | `await GetModule(path, { cacheDir })`, `node-comtypes generate`   |
| generated `CoClass` subclasses                               | `CoClass` objects; `CreateObject(coclass)`                        |
| `typeinfo.CreateTypeLib`, `RegisterTypeLib`                  | `createTypeLib`, `registerTypeLib`, `unregisterTypeLib`           |
| **Servers**                                                  |                                                                   |
| `class X(COMObject): _com_interfaces_ = [IFoo]`              | `class X extends COMObject { static _com_interfaces_ = [IFoo] }`  |
| `_outgoing_interfaces_`, `Fire_Event`                        | `_outgoing_interfaces_`, `fireEvent("Name", ...)`                 |
| `_final_release_`, `_reg_clsid_`, `_reg_progid_`             | same                                                              |
| `_reg_desc_`, `_reg_novers_progid_`                          | same                                                              |
| `_reg_typelib_`, `_typelib_path_`                            | same                                                              |
| `ReturnHRESULT(hresult, text)`                               | `throw new ReturnHRESULT(hresult, text)`                          |
| `server.register.UseCommandLine`, `register`                 | `UseCommandLine`, `registerServer`                                |
| `server.localserver.run`                                     | `runLocalServer`                                                  |
| **Other modules**                                            |                                                                   |
| `comtypes.persist`, `comtypes.viewobject`                    | `IPropertyBag`, `IPersistPropertyBag`(2), `IViewObject`(2), ...   |
| `IServiceProvider.QueryService(sid, interface)`              | `queryService(object, sid, interface)`                            |
| `comtypes.shelllink.ShellLink`                               | `ShellLink` (`CreateObject(ShellLink)`), `IShellLinkA`            |

## Compatibility and limits

- Methods run on one executor thread, in order: a slow call delays the next ones (not the event loop). Each costs a round trip to that thread; see [Synchronous calls](#synchronous-calls) for chatty loops. Win32 messages are pumped only on request ([Message pump](#message-pump)).
- Objects inside `ByRef` arguments and `VARIANT` structures passed by address are not handed over between apartments: in a single-threaded home, pass such arguments to objects created with `{ home: true }` or use `sync()`.
- Hand-rolled vtables (`COMObject`s aside) that are not thread-safe must not be passed to methods of objects shared with the executor in a multithreaded home.
- Objects returned without `{ dynamic: true }` expose the members of their type library only: properties added at run time (WMI class properties, `IDispatchEx`) need a dynamic object, as in Python. An unknown member of a typed pointer is `undefined` (Python raises `AttributeError`), and typed pointers cannot be called for their default member.
- A dynamic property whose get has only optional arguments is read as a value ([Dynamic objects](#dynamic-objects)).
- `named()` works at run time only: static signatures are positional.
- C arrays (`size_is`) have no codec: use STDMETHOD or `POINTER(T)` with your own buffers.
- `VT_RECORD` needs a record type with a GUID in its type library (an `IRecordInfo` requirement).
- Custom vtable interfaces cross apartments and processes only with a registered type library or proxy/stub; `IDispatch` always does.
- In-process servers (InprocServer32) cannot host Node.js, so `_reg_threading_` and `_reg_clsctx_` have no equivalent; a native shim forwarding to a LocalServer is a possible future addition.
- `_threading_ = "free"` objects deadlock when called from another thread while the JavaScript thread is blocked in a synchronous COM call (`sync()` or a call at home) waiting for that call.
- The Scripting runtime and the Shell do not implement `ISupportErrorInfo`: their vtable failures have no `details`.

Implementation flows, in Italian: [FLOWS.md](./FLOWS.md).

## Development

```sh
npm run build        # dist/
npm run typecheck
npm run test:types   # declaration and API type tests
npm run test:unit    # GUID, HRESULT, layouts, DECIMAL, DATE, VARIANT scalars, type library build and generation: any OS
npm test             # everything; COM tests are skipped outside Windows
npm pack --dry-run   # package contents
```

The [CI workflow](./.github/workflows/ci.yml) runs `npm test` with Node.js 24 and 26 on Windows x64 and ARM64. It is discovered only when this package is a repository root, not while nested in another repository. Pushing a `v<version>` tag that matches `package.json` publishes to npm once the whole matrix passes (npm Trusted Publishing; versions with a `-` go to `next`).

Examples, after a build (`node examples/<name>.ts`):

- [shortcut.ts](./examples/shortcut.ts): a shell link through `IShellLinkW` and `IPersistFile`;
- [automation.ts](./examples/automation.ts): Scripting.Dictionary, FileSystemObject and WMI;
- [stream.ts](./examples/stream.ts): `IStream` and COMMETHOD outputs;
- [server.ts](./examples/server.ts): a `COMObject` with events, used through the vtable, `IDispatch` and native code;
- [localserver.ts](./examples/localserver.ts): a LocalServer32 registered with `UseCommandLine`;
- [active-directory.ts](./examples/active-directory.ts): Active Directory search through ADSI `IDirectorySearch`, with the structures and enums of the Active DS type library;
- [active-directory-sync.ts](./examples/active-directory-sync.ts): the same search with every call synchronous on the JavaScript thread (`sync()`, `getModuleSync`, `CoGetObject`): faster, but it blocks the event loop for the whole search;
- [typelib.ts](./examples/typelib.ts): `GetModule` and coclasses.

## License

[MIT](./LICENSE). See [NOTICE](./NOTICE) for the Python comtypes attribution.
