import * as ctypes from "node-ctypes";
import { CLSID_ShellLink, COMInterface, COMPointer, CoCreateInstance, GUID, HRESULT, IPersistFile, IShellLinkW, IStream, IUnknown, STDMETHOD, apartment, sync } from "@node-ctypes/comtypes";
import type { COMError, InterfaceDef, InterfacePointer, SyncPointer, Synchronous } from "@node-ctypes/comtypes";
import { ole32, rawRelease } from "@node-ctypes/comtypes/low-level";
import { IClassFactory } from "@node-ctypes/comtypes/interfaces";

using scope = apartment();
const hresult: number = scope.hresult;

using link = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
const set: number = await link.SetPath("C:\\Windows");
await link.GetPath(ctypes.create_unicode_buffer(260), 260, null, 0);
const pointer: COMPointer<typeof IShellLinkW> = link;
// @ts-expect-error methods return promises
const notAwaited: number = link.SetPath("C:\Windows");
// The sync() view: the same methods, synchronous; properties and COMPointer members unchanged.
const view: SyncPointer<typeof IShellLinkW> = sync(link);
const sameView: Synchronous<typeof link> = view;
const setNow: number = view.SetPath("C:\Windows");
const viewAddress: bigint = view.address;
void [notAwaited, sameView, setNow, viewAddress];
const address: bigint = link.address;

using file = link.QueryInterface(IPersistFile);
await file.Save("x.lnk", 1);
const classId: GUID = await file.GetClassID();
const dirty: number = await file.IsDirty();
const maybe: InterfacePointer<typeof IStream> | null = link.tryQueryInterface(IStream);
const same: boolean = file.isSameObject(link);
using unknown = CoCreateInstance(CLSID_ShellLink);
const isUnknown: InterfacePointer<typeof IUnknown> = unknown;
const count: number = unknown.AddRef();
const raw: unknown = unknown.callMethod("Release");

// User-defined interfaces infer their method signatures, inherited ones included.
const IFoo = COMInterface({
  name: "IFoo",
  iid: "{11111111-2222-3333-4444-555555555555}",
  extends: IPersistFile,
  methods: [STDMETHOD(HRESULT, "Foo", [ctypes.c_int32, ctypes.c_wchar_p]), STDMETHOD(ctypes.c_uint32, "Count"), STDMETHOD(ctypes.c_double, "Ratio")],
});
const def: InterfaceDef = IFoo;
using foo = COMPointer.fromAddress(1n, IFoo);
const fooResult: number = await foo.Foo(1, "text");
const fooCount: number = await foo.Count();
const ratio: number = await foo.Ratio();
await foo.Save("y", 0);
const qi: InterfacePointer<typeof IUnknown> = foo.QueryInterface(IUnknown);

const factory: InterfacePointer<typeof IClassFactory> = COMPointer.fromAddress(1n, IClassFactory);
await factory.LockServer(1);
const freed: number = rawRelease(1n);
const native = ole32();
const error: COMError | null = null;

void [classId, hresult, set, pointer, address, dirty, maybe, same, isUnknown, count, raw, def, fooResult, fooCount, ratio, qi, factory, freed, native, error];

// @ts-expect-error SetPath takes a string.
link.SetPath(42);
// @ts-expect-error IShellLinkW has no Save; QueryInterface(IPersistFile) first.
link.Save("x.lnk", 1);
// @ts-expect-error Unknown method.
foo.Bar();
// @ts-expect-error Foo requires two arguments.
foo.Foo(1);
// @ts-expect-error QueryInterface takes an interface definition, not raw arguments.
link.QueryInterface(IStream.iid, null);

// ── Phase 2: automation ──────────────────────────────────────────────
import { CreateObject, GetObject, VARIANT, VT_I4, VT_ARRAY, typed, safeArray, named, putProperty, getProperty, missing, isDispatch, dispatchPointer, IDispatch as IDispatchDef, Dispatch, IEnumVARIANT } from "@node-ctypes/comtypes";
import type { DispatchObject, DispatchPointer, MemberInfo } from "@node-ctypes/comtypes";
import { invoke, getIDsOfNames } from "@node-ctypes/comtypes/automation";
import { CreateObject as CreateObjectClient } from "@node-ctypes/comtypes/client";

using dict = await CreateObject("Scripting.Dictionary", { dynamic: true });
const dynamic: DispatchObject = dict;
dict.Add("k", 1);
const anyValue: number = dict.Count;
const call: unknown = dict("k");
for (const key of dict) void key;
putProperty(dict, "Item", "k", 2);
const got: unknown = getProperty(dict, "Item", "k");
dict.Method(1, named({ Option: true }), missing);
const typedPointer: InterfacePointer<typeof IDispatchDef> = await CreateObject("Scripting.Dictionary", { interface: IDispatchDef });
const rawDispatch: DispatchPointer = dispatchPointer(dict);
const ids: number[] = getIDsOfNames(rawDispatch, ["Add"]);
const result: unknown = invoke(rawDispatch, ids[0], 1, ["a", 1]);
const wmi: DispatchObject = await GetObject("winmgmts:", { dynamic: true });
const sameCreate: typeof CreateObject = CreateObjectClient;
using variant = new VARIANT(typed(VT_ARRAY | VT_I4, [[1]], { dimensions: 2 }));
const value: unknown = variant.value;
const exact: string = variant.exact;
const array = new VARIANT(safeArray([1, 2]));
const isDyn: boolean = isDispatch(dict);
const info: MemberInfo | null = null;
void [dynamic, anyValue, call, got, typedPointer, ids, result, wmi, sameCreate, value, exact, array, isDyn, info];

// @ts-expect-error typed() takes a VARTYPE number first.
typed("VT_I4", 1);
// @ts-expect-error A typed interface pointer has no dynamic members.
typedPointer.Add("k", 1);

// ── Phase 3: COMMETHOD ───────────────────────────────────────────────
import { COMMETHOD, BSTR, SELF, SAFEARRAY, REFIID, VT_BSTR, IErrorInfo, IStream as IStreamDef, STATSTG } from "@node-ctypes/comtypes";

const IThing = COMInterface({
  name: "IThing",
  iid: "{22222222-3333-4444-5555-666666666666}",
  extends: IDispatchDef,
  methods: [
    COMMETHOD(["propget"], HRESULT, "Name", [["out", "retval"], BSTR, "pVal"]),
    COMMETHOD(["propput"], HRESULT, "Name", [["in"], BSTR, "value"]),
    COMMETHOD(["propget"], HRESULT, "Count", [["out", "retval"], ctypes.c_int32, "pVal"]),
    COMMETHOD(["propget"], HRESULT, "Item", [["in"], VARIANT, "index"], [["out", "retval"], SELF, "pItem"]),
    COMMETHOD(["propput"], HRESULT, "Secret", [["in"], BSTR, "value"]),
    COMMETHOD([], HRESULT, "Find", [["in"], BSTR, "what"], [["in", "optional"], ctypes.c_int32, "start", 0], [["out"], ctypes.c_int32, "index"], [["out", "retval"], SAFEARRAY(VT_BSTR), "matches"]),
    COMMETHOD([], HRESULT, "Reset", [["lcid"], ctypes.c_uint32, "lcid"]),
    COMMETHOD([], HRESULT, "Supports", [["in"], REFIID, "riid"]),
  ],
});
using thing = COMPointer.fromAddress(1n, IThing);
const thingName: string = thing.Name;
thing.Name = "renamed";
const viaAccessor: string = thing._get_Name();
const thingCount: number = thing.Count;
const item: InterfacePointer<typeof IThing> | null = thing.Item(1);
const itemName: string | undefined = item?.Name;
thing.Secret = "write only";
const found: [number, unknown[] | null] = await thing.Find("x");
const foundFrom: [number, unknown[] | null] = await thing.Find("x", 3);
const reset: number = await thing.Reset();
await thing.Supports("{00000000-0000-0000-C000-000000000046}");
await thing.Supports(IStreamDef.iid);
using errorInfo = COMPointer.fromAddress(1n, IErrorInfo);
const description: string = await errorInfo.GetDescription();
const context: number = await errorInfo.GetHelpContext();
using stream = COMPointer.fromAddress(1n, IStreamDef);
const clone: InterfacePointer<typeof IStreamDef> | null = await stream.Clone();
const position: bigint = await stream.Seek(0n);
const stat: STATSTG = await stream.Stat();
const copied: [bigint, bigint] = await stream.CopyTo(stream, 10n);
void [thingName, viaAccessor, thingCount, item, itemName, found, foundFrom, reset, description, context, clone, position, stat, copied];

// @ts-expect-error Count has no propput: read-only.
thing.Count = 3;
// @ts-expect-error Find requires `what`.
thing.Find();
// @ts-expect-error The lcid parameter is filled automatically.
thing.Reset(0);
// @ts-expect-error Name is a string.
thing.Name = 42;
// @ts-expect-error Item takes an index (parameterized property: a method).
const notAProperty: InterfacePointer<typeof IThing> = thing.Item;

// ── Phase 4: COM objects in JavaScript, events, pump ─────────────────
import { GetEvents, ShowEvents, startMessagePump, PumpEvents, waitFor, pumpMessages } from "@node-ctypes/comtypes";
import type { EventConnection, MessagePump } from "@node-ctypes/comtypes";
import { COMObject, registerClassObject, registerServer, unregisterServer, runLocalServer, getLiveObjectCount, VariantEnumerator } from "@node-ctypes/comtypes/server";
import type { ClassRegistration, InvokeArguments } from "@node-ctypes/comtypes/server";

class JsThing extends COMObject {
  static _com_interfaces_ = [IThing];
  static _outgoing_interfaces_ = [IThing];
  static _reg_clsid_ = "{55555555-5555-5555-5555-555555555555}";
  Name = "js";
  _invoke_(dispid: number, flags: number, args: InvokeArguments): unknown {
    return super._invoke_(dispid, flags, args);
  }
}
const jsThing = new JsThing();
using jsPointer = jsThing.QueryInterface(IThing);
const jsName: string = jsPointer.Name;
const refs: number = jsThing.refCount;
const fired: unknown[] = jsThing.fireEvent("Name", 1);
const connection: EventConnection = GetEvents(jsThing, { Name() {} });
const shown: EventConnection = ShowEvents(jsPointer, { log: (line: string) => void line });
const pump: MessagePump = startMessagePump({ interval: 5 });
PumpEvents(0);
const pumped: number = pumpMessages();
const waited: Promise<boolean> = waitFor(() => true, { timeout: 10 });
const registration: ClassRegistration = registerClassObject(JsThing);
registerServer(JsThing, { script: "server.ts" });
unregisterServer(JsThing);
const serving: Promise<void> = runLocalServer([JsThing]);
const liveObjects: number = getLiveObjectCount();
using enumerated = new VariantEnumerator([1, 2]).QueryInterface(IEnumVARIANT);
const dynamicFromObject: DispatchObject = Dispatch(jsThing);
void [jsName, refs, fired, connection, shown, pump, pumped, waited, registration, serving, liveObjects, dynamicFromObject];

// ── Phase 5: type libraries ──────────────────────────────────────────
import { DISPMETHOD, DISPPROPERTY, dispid, lazy, DATE, CURRENCY, GetModule, getModuleSync, loadTypeLib, readTypeLib, buildModule, generateModule, isCoClass } from "@node-ctypes/comtypes";
import type { CoClass, TypeLibModel, TypeLibModule, TypeLibPointer } from "@node-ctypes/comtypes";
import { GetModule as GetModuleTypelib, writeModule } from "@node-ctypes/comtypes/typelib";

const IFolderLike = COMInterface({
  name: "IFolderLike",
  iid: "{66666666-6666-6666-6666-666666666666}",
  extends: IDispatchDef,
  methods: [
    COMMETHOD(["propget"], HRESULT, "Parent", [["out", "retval"], lazy(() => IFolderLike2), "parent"]),
    COMMETHOD(["propget"], HRESULT, "Created", [["out", "retval"], DATE, "when"]),
    COMMETHOD(["propget"], HRESULT, "Price", [["out", "retval"], CURRENCY, "price"]),
  ],
});
const IFolderLike2 = COMInterface({ name: "IFolderLike2", iid: "{77777777-7777-7777-7777-777777777777}", extends: IDispatchDef, methods: [COMMETHOD(["propget"], HRESULT, "Name", [["out", "retval"], BSTR, "name"])] });
const IEventsLike = COMInterface({
  name: "IEventsLike",
  iid: "{88888888-8888-8888-8888-888888888888}",
  extends: IDispatchDef,
  dispMethods: [DISPMETHOD([dispid(1)], BSTR, "Describe", [["in"], VARIANT, "value"]), DISPPROPERTY([dispid(2)], ctypes.c_int32, "Level"), DISPPROPERTY([dispid(3), "readonly"], DATE, "Since")],
});
using folderLike = COMPointer.fromAddress(1n, IFolderLike);
const parentName: string | undefined = folderLike.Parent?.Name;
const created: Date = folderLike.Created;
const price: number = folderLike.Price;
using eventsLike = COMPointer.fromAddress(1n, IEventsLike);
const described: string = await eventsLike.Describe(1);
eventsLike.Level = 3;
const level: number = eventsLike.Level;
const since: Date = eventsLike.Since;
const module5: Promise<TypeLibModule> = GetModule("scrrun.dll");
const module5Sync: TypeLibModule = getModuleSync("scrrun.dll");
const sameGetModule: typeof GetModule = GetModuleTypelib;
const library5: TypeLibPointer = loadTypeLib("scrrun.dll");
const model5: TypeLibModel = readTypeLib(library5);
const built: TypeLibModule = buildModule(model5);
const code: string = generateModule(model5, { importFrom: "@node-ctypes/comtypes" });
const path5: string = writeModule(model5, "out");
declare const coclass5: CoClass<typeof IFolderLike>;
using created5 = await CreateObject(coclass5);
const typedFromCoClass: Date = created5.Created;
const isClass: boolean = isCoClass(coclass5);
void [parentName, created, price, described, level, since, module5, module5Sync, sameGetModule, built, code, path5, typedFromCoClass, isClass];

// @ts-expect-error Since is read-only.
eventsLike.Since = new Date();
// @ts-expect-error Describe takes a value.
eventsLike.Describe();

// ── Improvements: ByRef, records, threads ────────────────────────────
import { byRef, ByRef, record, recordInfo, RecordInfo, marshalInterface, unmarshalInterface, unmarshalDispatch, releaseMarshaled, GlobalInterfaceTable, VT_BOOL } from "@node-ctypes/comtypes";
import type { MarshaledInterface, RecordValue } from "@node-ctypes/comtypes";

const cancel: ByRef<boolean> = byRef(false);
const typedRef: ByRef<string> = byRef("x", VT_BSTR);
dict.Method(cancel, typedRef, byRef(1, VT_BOOL));
const cancelled: boolean = cancel.value;
declare const pointRecord: RecordValue;
const recordType: RecordInfo = pointRecord[recordInfo];
const fields: readonly string[] = recordType.fields;
const made: RecordValue = record(recordType, { x: 1 });
const token: MarshaledInterface = marshalInterface(dict, IDispatchDef);
const back: InterfacePointer<typeof IDispatchDef> = unmarshalInterface(token, IDispatchDef);
const backDynamic: DispatchObject = unmarshalDispatch(token);
releaseMarshaled(token);
using git = new GlobalInterfaceTable();
const cookie: number = git.register(dict);
const fromGit: DispatchObject = git.getDispatch(cookie);
void [cancelled, fields, made, back, backDynamic, fromGit];

// @ts-expect-error A token is consumed by unmarshalInterface with an interface definition.
unmarshalInterface(token);

// ── Parity: best interface, registration, type library creation ──────
import { GetBestInterface, wrap, createTypeLib, registerTypeLib, unregisterTypeLib, ICreateTypeLib2, TKIND_INTERFACE } from "@node-ctypes/comtypes";
import type { AutomationObject, DispatchTarget, TypeLibIdentity } from "@node-ctypes/comtypes";
import { UseCommandLine } from "@node-ctypes/comtypes/server";

using best = await CreateObject("Scripting.Dictionary");
const automation: AutomationObject = best;
best.Add("k", 1);
for (const key of best) void key;
const typedBest: COMPointer = GetBestInterface(link);
const sameWrap: typeof GetBestInterface = wrap;
const target: DispatchTarget = typedPointer;
const viaHelper: unknown = getProperty(typedPointer, "Item", "k");
using fromProgid = await CreateObject("Scripting.Dictionary", { interface: IDispatchDef });
const active: AutomationObject = await GetObject("winmgmts:");
using builder = createTypeLib("x.tlb");
const creator: InterfacePointer<typeof ICreateTypeLib2> = builder;
using createdInfo = await builder.CreateTypeInfo("IX", TKIND_INTERFACE);
const identity: TypeLibIdentity = registerTypeLib("x.tlb");
unregisterTypeLib(identity, { scope: "user" });
const handled: Promise<boolean> = UseCommandLine([JsThing]);
class Registered extends COMObject {
  static _com_interfaces_ = [IThing];
  static _reg_clsid_ = "{55555555-5555-5555-5555-555555555556}";
  static _reg_novers_progid_ = "Test.Registered";
  static _reg_typelib_ = ["{55555555-5555-5555-5555-555555555557}", 1, 0] as const;
  static _typelib_path_ = "x.tlb";
}
void [automation, typedBest, sameWrap, target, viaHelper, fromProgid, active, creator, createdInfo, handled, Registered];

// @ts-expect-error Only dynamic dispatch objects are callable.
best("k");

// ── Parity: Constants, minor interfaces, ReturnHRESULT ───────────────
import { Constants, queryService, ShellLink, IShellLinkA, IPropertyBag, IPersistPropertyBag, IViewObject2, SIZEL, DVASPECT_CONTENT } from "@node-ctypes/comtypes";
import type { ConstantsObject } from "@node-ctypes/comtypes";
import { ReturnHRESULT } from "@node-ctypes/comtypes/server";

const constants: ConstantsObject = Constants(best);
const temporary: number = constants.TemporaryFolder;
using shellLink = await CreateObject(ShellLink);
const wideLink: InterfacePointer<typeof IShellLinkW> = shellLink;
using ansiLink = shellLink.QueryInterface(IShellLinkA);
await ansiLink.SetPath("C:\Windows");
using service = queryService(shellLink, "{00000000-0000-0000-0000-000000000000}", IDispatchDef);
using bagPointer = COMPointer.fromAddress(1n, IPropertyBag);
const read: unknown = await bagPointer.Read("Name");
await bagPointer.Write("Name", 1);
using persistBag = COMPointer.fromAddress(1n, IPersistPropertyBag);
await persistBag.Load(bagPointer, null);
using viewObject = COMPointer.fromAddress(1n, IViewObject2);
const extent: SIZEL = await viewObject.GetExtent(DVASPECT_CONTENT, -1, null);
const returned: COMError = new ReturnHRESULT(1, "text");
void [temporary, wideLink, service, read, extent, returned];

// @ts-expect-error IShellLinkW has no Save.
shellLink.Save("x.lnk", 1);
