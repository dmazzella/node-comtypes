import * as ctypes from "node-ctypes";
import { inspect } from "node:util";
import { IMPLTYPEFLAG_FDEFAULT, IMPLTYPEFLAG_FSOURCE } from "../automation/constants.ts";
import { dispatchTypeInfo, documentation, typeMembers, withTypeAttr } from "../automation/members.ts";
import type { TypeInfoPointer } from "../automation/members.ts";
import { borrowReference } from "../automation/variant.ts";
import { COMError } from "../errors.ts";
import { GUID, toGUID } from "../guid.ts";
import type { GUIDLike } from "../guid.ts";
import { E_NOINTERFACE } from "../hresult.ts";
import { COMInterface, derivesFrom, findInterface } from "../interface.ts";
import type { InterfaceDef } from "../interface.ts";
import { GUIDKIND_DEFAULT_SOURCE_DISP_IID, IConnectionPointContainer, IProvideClassInfo2 } from "../interfaces/connectionpoint.ts";
import { IDispatch } from "../interfaces/dispatch.ts";
import { IProvideClassInfo, ITypeInfo, ITypeLib } from "../interfaces/typeinfo.ts";
import { IUnknown } from "../interfaces/unknown.ts";
import { COMPointer, sync } from "../pointer.ts";
import type { InvokeArguments } from "../server/automation.ts";
import { COMObject } from "../server/comobject.ts";

/** An advise connection; dispose (or `disconnect()`) to `Unadvise`. */
export interface EventConnection extends Disposable {
  readonly cookie: number;
  /** The event interface (given, or synthesized from its IID and type information). */
  readonly interface: InterfaceDef<any>;
  /** Event names by DISPID, when known. */
  readonly names: ReadonlyMap<number, string>;
  readonly connected: boolean;
  disconnect(): void;
}

export interface GetEventsOptions {
  /** Event interface: a definition (DISPIDs from its specs) or an IID; default: the source's default source interface. */
  readonly interface?: InterfaceDef<any> | GUIDLike;
}

type Handler = Record<string, unknown> & { onEvent?: (name: string, args: unknown[]) => unknown };

function sourcePointer(source: unknown): COMPointer {
  const borrowed = borrowReference(source);
  if (!borrowed) throw new TypeError("GetEvents expects a COM object");
  return COMPointer.fromAddress(borrowed.address, borrowed.interface, { addRef: true });
}

function typeInfoFromReference(typeInfo: TypeInfoPointer, reference: number): TypeInfoPointer {
  const out = new ctypes.c_void_p();
  sync(typeInfo).GetRefTypeInfo(reference, ctypes.byref(out));
  return COMPointer.fromAddress(out.value!, ITypeInfo);
}

/** Default source interface of a coclass type info. */
function defaultSourceOf(coclass: TypeInfoPointer): TypeInfoPointer | null {
  const count = withTypeAttr(coclass, (attr) => attr.cImplTypes);
  for (let index = 0; index < count; index++) {
    const flags = new ctypes.c_int32();
    sync(coclass).GetImplTypeFlags(index, ctypes.byref(flags));
    if ((flags.value & (IMPLTYPEFLAG_FDEFAULT | IMPLTYPEFLAG_FSOURCE)) !== (IMPLTYPEFLAG_FDEFAULT | IMPLTYPEFLAG_FSOURCE)) continue;
    const reference = new ctypes.c_uint32();
    sync(coclass).GetRefTypeOfImplType(index, ctypes.byref(reference));
    return typeInfoFromReference(coclass, reference.value);
  }
  return null;
}

function attempt<T>(work: () => T): T | null {
  try {
    return work();
  } catch {
    return null;
  }
}

/** Type information for event interface `iid`, from the coclass or the source's type library. */
function eventTypeInfo(source: COMPointer, iid: GUID | null): TypeInfoPointer | null {
  const provider = source.tryQueryInterface(IProvideClassInfo);
  if (provider) {
    using owned = provider;
    using coclass = attempt(() => sync(owned).GetClassInfo());
    if (coclass) {
      const fromCoclass = attempt(() => defaultSourceOf(coclass));
      if (fromCoclass && (!iid || withTypeAttr(fromCoclass, (attr) => attr.guid.equals(iid)))) return fromCoclass;
      fromCoclass?.Release();
      if (iid) {
        const fromLibrary = attempt(() => typeInfoInLibrary(coclass, iid));
        if (fromLibrary) return fromLibrary;
      }
    }
  }
  if (!iid) return null;
  const dispatch = source.tryQueryInterface(IDispatch);
  if (!dispatch) return null;
  using ownedDispatch = dispatch;
  using typeInfo = dispatchTypeInfo(ownedDispatch);
  return typeInfo ? attempt(() => typeInfoInLibrary(typeInfo, iid)) : null;
}

function typeInfoInLibrary(typeInfo: TypeInfoPointer, iid: GUID): TypeInfoPointer {
  const library = new ctypes.c_void_p();
  const index = new ctypes.c_uint32();
  sync(typeInfo).GetContainingTypeLib(ctypes.byref(library), ctypes.byref(index));
  using typeLib = COMPointer.fromAddress(library.value!, ITypeLib);
  const out = new ctypes.c_void_p();
  sync(typeLib).GetTypeInfoOfGuid(iid, ctypes.byref(out));
  return COMPointer.fromAddress(out.value!, ITypeInfo);
}

interface ResolvedInterface {
  readonly definition: InterfaceDef<any>;
  readonly names: Map<number, string>;
}

function resolveInterface(source: COMPointer, option: GetEventsOptions["interface"]): ResolvedInterface {
  const names = new Map<number, string>();
  let definition: InterfaceDef<any> | null = null;
  let iid: GUID | null = null;
  if (option && typeof option === "object" && !(option instanceof GUID)) {
    definition = option;
    iid = option.iid;
    for (const spec of [...option.allMethods, ...option.dispMethods]) if (spec.dispid !== null) names.set(spec.dispid, spec.property);
  } else if (option) {
    iid = toGUID(option as GUIDLike);
  } else {
    const provider = source.tryQueryInterface(IProvideClassInfo2);
    if (provider) {
      using owned = provider;
      iid = attempt(() => sync(owned).GetGUID(GUIDKIND_DEFAULT_SOURCE_DISP_IID));
    }
  }
  using typeInfo = eventTypeInfo(source, iid);
  let typeName: string | null = null;
  if (typeInfo) {
    iid ??= withTypeAttr(typeInfo, (attr) => new GUID(attr.guid));
    typeName = attempt(() => documentation(typeInfo).name);
    const members = attempt(() => typeMembers(typeInfo));
    for (const member of members?.values() ?? []) if (member.name && !names.has(member.memid)) names.set(member.memid, member.name);
  }
  if (!iid) throw new COMError(E_NOINTERFACE, { text: "The source has no default event interface: pass options.interface" });
  if (!definition) {
    // A registered definition (e.g. the outgoing interface of a JavaScript source) names its DISPIDs.
    const known = findInterface(iid);
    if (known && derivesFrom(known, IDispatch.iid)) {
      definition = known;
      for (const spec of [...known.allMethods, ...known.dispMethods]) if (spec.dispid !== null && !names.has(spec.dispid)) names.set(spec.dispid, spec.property);
    }
  }
  definition ??= COMInterface({ name: typeName ?? "EventSink", iid, extends: IDispatch, register: false });
  return { definition, names };
}

const sinkClasses = new WeakMap<InterfaceDef<any>, new (handler: Handler, names: ReadonlyMap<number, string>) => COMObject>();

function sinkClassFor(definition: InterfaceDef<any>): new (handler: Handler, names: ReadonlyMap<number, string>) => COMObject {
  let type = sinkClasses.get(definition);
  if (type) return type;
  const prefix = definition.name;
  type = class EventSink extends COMObject {
    static _com_interfaces_ = [definition];
    readonly handler: Handler;
    readonly names: ReadonlyMap<number, string>;

    constructor(handler: Handler, names: ReadonlyMap<number, string>) {
      super();
      this.handler = handler;
      this.names = names;
      // Vtable (non-dispatch) event interfaces: forward each method to the handler.
      for (let current: InterfaceDef<any> | null = definition; current && !current.iid.equals(IDispatch.iid) && !current.iid.equals(IUnknown.iid); current = current.base) {
        for (const spec of current.methods) {
          Object.defineProperty(this, spec.name, { value: (...args: unknown[]) => this.#deliver(spec.property, args), configurable: true });
        }
      }
    }

    #deliver(name: string, args: unknown[]): unknown {
      const handler = this.handler;
      const method = handler[name] ?? handler[`${prefix}_${name}`];
      if (typeof method === "function") return method.apply(handler, args);
      if (typeof handler.onEvent === "function") return handler.onEvent(name, args);
      return undefined;
    }

    _invoke_(dispid: number, _flags: number, args: InvokeArguments): unknown {
      // ByRef event parameters (e.g. Cancel) are passed as ByRef: set .value to answer.
      return this.#deliver(
        this.names.get(dispid) ?? `DISPID_${dispid}`,
        args.positional.map((value, index) => args.references[index] ?? value),
      );
    }
  };
  sinkClasses.set(definition, type);
  return type;
}

/**
 * Connect `handler` to the events of `source`. Python equivalent:
 * `comtypes.client.GetEvents`.
 *
 * Events call `handler[Name](...args)` (or `handler[Interface_Name]`, or
 * `handler.onEvent(name, args)`); unhandled events are ignored. Dispatch
 * arguments are dynamic objects lent for the call. The handler's result is
 * returned to the source. Out-of-process sources deliver events through
 * window messages: keep a message pump running (`startMessagePump`).
 */
export function GetEvents(source: unknown, handler: object, options: GetEventsOptions = {}): EventConnection {
  using pointer = sourcePointer(source);
  const { definition, names } = resolveInterface(pointer, options.interface);
  using container = pointer.QueryInterface(IConnectionPointContainer);
  const point = sync(container).FindConnectionPoint(definition.iid);
  if (!point) throw new COMError(E_NOINTERFACE, { text: "FindConnectionPoint returned NULL" });
  const Sink = sinkClassFor(definition);
  const sink = new Sink(handler as Handler, names);
  let cookie: number;
  try {
    cookie = sync(point).Advise(sink);
  } catch (error) {
    point.Release();
    throw error;
  }
  let connected = true;
  const disconnect = () => {
    if (!connected) return;
    connected = false;
    try {
      sync(point).Unadvise(cookie);
    } finally {
      point.Release();
    }
  };
  return {
    cookie,
    interface: definition,
    names,
    get connected() {
      return connected;
    },
    disconnect,
    [Symbol.dispose]: disconnect,
  };
}

/**
 * Print every event of `source` (`console.log` by default). Python
 * equivalent: `comtypes.client.ShowEvents`.
 */
export function ShowEvents(source: unknown, options: GetEventsOptions & { readonly log?: (line: string) => void } = {}): EventConnection {
  const log = options.log ?? ((line: string) => console.log(line));
  const handler = {
    onEvent(name: string, args: unknown[]) {
      log(`Event ${name}(${args.map((arg) => inspect(arg, { depth: 1 })).join(", ")})`);
    },
  };
  return GetEvents(source, handler, options);
}
