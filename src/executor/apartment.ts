// Apartments of the JavaScript ("home") thread and of the COM executor.
// Pointers belong to the apartment that obtained them: the executor (in the
// multithreaded apartment) can use a home pointer directly only when the home
// thread is in the MTA too or the object is agile; otherwise it gets its own
// reference through the Global Interface Table. Interface results of calls
// made on the executor are handed back to a single-threaded home the same way.
import * as ctypes from "node-ctypes";
import { VT_ARRAY, VT_BYREF, VT_DISPATCH, VT_UNKNOWN, VT_VARIANT } from "../automation/constants.ts";
import { APTTYPE_MTA, CLSCTX_INPROC_SERVER } from "../constants.ts";
import { COMError } from "../errors.ts";
import { GUID } from "../guid.ts";
import { CO_E_NOTINITIALIZED } from "../hresult.ts";
import { ole32, oleaut32 } from "../low-level/native.ts";
import { functionCache, rawQueryInterface, rawRelease, vtableEntry } from "../low-level/vtable.ts";
import { apartmentGeneration } from "../tracking.ts";
import { executorAlive, executorGeneration, executorTarget } from "./executor.ts";

const { c_int32, c_uint16, c_uint32, c_void_p } = ctypes;

const CLSID_StdGlobalInterfaceTable = new GUID("{00000323-0000-0000-C000-000000000046}");
const IID_IGlobalInterfaceTable = new GUID("{00000146-0000-0000-C000-000000000046}");
const IID_IAgileObject = new GUID("{94EA2B94-E9CC-49E0-C0FF-EE64CA8F5B90}");
const IID_IMarshal = new GUID("{00000003-0000-0000-C000-000000000046}");
const IID_IClientSecurity = new GUID("{0000013D-0000-0000-C000-000000000046}");
const IID_IDispatch = new GUID("{00020400-0000-0000-C000-000000000046}");
const IID_IUnknown = new GUID("{00000000-0000-0000-C000-000000000046}");
const CLSID_InProcFreeMarshaler = new GUID("{0000033A-0000-0000-C000-000000000046}");
const MSHCTX_INPROC = 3;
const MSHLFLAGS_NORMAL = 0;

const RegisterAt = functionCache(() => ctypes.WINFUNCTYPE(c_int32, c_void_p, c_void_p, c_void_p, c_void_p));
const RevokeAt = functionCache(() => ctypes.WINFUNCTYPE(c_int32, c_void_p, c_uint32));
const GetAt = functionCache(() => ctypes.WINFUNCTYPE(c_int32, c_void_p, c_uint32, c_void_p, c_void_p));
const QueryInterfaceAt = functionCache(() => ctypes.WINFUNCTYPE(c_int32, c_void_p, c_void_p, c_void_p));
const ReleaseAt = functionCache(() => ctypes.WINFUNCTYPE(c_uint32, c_void_p));
const GetUnmarshalClassAt = functionCache(() => ctypes.WINFUNCTYPE(c_int32, c_void_p, c_void_p, c_void_p, c_uint32, c_void_p, c_uint32, c_void_p));

/**
 * Apartment of the home thread: `mta` (explicit or implicit: shared with the
 * executor), `sta` (single-threaded, also NA/main STA), or `none` when COM is
 * not initialized there (objects are then called on the home thread).
 */
export function homeApartment(): "mta" | "sta" | "none" {
  const type = new c_int32();
  const qualifier = new c_int32();
  const hr = ole32().CoGetApartmentType(ctypes.byref(type), ctypes.byref(qualifier));
  if (hr === CO_E_NOTINITIALIZED || hr < 0) return "none";
  return type.value === APTTYPE_MTA ? "mta" : "sta";
}

function supports(address: bigint, iid: GUID): bigint | null {
  const [hr, pointer] = rawQueryInterface(address, iid);
  return hr >= 0 && pointer ? pointer : null;
}

/** A standard proxy (to another apartment or process): answered locally by the proxy manager. */
export function isProxy(address: bigint): boolean {
  const security = supports(address, IID_IClientSecurity);
  if (security) rawRelease(security);
  return security !== null;
}

/** Callable from any apartment: `IAgileObject`, or marshaled by the free-threaded marshaler. */
export function isAgile(address: bigint): boolean {
  const agile = supports(address, IID_IAgileObject);
  if (agile) {
    rawRelease(agile);
    return true;
  }
  const marshal = supports(address, IID_IMarshal);
  if (!marshal) return false;
  try {
    const clsid = new GUID();
    const hr = GetUnmarshalClassAt(vtableEntry(marshal, 3))(marshal, IID_IUnknown._buffer, address, MSHCTX_INPROC, null, MSHLFLAGS_NORMAL, clsid._buffer);
    return hr >= 0 && clsid.equals(CLSID_InProcFreeMarshaler);
  } finally {
    rawRelease(marshal);
  }
}

// The Global Interface Table of the home thread, per apartment generation.
let homeTable: { readonly address: bigint; readonly generation: number } | null = null;

function homeGlobalTable(): bigint {
  const generation = apartmentGeneration();
  if (homeTable?.generation === generation) return homeTable.address;
  const out = new c_void_p();
  const hr = ole32().CoCreateInstance(CLSID_StdGlobalInterfaceTable._buffer, null, CLSCTX_INPROC_SERVER, IID_IGlobalInterfaceTable._buffer, ctypes.byref(out));
  if (hr < 0 || !out.value) throw new COMError(hr, { method: "CoCreateInstance(StdGlobalInterfaceTable)" });
  homeTable = { address: out.value, generation };
  return out.value;
}

// The executor's own table, per executor generation (created on the executor).
let executorTable: { readonly address: Promise<bigint>; readonly generation: number } | null = null;

function executorGlobalTable(): Promise<bigint> {
  const target = executorTarget();
  const generation = executorGeneration();
  if (executorTable?.generation === generation) return executorTable.address;
  const out = new c_void_p();
  const address = ole32()
    .CoCreateInstance.callOn(target, CLSID_StdGlobalInterfaceTable._buffer, null, CLSCTX_INPROC_SERVER, IID_IGlobalInterfaceTable._buffer, ctypes.byref(out))
    .then((hr) => {
      if (hr < 0 || !out.value) throw new COMError(hr, { method: "CoCreateInstance(StdGlobalInterfaceTable)" });
      return out.value;
    });
  executorTable = { address, generation };
  return address;
}

/** Release a reference owned by the executor thread, after the calls already queued. */
export function releaseOnExecutor(address: bigint, generation: number): void {
  // An executor that has stopped took its references with its apartment.
  if (!executorAlive(generation)) return;
  ReleaseAt(vtableEntry(address, 2))
    .callOn(executorTarget(), address)
    .catch(() => undefined);
}

/**
 * A reference to the home pointer `address` usable on the executor (owned by
 * the executor), through the Global Interface Table; `null` when the
 * interface cannot be marshaled (no proxy/stub).
 */
export async function referenceForExecutor(address: bigint, iid: GUID): Promise<bigint | null> {
  const table = homeGlobalTable();
  const cookie = new c_uint32();
  if (RegisterAt(vtableEntry(table, 3))(table, address, iid._buffer, ctypes.byref(cookie)) < 0) return null;
  try {
    const executorTableAddress = await executorGlobalTable();
    const out = new c_void_p();
    const hr = await GetAt(vtableEntry(executorTableAddress, 5)).callOn(executorTarget(), executorTableAddress, cookie.value, iid._buffer, ctypes.byref(out));
    return hr >= 0 && out.value ? out.value : null;
  } finally {
    RevokeAt(vtableEntry(table, 4))(table, cookie.value);
  }
}

/** Memory receiving an interface result: an `out` interface pointer, a VARIANT or a `SAFEARRAY*`. */
export interface ResultLocation {
  readonly location: bigint;
  readonly kind: "pointer" | "variant" | "safearray";
  /** Interface of a `pointer` result. */
  readonly iid?: GUID;
}

const VARIANT_POINTER_OFFSET = 8;

/** Hand the interface pointer at `location + offset` over from the executor to the home apartment, in place. */
async function handOverPointer(location: bigint, offset: number, interfaceId: GUID): Promise<void> {
  const address = ctypes.readValue(location, c_void_p, offset) as bigint | null;
  if (!address) return;
  const target = executorTarget();
  // Agile objects are valid in both apartments: the reference stays as it is.
  const agile = new c_void_p();
  if ((await QueryInterfaceAt(vtableEntry(address, 0)).callOn(target, address, IID_IAgileObject._buffer, ctypes.byref(agile))) >= 0 && agile.value) {
    void ReleaseAt(vtableEntry(agile.value, 2))
      .callOn(target, agile.value)
      .catch(() => undefined);
    return;
  }
  const executorTableAddress = await executorGlobalTable();
  const cookie = new c_uint32();
  const registered = await RegisterAt(vtableEntry(executorTableAddress, 3)).callOn(target, executorTableAddress, address, interfaceId._buffer, ctypes.byref(cookie));
  // The executor's reference is no longer needed: the table holds its own.
  void ReleaseAt(vtableEntry(address, 2))
    .callOn(target, address)
    .catch(() => undefined);
  if (registered < 0) {
    ctypes.writeValue(location, c_void_p, null, offset);
    throw new COMError(registered, { method: "IGlobalInterfaceTable::RegisterInterfaceInGlobal" });
  }
  const table = homeGlobalTable();
  const out = new c_void_p();
  const hr = GetAt(vtableEntry(table, 5))(table, cookie.value, interfaceId._buffer, ctypes.byref(out));
  RevokeAt(vtableEntry(table, 4))(table, cookie.value);
  ctypes.writeValue(location, c_void_p, hr >= 0 ? out.value : null, offset);
  if (hr < 0) throw new COMError(hr, { method: "IGlobalInterfaceTable::GetInterfaceFromGlobal" });
}

/** Hand over the objects of a SAFEARRAY (of interfaces or VARIANTs), element by element. */
async function handOverArray(psa: bigint | null): Promise<void> {
  if (!psa) return;
  const api = oleaut32();
  const type = new c_uint16();
  if (api.SafeArrayGetVartype(psa, ctypes.byref(type)) < 0) return;
  const elements: Omit<ResultLocation, "location"> | null =
    type.value === VT_DISPATCH ? { kind: "pointer", iid: IID_IDispatch } : type.value === VT_UNKNOWN ? { kind: "pointer", iid: IID_IUnknown } : type.value === VT_VARIANT ? { kind: "variant" } : null;
  if (!elements) return;
  let count = 1;
  const bound = new c_int32();
  for (let dimension = 1; dimension <= api.SafeArrayGetDim(psa); dimension++) {
    api.SafeArrayGetLBound(psa, dimension, ctypes.byref(bound));
    const lower = bound.value;
    api.SafeArrayGetUBound(psa, dimension, ctypes.byref(bound));
    count *= bound.value - lower + 1;
  }
  const size = api.SafeArrayGetElemsize(psa);
  const data = new c_void_p();
  if (count <= 0 || api.SafeArrayAccessData(psa, ctypes.byref(data)) < 0 || !data.value) return;
  try {
    const base = data.value;
    await handOver(Array.from({ length: count }, (_, index) => ({ location: base + BigInt(index * size), ...elements })));
  } finally {
    api.SafeArrayUnaccessData(psa);
  }
}

/** Hand one interface result over from the executor to the home apartment, in place. */
async function handOverOne({ location, kind, iid }: ResultLocation): Promise<void> {
  if (kind === "pointer") return handOverPointer(location, 0, iid ?? IID_IUnknown);
  if (kind === "safearray") return handOverArray(ctypes.readValue(location, c_void_p) as bigint | null);
  const vt = ctypes.readValue(location, c_uint16) as number;
  if (vt & VT_BYREF) return;
  if (vt & VT_ARRAY) return handOverArray(ctypes.readValue(location, c_void_p, VARIANT_POINTER_OFFSET) as bigint | null);
  if (vt === VT_DISPATCH || vt === VT_UNKNOWN) return handOverPointer(location, VARIANT_POINTER_OFFSET, vt === VT_DISPATCH ? IID_IDispatch : IID_IUnknown);
}

/**
 * Hand the interface results of a call made on the executor over to a
 * single-threaded home apartment: each location then holds a reference owned
 * by the home thread (the same one for agile objects).
 */
export async function handOver(results: readonly ResultLocation[]): Promise<void> {
  const outcomes = await Promise.allSettled(results.map((result) => handOverOne(result)));
  const failure = outcomes.find((outcome) => outcome.status === "rejected");
  if (failure) throw (failure as PromiseRejectedResult).reason;
}
