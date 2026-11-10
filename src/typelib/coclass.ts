import type { GUID } from "../guid.ts";
import type { InterfaceDef } from "../interface.ts";

/**
 * A coclass from a type library: `CreateObject(coclass)` returns its default
 * interface. Python equivalent: the `CoClass` subclasses of generated modules.
 */
export interface CoClass<T extends InterfaceDef<any> = InterfaceDef<any>> {
  readonly kind: "coclass";
  readonly name: string;
  readonly clsid: GUID;
  /** Default (non-source) interface; `IDispatch` when unknown. */
  readonly defaultInterface: T;
  /** Default event interface, if any. */
  readonly sourceInterface: InterfaceDef<any> | null;
  readonly interfaces: readonly InterfaceDef<any>[];
  readonly sourceInterfaces: readonly InterfaceDef<any>[];
  readonly canCreate: boolean;
}

export function isCoClass(value: unknown): value is CoClass {
  return typeof value === "object" && value !== null && (value as { kind?: unknown }).kind === "coclass" && "clsid" in value;
}

/** Where to load the type information of interfaces built from a type library. */
export interface TypeLibLocation {
  readonly libid: string;
  readonly major: number;
  readonly minor: number;
  readonly lcid: number;
  /** File the library was loaded from, when not (or not necessarily) registered. */
  readonly path?: string;
}

const locations = new WeakMap<InterfaceDef<any>, TypeLibLocation>();

export function setTypeLibLocation(iface: InterfaceDef<any>, location: TypeLibLocation): void {
  locations.set(iface, location);
}

/** The type library an interface was built from (`GetModule`, generated modules). */
export function typeLibLocationOf(iface: InterfaceDef<any>): TypeLibLocation | undefined {
  return locations.get(iface);
}

/** Record classes built from a type library: how to get their `ITypeInfo` (an owned address), for `IRecordInfo`. */
const recordTypeInfos = new WeakMap<object, () => bigint>();

export function setRecordTypeInfo(type: object, typeInfo: () => bigint): void {
  recordTypeInfos.set(type, typeInfo);
}

export function recordTypeInfoOf(type: object): (() => bigint) | undefined {
  return recordTypeInfos.get(type);
}
