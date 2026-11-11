import * as ctypes from "node-ctypes";
import { setTimeout as delay } from "node:timers/promises";
import { CLSCTX_LOCAL_SERVER } from "../constants.ts";
import { checkHRESULT } from "../errors.ts";
import { toGUID } from "../guid.ts";
import type { GUID } from "../guid.ts";
import { ole32 } from "../low-level/native.ts";
import { deleteTree, HKEY_CURRENT_USER, writeString } from "../low-level/registry.ts";
import { pumpMessages } from "../pump.ts";
import { CoInitializeEx, CoUninitialize } from "../runtime.ts";
import { ClassFactory, COMObject, classTypeLib, getLiveObjectCount, queryObject } from "./comobject.ts";
import { registerTypeLib, unregisterTypeLib } from "../typelib/load.ts";
import { IUnknown } from "../interfaces/unknown.ts";
import { rawRelease } from "../low-level/vtable.ts";
import { markLocalClass, unmarkLocalClass } from "../executor/classes.ts";

export const REGCLS_SINGLEUSE = 0;
export const REGCLS_MULTIPLEUSE = 1;
export const REGCLS_MULTI_SEPARATE = 2;
export const REGCLS_SUSPENDED = 4;

type ComClass = (new () => COMObject) & typeof COMObject;

function classId(type: ComClass): GUID {
  if (!type._reg_clsid_) throw new TypeError(`${type.name}: static _reg_clsid_ is required`);
  return toGUID(type._reg_clsid_);
}

export interface ClassRegistration extends Disposable {
  readonly cookie: number;
  readonly factory: ClassFactory;
  revoke(): void;
}

/**
 * Register a class factory for `type` (`CoRegisterClassObject`) until
 * revoked. With `CLSCTX_INPROC_SERVER` the class is creatable in this process
 * without registry entries; LocalServer32 servers use `CLSCTX_LOCAL_SERVER`.
 */
export function registerClassObject(type: ComClass, options: { readonly clsctx?: number; readonly flags?: number } = {}): ClassRegistration {
  const factory = new ClassFactory(type);
  const address = queryObject(factory, IUnknown.iid)!;
  const cookie = new ctypes.c_uint32();
  try {
    checkHRESULT(ole32().CoRegisterClassObject(classId(type)._buffer, address, options.clsctx ?? CLSCTX_LOCAL_SERVER, options.flags ?? REGCLS_MULTIPLEUSE, ctypes.byref(cookie)), "CoRegisterClassObject");
  } finally {
    // COM keeps its own reference while registered.
    rawRelease(address);
  }
  // CreateObject creates objects of this class on this thread (see client/index.ts).
  const clsid = classId(type);
  markLocalClass(clsid);
  let registered = true;
  const revoke = () => {
    if (!registered) return;
    registered = false;
    unmarkLocalClass(clsid);
    checkHRESULT(ole32().CoRevokeClassObject(cookie.value), "CoRevokeClassObject");
  };
  return { cookie: cookie.value, factory, revoke, [Symbol.dispose]: revoke };
}

// ── registry (HKEY_CURRENT_USER\Software\Classes) ─────────────────────

const setString = (key: string, name: string | null, value: string) => writeString(HKEY_CURRENT_USER, key, name, value);

function quote(argument: string): string {
  return /[\s"]/.test(argument) ? `"${argument.replace(/"/g, '\\"')}"` : argument;
}

export interface ServerRegistration {
  /** Script run by the server process (COM appends `-Embedding`). */
  readonly script: string;
  /** Default: the current Node.js executable. */
  readonly executable?: string;
  /** Node.js options before the script (e.g. `--import` loaders). */
  readonly execArgv?: readonly string[];
  /** Extra arguments after the script. */
  readonly args?: readonly string[];
  /**
   * Also register the class type library file (`_typelib_path_`, or the file
   * `GetModule` loaded), so that its `[oleautomation]` and dual interfaces get
   * the type library marshaler as proxy/stub. Default `true`.
   */
  readonly typelib?: boolean;
}

/** The type library file to register with a class, if known. */
function typeLibFile(type: ComClass): string | null {
  return type._typelib_path_ ?? classTypeLib(type)?.path ?? null;
}

/**
 * Register `type` as a per-user LocalServer32 (`HKCU\Software\Classes`):
 * `CLSID\{clsid}\LocalServer32` runs `node script args...`, plus
 * - the ProgID keys (`_reg_progid_`, `_reg_novers_progid_`);
 * - `TypeLib` and `Version` for the class type library (`_reg_typelib_`, or
 *   the library of interfaces built by `GetModule`), whose file is registered
 *   too (`registerTypeLib`) unless `typelib: false`.
 *
 * Undo with {@link unregisterServer}. Python equivalent: `comtypes.server.register`.
 */
export function registerServer(type: ComClass, registration: ServerRegistration): void {
  const clsid = classId(type).toString();
  const key = `Software\\Classes\\CLSID\\${clsid}`;
  const command = [registration.executable ?? process.execPath, ...(registration.execArgv ?? []), registration.script, ...(registration.args ?? [])].map(quote).join(" ");
  const description = type._reg_desc_ ?? type._reg_progid_ ?? type.name;
  setString(key, null, description);
  setString(`${key}\\LocalServer32`, null, command);
  if (type._reg_progid_) setString(`${key}\\ProgID`, null, type._reg_progid_);
  if (type._reg_novers_progid_) {
    setString(`${key}\\VersionIndependentProgID`, null, type._reg_novers_progid_);
    if (type._reg_progid_) setString(`Software\\Classes\\${type._reg_novers_progid_}\\CurVer`, null, type._reg_progid_);
  }
  for (const progid of [type._reg_progid_, type._reg_novers_progid_]) {
    if (!progid) continue;
    setString(`Software\\Classes\\${progid}`, null, description);
    setString(`Software\\Classes\\${progid}\\CLSID`, null, clsid);
  }
  const library = classTypeLib(type);
  if (library) {
    setString(`${key}\\TypeLib`, null, library.libid);
    // Decimal, unlike the hexadecimal versions under TypeLib\{libid}.
    setString(`${key}\\Version`, null, `${library.major}.${library.minor}`);
    const file = typeLibFile(type);
    if (file && registration.typelib !== false) registerTypeLib(file);
  }
}

export interface UnregisterServerOptions {
  /** Also unregister the class type library (see {@link ServerRegistration.typelib}). Default `true`. */
  readonly typelib?: boolean;
}

/** Remove the per-user keys written by {@link registerServer}. */
export function unregisterServer(type: ComClass, options: UnregisterServerOptions = {}): void {
  deleteTree(HKEY_CURRENT_USER, `Software\\Classes\\CLSID\\${classId(type)}`);
  for (const progid of [type._reg_progid_, type._reg_novers_progid_]) if (progid) deleteTree(HKEY_CURRENT_USER, `Software\\Classes\\${progid}`);
  const library = classTypeLib(type);
  if (library && typeLibFile(type) && options.typelib !== false) {
    try {
      unregisterTypeLib(library);
    } catch {
      // Not registered, or already removed with another class of the library.
    }
  }
}

export interface LocalServerOptions {
  /** Exit after this long without objects once one was created (default 2 s). */
  readonly idleTimeout?: number;
  /** Exit if nothing is created within this time (default 60 s). */
  readonly startupTimeout?: number;
  /** Message pump interval (default 5 ms). */
  readonly interval?: number;
}

/**
 * Serve `types` from this process (the LocalServer32 entry point): initialize
 * an STA, register the class factories, pump messages until idle, then revoke
 * and uninitialize. Python equivalent: `comtypes.server.localserver.run`.
 */
export async function runLocalServer(types: readonly ComClass[], options: LocalServerOptions = {}): Promise<void> {
  CoInitializeEx();
  const registrations = types.map((type) => registerClassObject(type, { clsctx: CLSCTX_LOCAL_SERVER, flags: REGCLS_MULTIPLEUSE }));
  const live = () => getLiveObjectCount((object) => !(object instanceof ClassFactory));
  const locked = () => registrations.some((registration) => registration.factory.locks > 0);
  try {
    const started = performance.now();
    let served = false;
    let idleSince = performance.now();
    for (;;) {
      pumpMessages();
      const busy = live() > 0 || locked();
      if (busy) {
        served = true;
        idleSince = performance.now();
      }
      const now = performance.now();
      if (served && !busy && now - idleSince >= (options.idleTimeout ?? 2000)) break;
      if (!served && now - started >= (options.startupTimeout ?? 60_000)) break;
      await delay(options.interval ?? 5);
    }
  } finally {
    for (const registration of registrations) registration.revoke();
    CoUninitialize();
  }
}

const USAGE = `Usage: node <script> [/regserver | /unregserver | -Embedding]
  /regserver     register the classes for the current user (LocalServer32 running this script)
  /unregserver   remove the registration
  -Embedding     serve the classes (COM starts the server this way)`;

export interface UseCommandLineOptions {
  /** Default `process.argv.slice(2)`. */
  readonly argv?: readonly string[];
  /** Registration of `/regserver`; default: this script (`process.argv[1]`) with the current Node.js options. */
  readonly registration?: Partial<ServerRegistration>;
  /** Options of `-Embedding` ({@link runLocalServer}). */
  readonly server?: LocalServerOptions;
}

/**
 * Entry point of a LocalServer32 script: handles `/regserver`, `/unregserver`
 * and `-Embedding` (with `-` or `/`, in any case); without any of them, prints
 * the usage and returns `false`. Python equivalent:
 * `comtypes.server.register.UseCommandLine`.
 *
 * ```ts
 * class Greeter extends COMObject { static _reg_clsid_ = "{...}"; ... }
 * await UseCommandLine([Greeter]);
 * ```
 */
export async function UseCommandLine(types: readonly ComClass[], options: UseCommandLineOptions = {}): Promise<boolean> {
  const switches = new Set((options.argv ?? process.argv.slice(2)).filter((arg) => /^[-/]/.test(arg)).map((arg) => arg.slice(1).toLowerCase()));
  let handled = false;
  if (switches.has("unregserver")) {
    for (const type of types) unregisterServer(type);
    handled = true;
  }
  if (switches.has("regserver")) {
    const registration: ServerRegistration = { script: process.argv[1], execArgv: process.execArgv, ...options.registration };
    for (const type of types) registerServer(type, registration);
    handled = true;
  }
  if (switches.has("embedding")) {
    await runLocalServer(types, options.server);
    handled = true;
  }
  if (!handled) console.error(USAGE);
  return handled;
}
