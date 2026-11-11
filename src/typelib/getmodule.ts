import { buildModuleSteps } from "./build.ts";
import { writeModule } from "./generate.ts";
import type { TypeLibModule } from "./build.ts";
import { loadTypeLib, withLibAttr } from "./load.ts";
import type { TypeLibSource } from "./load.ts";
import { readTypeLibSteps } from "./model.ts";
import type { TypeLibModel } from "./model.ts";
import { runSliced, runSync } from "./steps.ts";
import type { Steps } from "./steps.ts";

export interface GetModuleOptions {
  /**
   * Also write the generated TypeScript module (`<Library>.ts`) to this
   * directory when missing (Python: `comtypes.gen`). Default:
   * `NODE_COMTYPES_CACHE`, or no file.
   */
  readonly cacheDir?: string | null;
}

const modules = new Map<string, TypeLibModule>();
/** `key|directory` pairs already written to a cache directory. */
const written = new Set<string>();
/** Asynchronous loads in progress, by library: concurrent calls share one build. */
const loading = new Map<string, Promise<TypeLibModule>>();

const moduleKey = (libid: string, major: number, minor: number, lcid: number) => `${libid}:${major}.${minor}:${lcid}`;

function isModel(value: unknown): value is TypeLibModel {
  return typeof value === "object" && value !== null && "types" in value && "libid" in value && !("address" in value);
}

/** Hooks of an asynchronous load: a load of the same library already in progress, and this load's start. */
interface Sharing {
  pending(key: string): Promise<TypeLibModule> | undefined;
  started(key: string): void;
}

/** Load, read and build a module, yielding after each type; the module cache is shared by both drivers. */
function* moduleSteps(source: TypeLibSource | TypeLibModel, options: GetModuleOptions, sharing?: Sharing): Steps<TypeLibModule> {
  const cacheDir = options.cacheDir === undefined ? process.env.NODE_COMTYPES_CACHE : options.cacheDir;
  const fresh = (key: string) => {
    const cached = modules.get(key);
    return cached && (!cacheDir || written.has(`${key}|${cacheDir}`)) ? cached : undefined;
  };
  let model: TypeLibModel;
  let path: string | undefined;
  if (isModel(source)) {
    model = source;
  } else {
    using library = loadTypeLib(source);
    // Look up the cache before reading the whole library.
    const libraryKey = withLibAttr(library, (attr) => moduleKey(attr.guid.toString(), attr.wMajorVerNum, attr.wMinorVerNum, attr.lcid));
    const cached = fresh(libraryKey);
    if (cached) return cached;
    const pending = sharing?.pending(libraryKey);
    if (pending) return (yield pending) as TypeLibModule;
    sharing?.started(libraryKey);
    model = yield* readTypeLibSteps(library);
    if (typeof source === "string" && /\.(dll|tlb|olb|ocx|exe)(\\\d+)?$/i.test(source)) path = source;
  }
  const key = moduleKey(model.libid, model.major, model.minor, model.lcid);
  let module = modules.get(key);
  if (!module) {
    const built = yield* buildModuleSteps(model, { path });
    // A synchronous load may have built it meanwhile: keep the first.
    module = modules.get(key) ?? built;
    modules.set(key, module);
  }
  if (cacheDir && !written.has(`${key}|${cacheDir}`)) {
    writeModule(model, cacheDir, { overwrite: false });
    written.add(`${key}|${cacheDir}`);
  }
  return module;
}

/**
 * Load a type library and build its definitions (cached per LIBID, version
 * and LCID). Python equivalent: `comtypes.client.GetModule`.
 *
 * ```ts
 * const Scripting = await GetModule("scrrun.dll");
 * using dict = await CreateObject(Scripting.Dictionary);   // typed IDictionary
 * ```
 *
 * The library is read and built in slices of a few milliseconds, yielding to
 * the event loop between them: a large library (e.g. mshtml) does not freeze
 * timers, I/O or an Electron UI. Concurrent calls for the same library share
 * one build. {@link getModuleSync} does the same work at once.
 *
 * The result is untyped (`Record<string, unknown>`); generate a module
 * (`node-comtypes generate`) for static types.
 */
export function GetModule(source: TypeLibSource | TypeLibModel, options: GetModuleOptions = {}): Promise<TypeLibModule> {
  let key: string | null = null;
  // runSliced starts after this function returns: `result` is set by then.
  const result: Promise<TypeLibModule> = runSliced(
    moduleSteps(source, options, {
      pending: (libraryKey) => loading.get(libraryKey),
      started: (libraryKey) => {
        key = libraryKey;
        loading.set(libraryKey, result);
      },
    }),
  ).finally(() => {
    if (key !== null && loading.get(key) === result) loading.delete(key);
  });
  return result;
}

/**
 * {@link GetModule} synchronously: the library is read and built at once, on
 * the calling thread (the cache is shared). For code that cannot wait, such as
 * the results of synchronous members.
 */
export function getModuleSync(source: TypeLibSource | TypeLibModel, options: GetModuleOptions = {}): TypeLibModule {
  return runSync(moduleSteps(source, options));
}
