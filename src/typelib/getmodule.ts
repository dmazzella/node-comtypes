import { buildModule } from "./build.ts";
import { writeModule } from "./generate.ts";
import type { TypeLibModule } from "./build.ts";
import { loadTypeLib, withLibAttr } from "./load.ts";
import type { TypeLibSource } from "./load.ts";
import { readTypeLib } from "./model.ts";
import type { TypeLibModel } from "./model.ts";

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

const moduleKey = (libid: string, major: number, minor: number, lcid: number) => `${libid}:${major}.${minor}:${lcid}`;

function isModel(value: unknown): value is TypeLibModel {
  return typeof value === "object" && value !== null && "types" in value && "libid" in value && !("address" in value);
}

/**
 * Load a type library and build its definitions (cached per LIBID, version
 * and LCID). Python equivalent: `comtypes.client.GetModule`.
 *
 * ```ts
 * const Scripting = GetModule("scrrun.dll");
 * using dict = CreateObject(Scripting.Dictionary);   // typed IDictionary
 * ```
 *
 * The result is untyped (`Record<string, unknown>`); generate a module
 * (`node-comtypes generate`) for static types.
 */
export function GetModule(source: TypeLibSource | TypeLibModel, options: GetModuleOptions = {}): TypeLibModule {
  const cacheDir = options.cacheDir === undefined ? process.env.NODE_COMTYPES_CACHE : options.cacheDir;
  let model: TypeLibModel;
  let path: string | undefined;
  if (isModel(source)) {
    model = source;
  } else {
    using library = loadTypeLib(source);
    // Look up the cache before reading the whole library.
    const libraryKey = withLibAttr(library, (attr) => moduleKey(attr.guid.toString(), attr.wMajorVerNum, attr.wMinorVerNum, attr.lcid));
    const cached = modules.get(libraryKey);
    if (cached && (!cacheDir || written.has(`${libraryKey}|${cacheDir}`))) return cached;
    model = readTypeLib(library);
    if (typeof source === "string" && /\.(dll|tlb|olb|ocx|exe)(\\\d+)?$/i.test(source)) path = source;
  }
  const key = moduleKey(model.libid, model.major, model.minor, model.lcid);
  let module = modules.get(key);
  if (!module) {
    module = buildModule(model, { path });
    modules.set(key, module);
  }
  if (cacheDir && !written.has(`${key}|${cacheDir}`)) {
    writeModule(model, cacheDir, { overwrite: false });
    written.add(`${key}|${cacheDir}`);
  }
  return module;
}
