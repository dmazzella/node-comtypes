import type { COMPointer } from "../pointer.ts";
import type { ComObject } from "../commethod.ts";
import type { TypeLibModule } from "../typelib/build.ts";
import { getModuleSync } from "../typelib/getmodule.ts";
import { typeLibOf } from "../typelib/load.ts";

/** Constants of a type library by name: enum members, module constants and enums. */
export type ConstantsObject = { readonly [name: string]: any };

const cache = new WeakMap<TypeLibModule, ConstantsObject>();

function isEnum(value: unknown): value is Readonly<Record<string, number>> {
  return typeof value === "object" && value !== null && Object.getPrototypeOf(value) === Object.prototype && Object.values(value).every((member) => typeof member === "number");
}

function isConstant(value: unknown): boolean {
  return typeof value === "number" || typeof value === "string" || typeof value === "bigint" || typeof value === "boolean" || value instanceof Date;
}

function constantsOf(module: TypeLibModule): ConstantsObject {
  let constants = cache.get(module);
  if (constants) return constants;
  const values: Record<string, unknown> = {};
  const enums: [string, unknown][] = [];
  for (const [name, value] of Object.entries(module)) {
    if (name === "library") continue;
    if (isEnum(value)) {
      enums.push([name, value]);
      for (const [member, number] of Object.entries(value)) if (!(member in values)) values[member] = number;
    } else if (isConstant(value) && !(name in values)) {
      values[name] = value;
    }
  }
  // Enum (and alias) names after the members: c.SpecialFolderConst.TemporaryFolder.
  for (const [name, value] of enums) if (!(name in values)) values[name] = value;
  const library = module.library.name ?? module.library.libid;
  constants = new Proxy(Object.freeze(values), {
    get(target, property) {
      if (typeof property === "symbol" || property in target) return target[property as string];
      // Not thenable, not serialized specially.
      if (property === "then" || property === "toJSON") return undefined;
      throw new TypeError(`${library} has no constant or enum ${property}`);
    },
  });
  cache.set(module, constants);
  return constants;
}

/**
 * The constants of an object's type library (through `IProvideClassInfo` or
 * `IDispatch::GetTypeInfo`): enum members and module constants by name, and
 * enums by name. Unknown names throw. Python equivalent:
 * `comtypes.client.Constants`.
 *
 * ```ts
 * using fso = CreateObject("Scripting.FileSystemObject");
 * const c = Constants(fso);
 * fso.GetSpecialFolder(c.TemporaryFolder);   // or c.SpecialFolderConst.TemporaryFolder
 * ```
 */
export function Constants(object: COMPointer | ComObject): ConstantsObject {
  using library = typeLibOf(object);
  return constantsOf(getModuleSync(library));
}
