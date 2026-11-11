import * as ctypes from "node-ctypes";
import { mkdirSync, writeFileSync, existsSync } from "node:fs";
import { join } from "node:path";
import { DECIMAL } from "../automation/decimal.ts";
import { VARIANT } from "../automation/variant.ts";
import { classifyParamType, resolveParamInterface } from "../commethod.ts";
import type { ComParamType, ParamInfo, ParamKind } from "../commethod.ts";
import { GUID } from "../guid.ts";
import { HRESULT } from "../interface.ts";
import type { InterfaceDef, MethodSpec } from "../interface.ts";
import { buildModule, builtinInterfaceDefs } from "./build.ts";
import type { TypeLibModule } from "./build.ts";
import type { CoClass } from "./coclass.ts";
import type { TypeLibModel } from "./model.ts";

export interface GenerateOptions {
  /** Module specifier of node-comtypes in the generated imports (default `"@node-ctypes/comtypes"`). */
  readonly importFrom?: string;
}

const RESERVED = new Set(
  "break case catch class const continue debugger default delete do else enum export extends false finally for function if import in instanceof new null return super switch this throw true try typeof var void while with yield let static implements interface package private protected public await arguments eval undefined NaN Infinity".split(
    " ",
  ),
);

/** A valid, non-reserved identifier for an export or a parameter. */
export function identifier(name: string): string {
  let result = name.replace(/[^A-Za-z0-9_$]/g, "_");
  if (!result || /^\d/.test(result)) result = `_${result}`;
  return RESERVED.has(result) ? `${result}_` : result;
}

/** A property key in a type literal: bare when possible, quoted otherwise. */
function key(name: string): string {
  return /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(name) ? name : JSON.stringify(name);
}

const NUMBER_TYPES = new Set<unknown>([ctypes.c_int8, ctypes.c_uint8, ctypes.c_int16, ctypes.c_uint16, ctypes.c_int32, ctypes.c_uint32, ctypes.c_float, ctypes.c_double]);
const BIGINT_TYPES = new Set<unknown>([ctypes.c_int64, ctypes.c_uint64, ctypes.c_ssize_t, ctypes.c_size_t]);

/** TypeScript source of a generated module. */
class Writer {
  readonly module: TypeLibModule;
  readonly model: TypeLibModel;
  /** Definition (interface, record class, coclass) -> export name. */
  readonly names = new Map<unknown, string>();
  /** node-comtypes names imported as types (aliased with `$`). */
  readonly imports = new Set<string>(["InterfaceDef", "InterfaceMethods", "InterfacePointer", "SyncInterfaceMethods", "ComObject"]);

  constructor(model: TypeLibModel) {
    this.model = model;
    this.module = buildModule(model);
    for (const [name, value] of Object.entries(this.module)) {
      if (name === "library" || this.names.has(value)) continue;
      if (typeof value === "object" || typeof value === "function") this.names.set(value, identifier(name));
    }
  }

  use(name: string): string {
    this.imports.add(name);
    return `${name}$`;
  }

  /** Type expression of an interface definition. */
  interfaceType(definition: InterfaceDef<any>): string {
    const builtin = builtinInterfaceDefs.get(definition.iid.toString());
    if (builtin && builtin.definition === definition) return `typeof ${this.use(builtin.name)}`;
    const name = this.names.get(definition);
    return name ? `typeof ${name}` : `typeof ${this.use("IDispatch")}`;
  }

  simple(type: unknown, direction: "in" | "out"): string {
    if (NUMBER_TYPES.has(type)) return "number";
    if (BIGINT_TYPES.has(type)) return direction === "in" ? "number | bigint" : "bigint";
    if (type === ctypes.wintypes.VARIANT_BOOL) return direction === "in" ? "boolean | number" : "boolean";
    if (type === ctypes.c_char_p) return direction === "in" ? "string | Buffer | null" : "string | null";
    if (type === ctypes.c_wchar_p) return "string | null";
    if (type === ctypes.c_void_p) return direction === "in" ? "unknown" : "bigint | null";
    return "unknown";
  }

  struct(type: unknown, direction: "in" | "out"): string {
    if (type === GUID) return direction === "in" ? this.use("GUIDLike") : this.use("GUID");
    if (type === DECIMAL) return direction === "in" ? `${this.use("DECIMAL")} | Record<string, unknown>` : this.use("DECIMAL");
    const name = this.names.get(type);
    if (!name) return direction === "in" ? "unknown" : "ctypes$.Structure";
    return direction === "in" ? `${name} | Partial<${name}Fields>` : name;
  }

  /** TypeScript type of a parameter kind, as accepted (`in`) or returned (`out`). */
  kind(kind: ParamKind, direction: "in" | "out", self: InterfaceDef<any>): string {
    switch (kind.kind) {
      case "bstr":
        return direction === "in" ? "string | null" : "string";
      case "lpolestr":
        return "string | null";
      case "refguid":
        return direction === "in" ? this.use("GUIDLike") : this.use("GUID");
      case "safearray":
        return direction === "in" ? `ArrayLike<unknown> | ${this.use("Typed")} | null` : "unknown[] | null";
      case "interface": {
        if (direction === "in") return "ComObject$ | null";
        const iface = resolveParamInterface(kind, self);
        // Plain IDispatch outputs are converted at run time (best interface, or a dynamic object), as OutValue says.
        if (builtinInterfaceDefs.get(iface.iid.toString())?.name === "IDispatch") return `${this.use("AutomationObject")} | null`;
        return `InterfacePointer$<${this.interfaceType(iface)}> | null`;
      }
      case "date":
        return direction === "in" ? "Date | number" : "Date";
      case "currency":
        return direction === "in" ? "number | string | bigint" : "number";
      case "variant":
        return "unknown";
      case "struct":
        return this.struct(kind.type, direction);
      case "simple":
        return this.simple(kind.type, direction);
      case "pointer":
        return direction === "in" ? "unknown" : this.kind(kind.base, "out", self);
      case "raw":
        return "unknown";
    }
  }

  valueType(type: ComParamType | null, direction: "in" | "out", self: InterfaceDef<any>): string {
    if (type === null) return "undefined";
    if (type === VARIANT) return "unknown";
    return this.kind(classifyParamType(type), direction, self);
  }

  /** `(a: A, b?: B)`: optional parameters only at the end. */
  parameters(params: readonly ParamInfo[], self: InterfaceDef<any>): string {
    const inputs = params.filter((param) => param.input);
    let firstOptional = inputs.length;
    while (firstOptional > 0 && inputs[firstOptional - 1].optional) firstOptional--;
    const used = new Set<string>();
    return inputs
      .map((param, index) => {
        let name = identifier(param.name);
        while (used.has(name)) name = `${name}_`;
        used.add(name);
        const type = this.valueType(param.type, "in", self);
        if (index >= firstOptional) return `${name}?: ${type}`;
        return param.optional ? `${name}: ${type} | undefined` : `${name}: ${type}`;
      })
      .join(", ");
  }

  /** Result of a COMMETHOD (outputs) or a dispatch member (value type). */
  result(spec: MethodSpec, self: InterfaceDef<any>): string {
    const params = (spec as MethodSpec & { readonly params?: readonly ParamInfo[] }).params;
    if (spec.dispatch) return this.valueType((spec as MethodSpec & { readonly valueType: ComParamType | null }).valueType, "out", self);
    if (!params) return spec.restype === HRESULT ? "number" : "unknown";
    const outputs = params.filter((param) => param.output).map((param) => this.valueType(param.type, "out", self));
    if (!outputs.length) return spec.restype === HRESULT ? "number" : spec.restype === null ? "undefined" : this.simple(spec.restype, "out");
    return outputs.length === 1 ? outputs[0] : `[${outputs.join(", ")}]`;
  }

  /** Signature of a member; methods (kind `method`) return promises unless `synchronous`. */
  signature(spec: MethodSpec, self: InterfaceDef<any>, synchronous = true): string {
    const params = (spec as MethodSpec & { readonly params?: readonly ParamInfo[] }).params;
    const result = params ? this.result(spec, self) : "number";
    const returned = synchronous || spec.kind !== "method" ? result : `Promise<${result}>`;
    return params ? `(${this.parameters(params, self)}): ${returned}` : `(...args: unknown[]): ${returned}`;
  }

  /** Members as installed on pointers (`synchronous`: on the `sync()` views). */
  interfaceBody(definition: InterfaceDef<any>, synchronous: boolean): string[] {
    const ownDisp = definition.dispMethods.slice(definition.base?.dispMethods.length ?? 0);
    const members = [...definition.methods, ...ownDisp];
    const lines: string[] = [];
    const memberNames = new Set<string>();
    for (const spec of members) {
      memberNames.add(spec.name);
      lines.push(`  ${key(spec.name)}${this.signature(spec, definition, synchronous)};`);
      if (spec.dispatch && spec.kind === "propget" && spec.writable) {
        memberNames.add(`_set_${spec.property}`);
        lines.push(`  ${key(`_set_${spec.property}`)}(value: ${this.valueType((spec as MethodSpec & { readonly valueType: ComParamType }).valueType, "in", definition)}): undefined;`);
      }
    }
    // Properties, as installed at runtime (see PropertiesOf).
    const groups = new Map<string, { get?: MethodSpec; put?: MethodSpec; writable?: boolean }>();
    for (const spec of members) {
      if (spec.kind === "method") continue;
      const group = groups.get(spec.property) ?? {};
      if (spec.kind === "propget") {
        group.get = spec;
        if (spec.writable) group.writable = true;
      } else if ((spec.inputArguments ?? 0) >= 1 && (spec.requiredArguments ?? 0) <= 1) {
        group.put ??= spec;
        group.writable = true;
      }
      groups.set(spec.property, group);
    }
    for (const [property, { get, put, writable }] of groups) {
      if (memberNames.has(property)) continue;
      if (get && (get.requiredArguments ?? 0) > 0) {
        lines.push(`  ${key(property)}${this.signature(get, definition)};`);
      } else if (get) {
        lines.push(`  ${writable ? "" : "readonly "}${key(property)}: ${this.result(get, definition)};`);
      } else if (put) {
        const params = (put as MethodSpec & { readonly params?: readonly ParamInfo[] }).params ?? [];
        const value = params.filter((param) => param.input).at(-1);
        lines.push(`  ${key(property)}: ${value ? this.valueType(value.type, "in", definition) : "unknown"};`);
      }
      memberNames.add(property);
    }
    return lines;
  }

  render(options: GenerateOptions): string {
    const from = options.importFrom ?? "@node-ctypes/comtypes";
    const body: string[] = [];
    const exported = new Set<string>(["library"]);
    const exportName = (name: string) => {
      const id = identifier(name);
      if (exported.has(id)) return null;
      exported.add(id);
      return id;
    };
    for (const type of this.model.types) {
      const value = this.module[type.name];
      switch (type.kind) {
        case "enum": {
          const name = exportName(type.name);
          if (!name) break;
          const members = type.members.map((member) => `${key(member.name)}: ${member.value}`);
          const values = [...new Set(type.members.map((member) => member.value))];
          body.push(
            `/** ${type.doc ?? `Enum ${type.name}`} */`,
            `export const ${name} = module$[${JSON.stringify(type.name)}] as { readonly ${members.join("; readonly ")} };`,
            `export type ${name} = ${values.length ? values.join(" | ") : "number"};`,
            "",
          );
          break;
        }
        case "record":
        case "union": {
          if (value === GUID) break;
          const name = exportName(type.name);
          if (!name) break;
          const fields = ((value as { _fields_?: readonly (readonly [string, unknown])[] })._fields_ ?? []).map(([field, fieldType]) => {
            const fieldName = key(field);
            if (fieldType === VARIANT) return `${fieldName}: ${this.use("VARIANT")}`;
            if (typeof fieldType === "function" && (fieldType as { prototype?: unknown }).prototype instanceof ctypes.Structure) return `${fieldName}: ${this.struct(fieldType, "out")}`;
            return `${fieldName}: ${this.simple(fieldType, "out")}`;
          });
          body.push(
            `/** ${type.doc ?? `Record ${type.name}`} */`,
            `export interface ${name}Fields {`,
            ...fields.map((field) => `  ${field};`),
            `}`,
            `export const ${name} = module$[${JSON.stringify(type.name)}] as unknown as { new (init?: Partial<${name}Fields> | Buffer): ${name}; readonly size: number };`,
            `export type ${name} = ctypes$.Structure & ${name}Fields;`,
            "",
          );
          break;
        }
        case "interface":
        case "dispinterface": {
          const definition = value as InterfaceDef<any>;
          const name = exportName(type.name);
          if (!name) break;
          const builtin = builtinInterfaceDefs.get(definition.iid.toString());
          if (builtin && builtin.definition === definition) {
            body.push(`export const ${name} = module$[${JSON.stringify(type.name)}] as typeof ${this.use(builtin.name)};`, "");
            break;
          }
          const own = [...new Set([...definition.methods, ...definition.dispMethods].map((spec) => [spec.name, spec.property]).flat())];
          const omitted = own.length ? own.map((member) => JSON.stringify(member)).join(" | ") : "never";
          const base = definition.base ? `Omit<InterfaceMethods$<${this.interfaceType(definition.base)}>, ${omitted}>` : "{}";
          const syncBase = definition.base ? `Omit<SyncInterfaceMethods$<${this.interfaceType(definition.base)}>, ${omitted}>` : "{}";
          body.push(
            `/** ${type.doc ?? `${type.kind === "dispinterface" ? "Dispinterface" : "Interface"} ${type.name}`} (${definition.iid}) */`,
            `export interface ${name}Methods extends ${base} {`,
            ...this.interfaceBody(definition, false),
            `}`,
            `/** ${type.name} members on the \`sync()\` views: every method synchronous. */`,
            `export interface ${name}SyncMethods extends ${syncBase} {`,
            ...this.interfaceBody(definition, true),
            `}`,
            `export const ${name} = module$[${JSON.stringify(type.name)}] as InterfaceDef$<${name}Methods, ${name}SyncMethods>;`,
            "",
          );
          break;
        }
        case "coclass": {
          const coclass = value as CoClass;
          const name = exportName(type.name);
          if (!name) break;
          body.push(`/** ${type.doc ?? `CoClass ${type.name}`} (${coclass.clsid}) */`, `export const ${name} = module$[${JSON.stringify(type.name)}] as ${this.use("CoClass")}<${this.interfaceType(coclass.defaultInterface)}>;`, "");
          break;
        }
        case "alias":
          // After every target: see below.
          break;
        case "module":
          for (const constant of type.constants) {
            const name = exportName(constant.name);
            if (!name) continue;
            const constantValue = this.module[constant.name];
            const tsType =
              typeof constantValue === "bigint" ? "bigint" : constantValue instanceof Date ? "Date" : typeof constantValue === "number" || typeof constantValue === "string" || typeof constantValue === "boolean" ? typeof constantValue : "unknown";
            body.push(`export const ${name} = module$[${JSON.stringify(constant.name)}] as ${tsType};`);
          }
          body.push("");
          break;
      }
    }
    // Aliases of enums, records and interfaces, after their targets.
    for (const type of this.model.types) {
      if (type.kind !== "alias") continue;
      const value = this.module[type.name];
      const target = value === undefined ? undefined : this.names.get(value);
      if (!target || target === identifier(type.name)) continue;
      const name = exportName(type.name);
      if (!name) continue;
      const isInterface = typeof value === "object" && value !== null && "Pointer" in value;
      body.push(`export const ${name} = ${target};`, ...(isInterface ? [`export type ${name}Methods = ${target}Methods;`, `export type ${name}SyncMethods = ${target}SyncMethods;`] : [`export type ${name} = ${target};`]), "");
    }
    const typeImports = [...this.imports, "TypeLibModel", "TypeLibModule"].sort().map((name) => `${name} as ${name}$`);
    const header = [
      `// Generated by node-comtypes from the type library ${this.model.name} ${this.model.major}.${this.model.minor} (${this.model.libid}, LCID ${this.model.lcid}). Do not edit.`,
      `import type * as ctypes$ from "node-ctypes";`,
      `import { buildModule as buildModule$ } from ${JSON.stringify(from)};`,
      `import type { ${typeImports.join(", ")} } from ${JSON.stringify(from)};`,
      "",
      `const model$ = JSON.parse(${JSON.stringify(JSON.stringify(this.model))}) as TypeLibModel$;`,
      `const module$: TypeLibModule$ = buildModule$(model$);`,
      "",
      `/** Library attributes: name, LIBID, version, LCID. */`,
      `export const library = module$.library;`,
      "",
    ];
    // ctypes$ is only needed by records.
    const source = [...header, ...body].join("\n");
    return source.includes("ctypes$.") ? source : source.replace(`import type * as ctypes$ from "node-ctypes";\n`, "");
  }
}

/**
 * TypeScript source of a module for a type library model: the model, built at
 * import time by `buildModule` (the same code as `GetModule`), plus static
 * types for every enum, record, interface, dispinterface, coclass and constant.
 */
export function generateModule(model: TypeLibModel, options: GenerateOptions = {}): string {
  return new Writer(model).render(options);
}

/** File name of the generated module of a model. */
export function moduleFileName(model: TypeLibModel): string {
  return `${identifier(model.name)}.ts`;
}

/** Write the generated module to `directory`; returns its path. */
export function writeModule(model: TypeLibModel, directory: string, options: GenerateOptions & { readonly overwrite?: boolean } = {}): string {
  mkdirSync(directory, { recursive: true });
  const path = join(directory, moduleFileName(model));
  if (options.overwrite !== false || !existsSync(path)) writeFileSync(path, generateModule(model, options));
  return path;
}
