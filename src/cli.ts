#!/usr/bin/env node
import { mkdirSync, writeFileSync } from "node:fs";
import { basename, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { parseArgs } from "node:util";
import { CoInitializeEx, CoUninitialize } from "./runtime.ts";
import { generateModule, moduleFileName } from "./typelib/generate.ts";
import { loadTypeLib } from "./typelib/load.ts";
import { readTypeLib } from "./typelib/model.ts";
import type { TypeLibModel } from "./typelib/model.ts";

const USAGE = `Usage:
  node-comtypes generate <typelib> [--out <dir>] [--import <specifier>] [--js]
  node-comtypes info <typelib>

<typelib> is a file ("scrrun.dll", "C:\\\\path\\\\lib.tlb", "file.dll\\\\2"),
a LIBID ("{420B2830-E718-11CF-893D-00A0C9054228}") or a ProgID ("Scripting.Dictionary").

generate  writes <Library>.ts (default directory: the current one); with --js,
          <Library>.js and <Library>.d.ts instead (requires the typescript package).
info      prints the library and its types.`;

function read(source: string): TypeLibModel {
  using library = loadTypeLib(source);
  return readTypeLib(library);
}

/** Declaration files of node-comtypes and node-ctypes as seen from this CLI, for output directories outside a project. */
function typePaths(specifier: string): Record<string, string[]> {
  const paths: Record<string, string[]> = {};
  if (specifier === "@node-ctypes/comtypes") paths["@node-ctypes/comtypes"] = [fileURLToPath(new URL("./index.d.ts", import.meta.url))];
  try {
    paths["node-ctypes"] = [fileURLToPath(import.meta.resolve("node-ctypes")).replace(/\.js$/, ".d.ts")];
  } catch {
    // Resolved normally from the output directory.
  }
  return paths;
}

async function emitJavaScript(directory: string, file: string, source: string, specifier: string): Promise<string[]> {
  let ts: typeof import("typescript");
  try {
    ts = (await import("typescript")).default as typeof import("typescript");
  } catch {
    throw new Error("--js requires the typescript package (npm install --save-dev typescript)");
  }
  const input = join(directory, file);
  writeFileSync(input, source);
  const program = ts.createProgram([input], {
    target: ts.ScriptTarget.ES2022,
    module: ts.ModuleKind.NodeNext,
    moduleResolution: ts.ModuleResolutionKind.NodeNext,
    declaration: true,
    skipLibCheck: true,
    strict: true,
    outDir: directory,
    rootDir: directory,
    paths: typePaths(specifier),
  });
  const result = program.emit();
  const errors = ts
    .getPreEmitDiagnostics(program)
    .concat(result.diagnostics)
    .filter((diagnostic) => diagnostic.category === ts.DiagnosticCategory.Error);
  if (errors.length) throw new Error(ts.formatDiagnostics(errors, { getCanonicalFileName: (name) => name, getCurrentDirectory: () => directory, getNewLine: () => "\n" }));
  const { rmSync } = await import("node:fs");
  rmSync(input);
  const stem = basename(file, ".ts");
  return [join(directory, `${stem}.js`), join(directory, `${stem}.d.ts`)];
}

async function main(argv: readonly string[]): Promise<number> {
  const { values, positionals } = parseArgs({
    args: [...argv],
    allowPositionals: true,
    options: { out: { type: "string", short: "o" }, import: { type: "string" }, js: { type: "boolean" }, help: { type: "boolean", short: "h" } },
  });
  const [command, source] = positionals;
  if (values.help || !command) {
    console.log(USAGE);
    return command || values.help ? 0 : 1;
  }
  if (!source || (command !== "generate" && command !== "info")) {
    console.error(USAGE);
    return 1;
  }
  CoInitializeEx();
  try {
    const model = read(source);
    if (command === "info") {
      console.log(`${model.name} ${model.major}.${model.minor} ${model.libid} (LCID ${model.lcid})${model.doc ? ` - ${model.doc}` : ""}`);
      for (const type of model.types) console.log(`  ${type.kind.padEnd(13)} ${type.name}${"iid" in type ? ` ${type.iid}` : "clsid" in type ? ` ${type.clsid}` : ""}`);
      return 0;
    }
    const directory = resolve(values.out ?? ".");
    mkdirSync(directory, { recursive: true });
    const code = generateModule(model, { importFrom: values.import });
    const file = moduleFileName(model);
    const written = values.js ? await emitJavaScript(directory, file, code, values.import ?? "@node-ctypes/comtypes") : [join(directory, file)];
    if (!values.js) writeFileSync(written[0], code);
    for (const path of written) console.log(path);
    return 0;
  } finally {
    CoUninitialize();
  }
}

main(process.argv.slice(2))
  .then(
    (code) => {
      process.exitCode = code;
    },
    (error: unknown) => {
      console.error(error instanceof Error ? error.message : error);
      process.exitCode = 1;
    },
  )
  // One more event loop turn before exiting: exiting right after the work
  // (and CoUninitialize) races Node's teardown with tasks posted by V8's
  // concurrent GC, and can abort with "!(handle->flags & UV_HANDLE_CLOSING)".
  .then(() => new Promise((resolve) => setTimeout(resolve, 0)));
