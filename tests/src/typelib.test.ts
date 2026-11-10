import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { after, describe, it } from "node:test";
import { fileURLToPath } from "node:url";
import * as ctypes from "node-ctypes";
import ts from "typescript";

import {
  COMError,
  COMPointer,
  CreateObject,
  Dispatch,
  GUID,
  GetModule,
  TYPE_E_CANTLOADLIBRARY,
  apartment,
  getLivePointerCount,
  loadTypeLib,
  readTypeLib,
  registeredTypeLibPath,
  registeredVersions,
  typeLibOf,
  typeName,
  writeModule,
} from "@node-ctypes/comtypes";
import type { CoClass, InterfaceDef } from "@node-ctypes/comtypes";
import { COMObject } from "@node-ctypes/comtypes/server";

const skip = process.platform !== "win32" && "COM requires Windows";
const SCRIPTING_LIBID = "{420B2830-E718-11CF-893D-00A0C9054228}";
const root = fileURLToPath(new URL("../..", import.meta.url));
const generated = join(root, "tests", ".generated");
const temporary = mkdtempSync(join(tmpdir(), "node-comtypes-typelib-"));
after(() => rmSync(temporary, { recursive: true, force: true }));

describe("loading type libraries", { skip }, () => {
  it("loads from a file, a LIBID, a ProgID and a live object", () => {
    using _ = apartment();
    const names = [];
    for (const source of ["scrrun.dll", SCRIPTING_LIBID, "Scripting.Dictionary", { libid: SCRIPTING_LIBID, major: 1, minor: 0 }] as const) {
      using library = loadTypeLib(source);
      names.push(readTypeLib(library).name);
    }
    using dict = CreateObject("Scripting.Dictionary");
    using fromObject = typeLibOf(dict);
    names.push(readTypeLib(fromObject).name);
    assert.deepEqual(names, ["Scripting", "Scripting", "Scripting", "Scripting", "Scripting"]);
    assert.deepEqual(registeredVersions(SCRIPTING_LIBID)[0], { major: 1, minor: 0 });
    assert.match(registeredTypeLibPath({ libid: SCRIPTING_LIBID }), /scrrun\.dll$/i);
    assert.throws(() => loadTypeLib("node-comtypes-missing.tlb"), COMError);
    assert.throws(
      () => loadTypeLib("{00000000-1111-2222-3333-444444444444}"),
      (error: unknown) => error instanceof COMError && error.hresult === TYPE_E_CANTLOADLIBRARY,
    );
  });

  it("describes types, members and coclasses", () => {
    using _ = apartment();
    using library = loadTypeLib("scrrun.dll");
    const model = readTypeLib(library);
    assert.equal(model.libid, SCRIPTING_LIBID);
    const compare = model.types.find((type) => type.name === "CompareMethod");
    assert.deepEqual(compare?.kind === "enum" && compare.members.map((member) => [member.name, member.value]), [
      ["BinaryCompare", 0],
      ["TextCompare", 1],
      ["DatabaseCompare", 2],
    ]);
    const dictionary = model.types.find((type) => type.name === "IDictionary");
    assert.ok(dictionary?.kind === "interface" && dictionary.dual);
    assert.equal(dictionary.base?.kind === "user" && dictionary.base.name, "IDispatch");
    assert.deepEqual(
      dictionary.methods.slice(0, 4).map((method) => [method.name, method.invkind, method.slot]),
      [
        ["Item", "propputref", 7],
        ["Item", "propput", 8],
        ["Item", "propget", 9],
        ["Add", "method", 10],
      ],
    );
    const coclass = model.types.find((type) => type.name === "Dictionary");
    assert.ok(coclass?.kind === "coclass" && coclass.canCreate && coclass.interfaces[0].default);
    assert.deepEqual(JSON.parse(JSON.stringify(model)), model, "the model is JSON-serializable");
  });
});

describe("GetModule", { skip }, () => {
  it("builds typed definitions, cached per library", () => {
    using _ = apartment();
    const Scripting = GetModule("scrrun.dll");
    assert.equal(GetModule(SCRIPTING_LIBID), Scripting);
    assert.equal(Scripting.library.name, "Scripting");
    assert.deepEqual({ ...(Scripting.SpecialFolderConst as object) }, { WindowsFolder: 0, SystemFolder: 1, TemporaryFolder: 2 });
    const before = getLivePointerCount();
    {
      using dict = CreateObject(Scripting.Dictionary as CoClass) as unknown as Record<string, (...args: unknown[]) => unknown> & COMPointer;
      assert.equal(dict.interface.name, "IDictionary", "the coclass default interface (vtable)");
      dict.Add("a", 1);
      dict.Add("b", [1, 2]);
      assert.equal((dict as unknown as { Count: number }).Count, 2);
      assert.deepEqual(dict.Items(), [1, [1, 2]]);
      assert.throws(
        () => dict.Add("a", 3),
        (error: unknown) => error instanceof COMError && error.method === "IDictionary::Add",
      );
    }
    {
      using fso = CreateObject(Scripting.FileSystemObject as CoClass) as unknown as COMPointer & Record<string, (...args: unknown[]) => COMPointer & Record<string, unknown>>;
      using folder = fso.GetSpecialFolder(2);
      assert.equal(folder.interface.name, "IFolder", "interface results are typed (lazy references)");
      assert.equal(typeof folder.Path, "string");
    }
    assert.equal(getLivePointerCount(), before);
  });

  it("calls dispinterface properties through IDispatch (stdole Font)", () => {
    using _ = apartment();
    const stdole = GetModule("stdole2.tlb");
    const Font = stdole.Font as InterfaceDef<any>;
    assert.equal(Font.allMethods.length, 7);
    class FONTDESC extends ctypes.Structure {
      static _fields_ = [
        ["cbSizeofstruct", ctypes.c_uint32],
        ["lpstrName", ctypes.c_void_p],
        ["cySize", ctypes.c_int64],
        ["sWeight", ctypes.c_int16],
        ["sCharset", ctypes.c_int16],
        ["fItalic", ctypes.c_int32],
        ["fUnderline", ctypes.c_int32],
        ["fStrikethrough", ctypes.c_int32],
      ];
    }
    const name = Buffer.from("Arial\0", "utf16le");
    const description = new FONTDESC();
    description.cbSizeofstruct = ctypes.sizeof(FONTDESC);
    description.lpstrName = ctypes.addressof(name);
    description.cySize = 120000n;
    description.sWeight = 700;
    const OleCreateFontIndirect = new ctypes.WinDLL("oleaut32.dll").func("OleCreateFontIndirect", ctypes.c_int32, [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p]);
    const out = new ctypes.c_void_p();
    assert.equal(OleCreateFontIndirect(description._buffer, Font.iid._buffer, ctypes.byref(out)), 0);
    using font = COMPointer.fromAddress(out.value!, Font) as unknown as COMPointer & { Name: string; Size: number; Bold: boolean; Italic: boolean };
    assert.deepEqual([font.Name, font.Size, font.Bold, font.Italic], ["Arial", 12, true, false]);
    font.Name = "Courier New";
    font.Italic = true;
    font.Size = 10;
    assert.deepEqual([font.Name, font.Italic], ["Courier New", true]);
    assert.ok(Math.abs(font.Size - 10) < 0.5, "GDI rounds the size to the device resolution");
    void name;
  });

  it("writes the generated module to a cache directory", () => {
    using _ = apartment();
    GetModule("scrrun.dll", { cacheDir: temporary });
    assert.ok(existsSync(join(temporary, "Scripting.ts")));
  });

  it("gives JavaScript objects the type information of type library interfaces", () => {
    using _ = apartment();
    const Scripting = GetModule("scrrun.dll");
    class JsDictionary extends COMObject {
      static _com_interfaces_ = [Scripting.IDictionary as InterfaceDef<any>];
      #items = new Map<string, unknown>();
      get Count() {
        return this.#items.size;
      }
      Add(key: string, item: unknown) {
        this.#items.set(key, item);
      }
      Exists(key: string) {
        return this.#items.has(key);
      }
    }
    using dynamic = Dispatch(new JsDictionary());
    assert.equal(typeName(dynamic), "IDictionary");
    dynamic.Add("k", 1);
    assert.equal(dynamic.Count, 1);
    assert.equal(dynamic.Exists("k"), true);
  });
});

describe("generated modules", { skip }, () => {
  it("compile in strict mode and work at runtime", () => {
    using _ = apartment();
    rmSync(generated, { recursive: true, force: true });
    for (const source of ["scrrun.dll", "stdole2.tlb", "msxml6.dll", "C:\\Windows\\System32\\wbem\\wbemdisp.tlb", "shell32.dll"]) {
      using library = loadTypeLib(source);
      writeModule(readTypeLib(library), generated);
    }
    const usage = join(root, "tests", "typelib", "usage.ts");
    const program = ts.createProgram([usage], {
      target: ts.ScriptTarget.ES2022,
      module: ts.ModuleKind.NodeNext,
      moduleResolution: ts.ModuleResolutionKind.NodeNext,
      strict: true,
      noEmit: true,
      skipLibCheck: true,
      allowImportingTsExtensions: true,
      lib: ["lib.es2022.d.ts", "lib.esnext.disposable.d.ts"],
      types: ["node"],
    });
    const errors = ts.getPreEmitDiagnostics(program);
    assert.equal(errors.length, 0, ts.formatDiagnostics(errors, { getCanonicalFileName: (name) => name, getCurrentDirectory: () => root, getNewLine: () => "\n" }));
    const output = execFileSync(process.execPath, [usage], { cwd: root, encoding: "utf8", timeout: 60_000 });
    const results = JSON.parse(output.trim().split("\n").at(-1)!);
    assert.deepEqual(results.dictionary, [1, true, 1, ["Key"]]);
    assert.deepEqual(results.fso, [true, "string", true]);
    assert.deepEqual(results.msxml, [true, "root", 2, "a", true]);
    assert.deepEqual(results.wmi, [true, ["node.exe"], "ISWbemSinkEvents"]);
    assert.deepEqual(results.shell, [true]);
    assert.deepEqual(results.stdole, ["object", "string"]);
  });
});

describe("CLI", { skip }, () => {
  const cli = join(root, "dist", "cli.js");

  it("prints library information", () => {
    const output = execFileSync(process.execPath, [cli, "info", "Scripting.Dictionary"], { encoding: "utf8" });
    assert.match(output, /^Scripting 1\.0 \{420B2830-E718-11CF-893D-00A0C9054228\}/);
    assert.match(output, /coclass\s+Dictionary \{EE09B103-97E0-11CF-978F-00A02463E06F\}/);
  });

  it("generates TypeScript and JavaScript modules", () => {
    const directory = join(temporary, "cli");
    execFileSync(process.execPath, [cli, "generate", "scrrun.dll", "--out", directory], { encoding: "utf8" });
    assert.match(readFileSync(join(directory, "Scripting.ts"), "utf8"), /export const Dictionary = module\$\["Dictionary"\] as CoClass\$<typeof IDictionary>;/);
    execFileSync(process.execPath, [cli, "generate", "scrrun.dll", "--js", "--out", directory], { encoding: "utf8" });
    assert.ok(existsSync(join(directory, "Scripting.js")) && existsSync(join(directory, "Scripting.d.ts")));
    assert.throws(() => execFileSync(process.execPath, [cli, "generate"], { stdio: "pipe" }));
    void GUID;
  });
});
