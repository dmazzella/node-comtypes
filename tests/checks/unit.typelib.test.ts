import assert from "node:assert/strict";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { describe, it } from "node:test";
import { fileURLToPath } from "node:url";
import * as ctypes from "node-ctypes";
import ts from "typescript";

import { BSTR, COMInterface, CURRENCY, DATE, DISPMETHOD, DISPPROPERTY, GUID, IDispatch, IUnknown, VARIANT, VT_BSTR, buildModule, dispid, generateModule, identifier, lazy, moduleFileName } from "@node-ctypes/comtypes";
import type { CoClass, InterfaceDef, TypeLibModel } from "@node-ctypes/comtypes";

// TYPEKIND / VARENUM / PARAMFLAG values used by the synthetic model.
const TKIND = { ENUM: 0, RECORD: 1, INTERFACE: 3, DISPATCH: 4, COCLASS: 5, ALIAS: 6 };
const VT = { I2: 2, I4: 3, R8: 5, CY: 6, DATE: 7, BSTR: 8, DISPATCH: 9, BOOL: 11, VARIANT: 12, UNKNOWN: 13, I8: 20, VOID: 24, HRESULT: 25, PTR: 26, SAFEARRAY: 27, USERDEFINED: 29 };
const IN = 1;
const OUT = 2;
const LCID = 4;
const RETVAL = 8;
const OPT = 16;
const HASDEFAULT = 32;

const vt = (value: number) => ({ kind: "vt", vt: value }) as const;
const ptr = (to: unknown) => ({ kind: "ptr", to }) as never;
const user = (name: string, typekind: number, guid: string | null = null) => ({ kind: "user", name, typekind, guid, library: null }) as const;
const external = (name: string, guid: string) => ({ kind: "user", name, typekind: TKIND.INTERFACE, guid, library: "stdole" }) as const;
const func = (name: string, slot: number, params: unknown[], invkind = "method", memid = slot) => ({ name, memid, invkind, dispatch: slot < 0, params, returns: vt(VT.HRESULT), slot, hidden: false, restricted: false, doc: null });

const IID_ITHING = "{A0000000-0000-0000-0000-000000000001}";
const IID_IOTHER = "{A0000000-0000-0000-0000-000000000002}";
const IID_EVENTS = "{A0000000-0000-0000-0000-000000000003}";

export const model = {
  name: "Synthetic",
  doc: "Synthetic test library",
  libid: "{A0000000-0000-0000-0000-0000000000FF}",
  major: 1,
  minor: 2,
  lcid: 0,
  syskind: 3,
  types: [
    {
      kind: "enum",
      name: "__MIDL_Color",
      doc: null,
      flags: 0,
      members: [
        { name: "Red", value: 1, doc: null },
        { name: "Blue", value: 4, doc: null },
      ],
    },
    { kind: "alias", name: "Color", doc: null, flags: 0, type: user("__MIDL_Color", TKIND.ENUM) },
    {
      kind: "record",
      name: "Point",
      doc: "A point",
      flags: 0,
      guid: null,
      size: 16,
      alignment: 8,
      fields: [
        { name: "x", type: vt(VT.I4), offset: 0, doc: null },
        { name: "y", type: vt(VT.R8), offset: 8, doc: null },
      ],
    },
    {
      kind: "interface",
      name: "IThing",
      doc: "A thing",
      flags: 0x1040,
      iid: IID_ITHING,
      base: external("IDispatch", "{00020400-0000-0000-C000-000000000046}"),
      dual: true,
      methods: [
        func("Name", 7, [{ name: "value", type: ptr(vt(VT.BSTR)), flags: OUT | RETVAL }], "propget", 1),
        func("Name", 8, [{ name: "value", type: vt(VT.BSTR), flags: IN }], "propput", 1),
        func("Count", 9, [{ name: "value", type: ptr(vt(VT.I4)), flags: OUT | RETVAL }], "propget", 2),
        func(
          "Item",
          10,
          [
            { name: "index", type: vt(VT.VARIANT), flags: IN },
            { name: "item", type: ptr(ptr(user("IThing", TKIND.DISPATCH, IID_ITHING))), flags: OUT | RETVAL },
          ],
          "propget",
          0,
        ),
        func(
          "Find",
          11,
          [
            { name: "what", type: vt(VT.BSTR), flags: IN },
            { name: "start", type: vt(VT.I4), flags: IN | OPT | HASDEFAULT, default: 5 },
            { name: "lcid", type: vt(VT.I4), flags: LCID },
            { name: "other", type: ptr(ptr(user("IOther", TKIND.INTERFACE, IID_IOTHER))), flags: OUT },
            { name: "matches", type: ptr({ kind: "safearray", of: vt(VT.BSTR) }), flags: OUT | RETVAL },
          ],
          "method",
          3,
        ),
        func(
          "Stamp",
          12,
          [
            { name: "when", type: vt(VT.DATE), flags: IN },
            { name: "price", type: ptr(vt(VT.CY)), flags: IN | OUT },
            { name: "point", type: ptr(user("Point", TKIND.RECORD)), flags: OUT },
          ],
          "method",
          4,
        ),
        func("Paint", 13, [{ name: "color", type: user("Color", TKIND.ALIAS), flags: IN }], "method", 5),
      ],
    },
    {
      kind: "interface",
      name: "IOther",
      doc: null,
      flags: 0,
      iid: IID_IOTHER,
      base: external("IUnknown", "{00000000-0000-0000-C000-000000000046}"),
      dual: false,
      methods: [
        func("Back", 3, [{ name: "thing", type: ptr(ptr(user("IThing", TKIND.DISPATCH, IID_ITHING))), flags: OUT | RETVAL }]),
        func("Parent", 4, [{ name: "parent", type: ptr(ptr(external("IDispatch", "{00020400-0000-0000-C000-000000000046}"))), flags: OUT | RETVAL }]),
      ],
    },
    {
      kind: "dispinterface",
      name: "ThingEvents",
      doc: null,
      flags: 0x1000,
      iid: IID_EVENTS,
      methods: [{ ...func("Changed", -1, [{ name: "count", type: vt(VT.I4), flags: IN }], "method", 1), returns: vt(VT.VOID) }],
      properties: [
        { name: "Enabled", memid: 2, type: vt(VT.BOOL), readonly: false, doc: null },
        { name: "Version", memid: 3, type: vt(VT.BSTR), readonly: true, doc: null },
      ],
    },
    {
      kind: "coclass",
      name: "Thing",
      doc: "The thing",
      flags: 2,
      clsid: "{A0000000-0000-0000-0000-000000000010}",
      canCreate: true,
      interfaces: [
        { type: user("IThing", TKIND.DISPATCH, IID_ITHING), default: true, source: false, restricted: false },
        { type: user("ThingEvents", TKIND.DISPATCH, IID_EVENTS), default: true, source: true, restricted: false },
      ],
    },
    {
      kind: "module",
      name: "Constants",
      doc: null,
      flags: 0,
      constants: [
        { name: "MAX_THINGS", value: 10, doc: null },
        { name: "BIG", value: { $bigint: "9007199254740993" }, doc: null },
      ],
    },
  ],
} as unknown as TypeLibModel;

describe("DISPMETHOD and DISPPROPERTY", () => {
  it("describe dispatch members without vtable slots", () => {
    const method = DISPMETHOD([dispid(5)], BSTR, "Describe", [["in"], VARIANT, "value"], [["in", "optional"], ctypes.c_int32, "depth", 1]);
    assert.equal(method.name, "Describe");
    assert.equal(method.dispid, 5);
    assert.equal(method.requiredArguments, 1);
    assert.equal(method.inputArguments, 2);
    assert.deepEqual(method.argtypes, []);
    assert.equal(typeof method.dispatch, "function");
    const getter = DISPMETHOD(["propget", dispid(6)], ctypes.c_int32, "Size");
    assert.equal(getter.name, "_get_Size");
    const property = DISPPROPERTY([dispid(7), "readonly"], DATE, "Created");
    assert.equal(property.writable, false);
    assert.equal(DISPPROPERTY([dispid(8)], CURRENCY, "Price").writable, true);
    assert.throws(() => DISPMETHOD([], null, "NoDispid"), /dispid/);
  });

  it("install methods and accessors, not vtable slots", () => {
    const IEvents = COMInterface({
      name: "IEvents",
      iid: "{B0000000-0000-0000-0000-000000000001}",
      extends: IDispatch,
      dispMethods: [DISPMETHOD([dispid(1)], null, "Fire", [["in"], BSTR, "what"]), DISPPROPERTY([dispid(2)], BSTR, "Label"), DISPPROPERTY([dispid(3), "readonly"], ctypes.c_int32, "Count")],
    });
    assert.equal(IEvents.allMethods.length, IDispatch.allMethods.length, "no vtable slots beyond IDispatch");
    assert.deepEqual(
      IEvents.dispMethods.map((member) => member.name),
      ["Fire", "_get_Label", "_get_Count"],
    );
    const prototype = IEvents.Pointer.prototype;
    assert.equal(typeof Object.getOwnPropertyDescriptor(prototype, "Fire")?.value, "function");
    assert.equal(typeof Object.getOwnPropertyDescriptor(prototype, "_set_Label")?.value, "function");
    const label = Object.getOwnPropertyDescriptor(prototype, "Label")!;
    assert.ok(label.get && label.set);
    const count = Object.getOwnPropertyDescriptor(prototype, "Count")!;
    assert.ok(count.get && !count.set);
    assert.throws(() => COMInterface({ name: "IBad", iid: "{B0000000-0000-0000-0000-000000000002}", extends: IDispatch, methods: [DISPMETHOD([dispid(1)], null, "X")] as never }), /dispMethods/);
  });

  it("lazy references resolve when called", () => {
    let resolved = 0;
    const reference = lazy(() => {
      resolved++;
      return IUnknown;
    });
    assert.equal(resolved, 0);
    assert.equal(reference.resolve(), IUnknown);
    assert.throws(() => lazy(42 as never), TypeError);
  });
});

describe("buildModule", () => {
  const module = buildModule(model);

  it("builds enums, aliases, records and constants", () => {
    assert.deepEqual({ ...(module.__MIDL_Color as object) }, { Red: 1, Blue: 4 });
    assert.equal(module.Color, module.__MIDL_Color, "aliases name the same definition");
    const Point = module.Point as new () => ctypes.Structure;
    assert.equal(ctypes.sizeof(Point), 16);
    const point = new Point();
    point.x = 3;
    point.y = 1.5;
    assert.deepEqual([point.x, point.y], [3, 1.5]);
    assert.equal(module.MAX_THINGS, 10);
    assert.equal(module.BIG, 9007199254740993n);
    assert.equal(module.library.name, "Synthetic");
    assert.ok((module.library.guid as GUID).equals(model.libid));
  });

  it("builds dual interfaces with COMMETHOD specs in slot order", () => {
    const IThing = module.IThing as InterfaceDef<any>;
    assert.equal(IThing.base, IDispatch, "built-in base interfaces are reused");
    assert.deepEqual(
      IThing.methods.map((spec) => spec.name),
      ["_get_Name", "_set_Name", "_get_Count", "_get_Item", "Find", "Stamp", "Paint"],
    );
    assert.equal(IThing.methodIndex.get("Find"), 11);
    const find = IThing.methods[4] as unknown as { params: { name: string; optional: boolean; lcid: boolean; output: boolean; defaultValue: unknown }[]; requiredArguments: number; dispid: number };
    assert.equal(find.dispid, 3);
    assert.equal(find.requiredArguments, 1);
    assert.deepEqual(
      find.params.map((param) => [param.name, param.optional, param.lcid, param.output]),
      [
        ["what", false, false, false],
        ["start", true, false, false],
        ["lcid", false, true, false],
        ["other", false, false, true],
        ["matches", false, false, true],
      ],
    );
    assert.equal(find.params[1].defaultValue, 5);
    const properties = IThing.Pointer.prototype;
    assert.ok(Object.getOwnPropertyDescriptor(properties, "Name")?.set);
    assert.ok(!Object.getOwnPropertyDescriptor(properties, "Count")?.set);
    assert.equal(typeof Object.getOwnPropertyDescriptor(properties, "Item")?.value, "function");
    const IOther = module.IOther as InterfaceDef<any>;
    assert.equal(IOther.base, IUnknown);
  });

  it("builds dispinterfaces and coclasses", () => {
    const events = module.ThingEvents as InterfaceDef<any>;
    assert.deepEqual(
      events.dispMethods.map((spec) => [spec.name, spec.dispid]),
      [
        ["_get_Enabled", 2],
        ["_get_Version", 3],
        ["Changed", 1],
      ],
    );
    const thing = module.Thing as CoClass;
    assert.equal(thing.defaultInterface, module.IThing);
    assert.equal(thing.sourceInterface, events);
    assert.ok(thing.clsid.equals("{A0000000-0000-0000-0000-000000000010}"));
  });
});

describe("generateModule", () => {
  const source = generateModule(model);

  it("emits typed declarations", () => {
    assert.equal(moduleFileName(model), "Synthetic.ts");
    assert.match(source, /export const __MIDL_Color = module\$\["__MIDL_Color"\] as \{ readonly Red: 1; readonly Blue: 4 \};/);
    assert.match(source, /export const Color = __MIDL_Color;\nexport type Color = __MIDL_Color;/);
    assert.match(source, /export interface IThingMethods extends Omit<InterfaceMethods\$<typeof IDispatch\$>/);
    assert.match(source, /  Name: string;/);
    assert.match(source, /  readonly Count: number;/);
    assert.match(source, /  Item\(index: unknown\): InterfacePointer\$<typeof IThing> \| null;/);
    assert.match(source, /  Find\(what: string \| null, start\?: number\): \[InterfacePointer\$<typeof IOther> \| null, unknown\[\] \| null\];/);
    assert.match(source, /  Stamp\(when: Date \| number, price: number \| string \| bigint\): \[number, Point\];/);
    assert.match(source, /  readonly Version: string;/);
    // Plain IDispatch outputs become their best interface or a dynamic object at run time.
    assert.match(source, /  Parent\(\): AutomationObject\$ \| null;/);
    assert.match(source, /export const Thing = module\$\["Thing"\] as CoClass\$<typeof IThing>;/);
    assert.match(source, /export const BIG = module\$\["BIG"\] as bigint;/);
    assert.equal(identifier("delete"), "delete_");
    assert.equal(identifier("a-b"), "a_b");
  });

  it("compiles in strict mode with typed usage", () => {
    const directory = mkdtempSync(join(fileURLToPath(new URL("../..", import.meta.url)), ".generated-"));
    try {
      writeFileSync(join(directory, "Synthetic.ts"), source);
      writeFileSync(
        join(directory, "usage.ts"),
        [
          `import * as Synthetic from "./Synthetic.ts";`,
          `import type { InterfacePointer } from "@node-ctypes/comtypes";`,
          `declare const thing: InterfacePointer<typeof Synthetic.IThing>;`,
          `const name: string = thing.Name;`,
          `thing.Name = "x";`,
          `const count: number = thing.Count;`,
          `const [other, matches] = thing.Find("x");`,
          `const back = other?.Back();`,
          `const again: string | undefined = back?.Name;`,
          `thing.Paint(Synthetic.Color.Blue);`,
          `const point: Synthetic.Point = thing.Stamp(new Date(), "1.5")[1];`,
          `declare const events: InterfacePointer<typeof Synthetic.ThingEvents>;`,
          `events.Enabled = true;`,
          `const version: string = events.Version;`,
          `events.Changed(3);`,
          `// @ts-expect-error Count is read-only.`,
          `thing.Count = 1;`,
          `// @ts-expect-error Version is read-only.`,
          `events.Version = "2";`,
          `// @ts-expect-error Find requires what.`,
          `thing.Find();`,
          `void [name, count, matches, again, point, version];`,
        ].join("\n"),
      );
      const program = ts.createProgram([join(directory, "usage.ts")], {
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
      assert.equal(errors.length, 0, ts.formatDiagnostics(errors, { getCanonicalFileName: (name) => name, getCurrentDirectory: () => directory, getNewLine: () => "\n" }));
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  it("keeps VT_BSTR safe arrays typed", () => {
    assert.ok(source.includes("unknown[] | null"));
    void VT_BSTR;
  });
});
