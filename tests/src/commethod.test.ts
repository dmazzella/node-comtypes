import assert from "node:assert/strict";
import { describe, it } from "node:test";
import * as ctypes from "node-ctypes";

import {
  BSTR,
  COMError,
  COMInterface,
  COMMETHOD,
  COMPointer,
  CreateObject,
  HRESULT,
  IDispatch,
  ITypeInfo,
  IUnknown,
  INVOKE_FUNC,
  INVOKE_PROPERTYGET,
  INVOKE_PROPERTYPUT,
  SELF,
  STDMETHOD,
  VARIANT,
  apartment,
  dispatchPointer,
  dispatchTypeInfo,
  documentation,
  getLivePointerCount,
  named,
  sync,
  withFuncDesc,
  withTypeAttr,
} from "@node-ctypes/comtypes";
import type { DispatchObject, InterfacePointer } from "@node-ctypes/comtypes";
import { COMObject } from "@node-ctypes/comtypes/server";

const skip = process.platform !== "win32" && "COM requires Windows";
const { c_int32, c_int16 } = ctypes;
const VARIANT_BOOL = ctypes.wintypes.VARIANT_BOOL;

// msxml6.idl: IXMLDOMNode (dual). Node-valued results are IXMLDOMNode (SELF);
// other interfaces (node lists, maps, documents) are declared as IDispatch.
const IXMLDOMNode = COMInterface({
  name: "IXMLDOMNode",
  iid: "{2933BF80-7B36-11D2-B20E-00C04F983E60}",
  extends: IDispatch,
  methods: [
    COMMETHOD(["propget"], HRESULT, "nodeName", [["out", "retval"], BSTR, "name"]),
    COMMETHOD(["propget"], HRESULT, "nodeValue", [["out", "retval"], VARIANT, "value"]),
    COMMETHOD(["propput"], HRESULT, "nodeValue", [["in"], VARIANT, "value"]),
    COMMETHOD(["propget"], HRESULT, "nodeType", [["out", "retval"], c_int32, "type"]),
    COMMETHOD(["propget"], HRESULT, "parentNode", [["out", "retval"], SELF, "parent"]),
    COMMETHOD(["propget"], HRESULT, "childNodes", [["out", "retval"], IDispatch, "childList"]),
    COMMETHOD(["propget"], HRESULT, "firstChild", [["out", "retval"], SELF, "firstChild"]),
    COMMETHOD(["propget"], HRESULT, "lastChild", [["out", "retval"], SELF, "lastChild"]),
    COMMETHOD(["propget"], HRESULT, "previousSibling", [["out", "retval"], SELF, "previousSibling"]),
    COMMETHOD(["propget"], HRESULT, "nextSibling", [["out", "retval"], SELF, "nextSibling"]),
    COMMETHOD(["propget"], HRESULT, "attributes", [["out", "retval"], IDispatch, "attributeMap"]),
    COMMETHOD([], HRESULT, "insertBefore", [["in"], SELF, "newChild"], [["in", "optional"], VARIANT, "refChild"], [["out", "retval"], SELF, "outNewChild"]),
    COMMETHOD([], HRESULT, "replaceChild", [["in"], SELF, "newChild"], [["in"], SELF, "oldChild"], [["out", "retval"], SELF, "outOldChild"]),
    COMMETHOD([], HRESULT, "removeChild", [["in"], SELF, "childNode"], [["out", "retval"], SELF, "oldChild"]),
    COMMETHOD([], HRESULT, "appendChild", [["in"], SELF, "newChild"], [["out", "retval"], SELF, "outNewChild"]),
    COMMETHOD([], HRESULT, "hasChildNodes", [["out", "retval"], VARIANT_BOOL, "hasChild"]),
    COMMETHOD(["propget"], HRESULT, "ownerDocument", [["out", "retval"], IDispatch, "XMLDOMDocument"]),
    COMMETHOD([], HRESULT, "cloneNode", [["in"], VARIANT_BOOL, "deep"], [["out", "retval"], SELF, "cloneRoot"]),
    COMMETHOD(["propget"], HRESULT, "nodeTypeString", [["out", "retval"], BSTR, "nodeType"]),
    COMMETHOD(["propget"], HRESULT, "text", [["out", "retval"], BSTR, "text"]),
    COMMETHOD(["propput"], HRESULT, "text", [["in"], BSTR, "text"]),
    COMMETHOD(["propget"], HRESULT, "specified", [["out", "retval"], VARIANT_BOOL, "isSpecified"]),
    COMMETHOD(["propget"], HRESULT, "definition", [["out", "retval"], SELF, "definitionNode"]),
    COMMETHOD(["propget"], HRESULT, "nodeTypedValue", [["out", "retval"], VARIANT, "typedValue"]),
    COMMETHOD(["propput"], HRESULT, "nodeTypedValue", [["in"], VARIANT, "typedValue"]),
    COMMETHOD(["propget"], HRESULT, "dataType", [["out", "retval"], VARIANT, "dataTypeName"]),
    COMMETHOD(["propput"], HRESULT, "dataType", [["in"], BSTR, "dataTypeName"]),
    COMMETHOD(["propget"], HRESULT, "xml", [["out", "retval"], BSTR, "xmlString"]),
    COMMETHOD([], HRESULT, "transformNode", [["in"], SELF, "stylesheet"], [["out", "retval"], BSTR, "xmlString"]),
    COMMETHOD([], HRESULT, "selectNodes", [["in"], BSTR, "queryString"], [["out", "retval"], IDispatch, "resultList"]),
    COMMETHOD([], HRESULT, "selectSingleNode", [["in"], BSTR, "queryString"], [["out", "retval"], SELF, "resultNode"]),
  ],
});

type Node = InterfacePointer<typeof IXMLDOMNode>;
const NODE_ELEMENT = 1;
const NODE_TEXT = 3;

async function documentNode(xml: string): Promise<{ document: DispatchObject; root: Node }> {
  const document = await CreateObject("MSXML2.DOMDocument.6.0", { dynamic: true });
  document.async = false;
  assert.equal(await document.loadXML(xml), true);
  // Setup: the document element, read through a synchronous view.
  using node = dispatchPointer(document).QueryInterface(IXMLDOMNode);
  const root = sync(node).selectSingleNode("/*")!;
  return { document, root };
}

/** Vtable slots by "name/invkind", walking a dual interface and its bases. */
function vtableLayout(typeInfo: InterfacePointer<typeof ITypeInfo>): Map<string, number> {
  const layout = new Map<string, number>();
  const reference = new ctypes.c_uint32();
  const out = new ctypes.c_void_p();
  sync(typeInfo).GetRefTypeOfImplType(-1 >>> 0, ctypes.byref(reference));
  sync(typeInfo).GetRefTypeInfo(reference.value, ctypes.byref(out));
  let current: InterfacePointer<typeof ITypeInfo> | null = COMPointer.fromAddress(out.value!, ITypeInfo);
  while (current) {
    using info = current;
    current = null;
    const { functions, implemented } = withTypeAttr(info, (attr) => ({ functions: attr.cFuncs, implemented: attr.cImplTypes }));
    for (let index = 0; index < functions; index++) {
      withFuncDesc(info, index, (desc) => layout.set(`${documentation(info, desc.memid).name}/${desc.invkind}`, desc.oVft / ctypes.sizeof(ctypes.c_void_p)));
    }
    if (implemented) {
      const base = new ctypes.c_uint32();
      const next = new ctypes.c_void_p();
      sync(info).GetRefTypeOfImplType(0, ctypes.byref(base));
      sync(info).GetRefTypeInfo(base.value, ctypes.byref(next));
      current = COMPointer.fromAddress(next.value!, ITypeInfo);
    }
  }
  return layout;
}

describe("COMMETHOD definitions", () => {
  it("names accessors as comtypes and computes native argument types", () => {
    const get = COMMETHOD(["propget"], HRESULT, "Name", [["out", "retval"], BSTR, "p"]);
    const put = COMMETHOD(["propput"], HRESULT, "Name", [["in"], BSTR, "v"]);
    const putref = COMMETHOD(["propputref"], HRESULT, "Name", [["in"], IUnknown, "v"]);
    assert.deepEqual([get.name, put.name, putref.name], ["_get_Name", "_set_Name", "_setref_Name"]);
    assert.deepEqual([get.kind, put.kind, putref.kind], ["propget", "propput", "propputref"]);
    assert.equal(get.property, "Name");
    assert.deepEqual(get.argtypes, [ctypes.c_void_p]);
    const variant = COMMETHOD([], HRESULT, "M", [["in"], VARIANT, "v"], [["in", "optional"], c_int16, "n", 5], [["lcid"], ctypes.c_uint32, "lcid"]);
    assert.deepEqual(variant.argtypes, [VARIANT, c_int16, ctypes.c_uint32]);
    assert.equal(variant.requiredArguments, 1);
    assert.equal(variant.inputArguments, 2);
    assert.equal(variant.params[1].defaultValue, 5);
  });

  it("validates flags and output types", () => {
    assert.throws(() => COMMETHOD(["propget", "propput"], HRESULT, "X"), TypeError);
    assert.throws(() => COMMETHOD([], HRESULT, "X", [["inout" as "in"], c_int32, "x"]), TypeError);
    assert.throws(() => COMMETHOD([], HRESULT, "X", [["out"], ctypes.c_void_p, "x"]) && COMMETHOD([], HRESULT, "Y", [["out"], ctypes.array(c_int32, 2), "x"]), TypeError);
    assert.throws(() => COMMETHOD([], HRESULT, ""), TypeError);
  });

  it("installs properties and parameterized properties", () => {
    const IThing = COMInterface({
      name: "IThing",
      iid: "{33333333-3333-3333-3333-333333333333}",
      extends: IUnknown,
      methods: [
        COMMETHOD(["propget"], HRESULT, "Name", [["out", "retval"], BSTR, "p"]),
        COMMETHOD(["propput"], HRESULT, "Name", [["in"], BSTR, "v"]),
        COMMETHOD(["propget"], HRESULT, "Item", [["in"], c_int32, "i"], [["out", "retval"], BSTR, "p"]),
        COMMETHOD(["propput"], HRESULT, "Secret", [["in"], BSTR, "v"]),
        STDMETHOD(HRESULT, "Raw"),
      ],
    });
    const prototype = IThing.Pointer.prototype;
    const name = Object.getOwnPropertyDescriptor(prototype, "Name")!;
    assert.equal(typeof name.get, "function");
    assert.equal(typeof name.set, "function");
    assert.equal(typeof Object.getOwnPropertyDescriptor(prototype, "Item")!.value, "function");
    const secret = Object.getOwnPropertyDescriptor(prototype, "Secret")!;
    assert.equal(secret.get, undefined);
    assert.equal(typeof secret.set, "function");
    assert.deepEqual(
      ["_get_Name", "_set_Name", "_get_Item", "_set_Secret", "Raw"].map((method) => IThing.methodIndex.get(method)),
      [3, 4, 5, 6, 7],
    );
  });
});

describe("COMMETHOD calls", { skip }, () => {
  it("declares IXMLDOMNode in the order of its type information", async () => {
    using _ = apartment();
    using document = await CreateObject("MSXML2.DOMDocument.6.0");
    using typeInfo = dispatchTypeInfo(dispatchPointer(document))!;
    const layout = vtableLayout(typeInfo);
    const kinds = { method: INVOKE_FUNC, propget: INVOKE_PROPERTYGET, propput: INVOKE_PROPERTYPUT, propputref: 8 };
    for (const method of IXMLDOMNode.methods) {
      assert.equal(layout.get(`${method.property}/${kinds[method.kind]}`), IXMLDOMNode.methodIndex.get(method.name), method.name);
    }
  });

  it("converts BSTR, VARIANT, VARIANT_BOOL and interface results", async () => {
    using _ = apartment();
    const { document, root } = await documentNode("<root a='1'><child>text</child><child>two</child></root>");
    using __ = document;
    using rootNode = root;
    assert.equal(rootNode.nodeName, "root");
    assert.equal(rootNode._get_nodeName(), "root");
    assert.equal(rootNode.nodeType, NODE_ELEMENT);
    assert.equal(rootNode.nodeTypeString, "element");
    assert.equal(await rootNode.hasChildNodes(), true);
    assert.equal(rootNode.nodeValue, null, "elements have a VT_NULL node value");
    using first = rootNode.firstChild!;
    assert.ok(first instanceof IXMLDOMNode.Pointer);
    assert.equal(first.text, "text");
    using second = first.nextSibling!;
    assert.equal(second.text, "two");
    assert.equal(second.nextSibling, null, "NULL interface results are null");
    using parent = second.parentNode!;
    assert.ok(parent.isSameObject(rootNode));
    using textNode = first.firstChild!;
    assert.equal(textNode.nodeType, NODE_TEXT);
    textNode.nodeValue = "changed";
    assert.equal(first.text, "changed");
    first.text = "set through propput";
    // Setting text replaces the child text node.
    using replaced = first.firstChild!;
    assert.equal(replaced.nodeValue, "set through propput");
  });

  it("passes interfaces in and returns them, by position or by name", async () => {
    using _ = apartment();
    const { document, root } = await documentNode("<root><a/></root>");
    using __ = document;
    using rootNode = root;
    using a = (await rootNode.selectSingleNode("a"))!;
    using clone = (await a.cloneNode(true))!;
    using appended = (await rootNode.appendChild(clone))!;
    assert.ok(appended.isSameObject(clone));
    using fresh = (await a.cloneNode(false))!;
    // Optional VARIANT omitted: VT_ERROR/DISP_E_PARAMNOTFOUND appends at the end.
    using inserted = (await rootNode.insertBefore(fresh))!;
    assert.ok(inserted.isSameObject(fresh));
    using another = (await a.cloneNode(false))!;
    using before = await (rootNode.insertBefore as (...args: unknown[]) => Promise<Node>)(named({ newChild: another, refChild: a }));
    assert.ok(before.isSameObject(another));
    assert.equal(rootNode.xml, "<root><a/><a/><a/><a/></root>");
    using removed = (await rootNode.removeChild(a))!;
    assert.ok(removed.isSameObject(a));
    // Dynamic objects are accepted for interface parameters (QueryInterface to IXMLDOMNode).
    using created = await document.createElement("dynamic");
    using appendedDynamic = (await rootNode.appendChild(created))!;
    assert.equal(appendedDynamic.nodeName, "dynamic");
  });

  it("validates arguments before calling", async () => {
    using _ = apartment();
    const { document, root } = await documentNode("<root/>");
    using __ = document;
    using rootNode = root;
    const untyped = rootNode as unknown as Record<string, (...args: unknown[]) => Promise<unknown>>;
    await assert.rejects(untyped.selectSingleNode(), /missing argument queryString/);
    await assert.rejects(untyped.selectSingleNode("a", "b"), /at most 1 arguments/);
    await assert.rejects(untyped.selectSingleNode(named({ nope: 1 })), /unknown argument/);
    await assert.rejects(untyped.appendChild(42), /Expected a COM object/);
  });

  it("raises COMError and frees outputs on failure", async () => {
    using _ = apartment();
    const { document, root } = await documentNode("<root/>");
    using __ = document;
    using rootNode = root;
    const before = getLivePointerCount();
    await assert.rejects(rootNode.selectSingleNode("///"), (error: unknown) => error instanceof COMError && error.method === "IXMLDOMNode::selectSingleNode" && error.hresult < 0);
    // MSXML reports the reason through IErrorInfo.
    await assert.rejects(rootNode.appendChild(rootNode), (error: unknown) => error instanceof COMError && error.hresult < 0 && typeof error.details?.description === "string");
    assert.equal(getLivePointerCount(), before);
  });

  it("does not leak interface results", async () => {
    using _ = apartment();
    const { document, root } = await documentNode("<root><a>x</a></root>");
    using __ = document;
    using rootNode = root;
    const before = getLivePointerCount();
    for (let i = 0; i < 1000; i++) {
      using child = rootNode.firstChild!;
      assert.equal(child.text, "x");
    }
    assert.equal(getLivePointerCount(), before);
  });
});

describe("COMMETHOD [in, out] and lcid parameters", { skip }, () => {
  it("passes and returns in/out values and fills the lcid", async () => {
    const { allocBSTR, oleaut32, readBSTR } = await import("@node-ctypes/comtypes/low-level");
    const { VARIANT_SIZE, writeVariant, setDefaultLCID } = await import("@node-ctypes/comtypes");
    const ITwice = COMInterface({
      name: "ITwice",
      iid: "{44444444-4444-4444-4444-444444444444}",
      extends: IUnknown,
      methods: [COMMETHOD([], HRESULT, "Twice", [["in", "out"], c_int32, "number"], [["in", "out"], BSTR, "text"], [["lcid"], ctypes.c_uint32, "lcid"], [["out", "retval"], VARIANT, "seenLcid"])],
    });
    // A minimal COM object implemented with node-ctypes callbacks.
    let references = 1;
    const queryInterface = ctypes.callback(() => -2147467262, c_int32, [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p]);
    const addRef = ctypes.callback(() => ++references, ctypes.c_uint32, [ctypes.c_void_p]);
    const release = ctypes.callback(() => --references, ctypes.c_uint32, [ctypes.c_void_p]);
    const twice = ctypes.callback(
      (_self: bigint | null, number: bigint | null, text: bigint | null, lcid: number, result: bigint | null) => {
        ctypes.writeValue(number!, c_int32, ctypes.readValue(number!, c_int32) * 2);
        const previous = ctypes.readValue(text!, ctypes.c_void_p);
        const value = readBSTR(previous) ?? "";
        oleaut32().SysFreeString(previous);
        ctypes.writeValue(text!, ctypes.c_void_p, allocBSTR(value + value));
        const variant = Buffer.alloc(VARIANT_SIZE);
        writeVariant(variant, lcid);
        ctypes.memmove(result!, variant, VARIANT_SIZE);
        return 0;
      },
      c_int32,
      [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint32, ctypes.c_void_p],
    );
    const pointerSize = ctypes.sizeof(ctypes.c_void_p);
    const vtable = Buffer.alloc(4 * pointerSize);
    [queryInterface, addRef, release, twice].forEach((callback, index) => ctypes.writeValue(vtable, ctypes.c_void_p, callback.pointer, index * pointerSize));
    const object = Buffer.alloc(pointerSize);
    ctypes.writeValue(object, ctypes.c_void_p, ctypes.addressof(vtable));
    try {
      // A single-threaded home: this object (plain callbacks, neither agile nor a proxy) is called on it.
      using _ = apartment();
      using pointer = COMPointer.fromAddress(ctypes.addressof(object), ITwice);
      assert.deepEqual(await pointer.Twice(21, "ab"), [42, "abab", 0]);
      setDefaultLCID(0x409);
      try {
        // named() is runtime-only: the static signature is positional.
        assert.deepEqual(await (pointer.Twice as (...args: unknown[]) => Promise<unknown>)(named({ text: "x", number: -1 })), [-2, "xx", 0x409]);
      } finally {
        setDefaultLCID(0);
      }
    } finally {
      assert.equal(references, 0);
      for (const callback of [queryInterface, addRef, release, twice]) callback.release();
      void [vtable, object];
    }
  });
});

describe("COMMETHOD POINTER(T) parameters", { skip }, () => {
  it("pass node-ctypes arrays as raw pointers", async () => {
    using _ = apartment();
    const ISum = COMInterface({
      name: "ISum",
      iid: "{0E8C5A30-0000-4000-8000-000000000001}",
      extends: IUnknown,
      methods: [COMMETHOD([], HRESULT, "Sum", [["in"], ctypes.POINTER(ctypes.c_int32), "values"], [["in"], ctypes.c_uint32, "count"], [["out", "retval"], ctypes.c_int32, "total"])],
    });
    class Sum extends COMObject {
      static _com_interfaces_ = [ISum];
      received = 0n;
      // The server receives the raw address of the array.
      Sum(values: bigint, count: number) {
        this.received = values;
        return Array.from({ length: count }, (_, index) => ctypes.readValue(values + BigInt(index * 4), ctypes.c_int32) as number).reduce((a, b) => a + b, 0);
      }
    }
    const object = new Sum();
    using sum = object.QueryInterface(ISum);
    // A node-ctypes array is passed as its address,
    // not copied into a temporary c_int32.
    const values = ctypes.array(ctypes.c_int32, 3)(1, 2, 39);
    assert.equal(await sum.Sum(values, 3), 42);
    assert.equal(object.received, ctypes.addressof(values));
    // A plain number still goes through a temporary.
    assert.equal(await sum.Sum(7, 1), 7);
  });
});
