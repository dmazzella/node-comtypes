import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { inspect } from "node:util";
import * as ctypes from "node-ctypes";

import { GUID, GUID_NULL, toGUID, IID_IUnknown } from "@node-ctypes/comtypes";

describe("GUID", () => {
  it("is 16 bytes with the Windows layout", () => {
    assert.equal(ctypes.sizeof(GUID), 16);
    const guid = new GUID("{00021401-0000-0000-C000-000000000046}");
    assert.equal(guid.Data1, 0x00021401);
    assert.equal(guid.Data2, 0);
    assert.equal(guid.Data3, 0);
    assert.deepEqual([...guid._buffer], [0x01, 0x14, 0x02, 0x00, 0, 0, 0, 0, 0xc0, 0, 0, 0, 0, 0, 0, 0x46]);
  });

  it("parses with or without braces and formats uppercase", () => {
    const text = "{1cf2b120-547d-101b-8e65-08002b2bc5bb}";
    assert.equal(new GUID(text).toString(), text.toUpperCase());
    assert.equal(new GUID(text.slice(1, -1)).toString(), text.toUpperCase());
    assert.equal(String(new GUID(` ${text} `)), text.toUpperCase());
    assert.equal(inspect(new GUID(text)), `GUID("${text.toUpperCase()}")`);
  });

  it("round-trips high values without sign issues", () => {
    const text = "{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}";
    const guid = new GUID(text);
    assert.equal(guid.Data1, 0xffffffff);
    assert.equal(guid.toString(), text);
  });

  it("rejects malformed strings and buffers", () => {
    for (const bad of ["", "{}", "00000000-0000-0000-C000-00000000004", "{00000000-0000-0000-C000-00000000004G}", "Scripting.Dictionary", "{00000000000000000000000000000000}"]) {
      assert.throws(() => new GUID(bad), SyntaxError, bad);
    }
    assert.throws(() => new GUID(Buffer.alloc(15)), TypeError);
    assert.equal(GUID.isGUIDString("{00000000-0000-0000-C000-000000000046}"), true);
    assert.equal(GUID.isGUIDString("Scripting.Dictionary"), false);
  });

  it("copies from bytes and other GUIDs", () => {
    const source = new GUID("{0000010B-0000-0000-C000-000000000046}");
    const copy = new GUID(source);
    assert.ok(copy.equals(source));
    copy.Data1 = 1;
    assert.ok(!copy.equals(source));
    assert.ok(new GUID(Buffer.from(source._buffer)).equals(source));
  });

  it("compares by value with GUIDs and strings", () => {
    assert.ok(IID_IUnknown.equals("{00000000-0000-0000-c000-000000000046}"));
    assert.ok(!IID_IUnknown.equals(null));
    assert.ok(GUID_NULL.isNull());
    assert.ok(new GUID().equals(GUID_NULL));
    assert.ok(!IID_IUnknown.isNull());
  });

  it("toGUID accepts GUIDs and strings only", () => {
    assert.equal(toGUID(IID_IUnknown), IID_IUnknown);
    assert.ok(toGUID("{00000000-0000-0000-C000-000000000046}").equals(IID_IUnknown));
    assert.throws(() => toGUID(42 as never), TypeError);
  });
});
