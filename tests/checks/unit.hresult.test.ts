import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { COMError, E_FAIL, E_INVALIDARG, E_NOINTERFACE, FAILED, HRESULT_CODE, HRESULT_FACILITY, HRESULT_FROM_WIN32, HRESULT_SEVERITY, MAKE_HRESULT, S_FALSE, S_OK, SUCCEEDED, checkHRESULT, formatHRESULT, toHRESULT } from "@node-ctypes/comtypes";

describe("HRESULT", () => {
  it("constants are signed 32-bit values", () => {
    assert.equal(E_FAIL, -2147467259);
    assert.equal(E_NOINTERFACE, -2147467262);
    assert.equal(E_INVALIDARG >>> 0, 0x80070057);
  });

  it("SUCCEEDED / FAILED accept signed, unsigned and bigint values", () => {
    for (const hr of [S_OK, S_FALSE, 0x7fffffff]) assert.ok(SUCCEEDED(hr) && !FAILED(hr));
    for (const hr of [E_FAIL, E_FAIL >>> 0, BigInt(E_FAIL >>> 0), BigInt(E_FAIL)]) assert.ok(FAILED(hr) && !SUCCEEDED(hr));
    assert.equal(toHRESULT(0x80004005), E_FAIL);
    assert.equal(toHRESULT(0x80004005n), E_FAIL);
  });

  it("builds and decomposes HRESULTs", () => {
    assert.equal(HRESULT_FROM_WIN32(87), E_INVALIDARG);
    assert.equal(HRESULT_FROM_WIN32(0), S_OK);
    assert.equal(HRESULT_FROM_WIN32(E_FAIL), E_FAIL);
    assert.equal(MAKE_HRESULT(1, 7, 87), E_INVALIDARG);
    assert.equal(HRESULT_CODE(E_INVALIDARG), 87);
    assert.equal(HRESULT_FACILITY(E_INVALIDARG), 7);
    assert.equal(HRESULT_SEVERITY(E_INVALIDARG), 1);
    assert.equal(HRESULT_SEVERITY(S_FALSE), 0);
    assert.equal(formatHRESULT(E_FAIL), "0x80004005");
    assert.equal(formatHRESULT(1), "0x00000001");
  });
});

describe("COMError", () => {
  it("carries hresult, text, details and method", () => {
    const error = new COMError(0x80004005, { text: "Unspecified error", method: "IFoo::Bar", details: { description: "boom", source: "Foo", helpFile: null, helpContext: 0, guid: null } });
    assert.ok(error instanceof Error);
    assert.equal(error.name, "COMError");
    assert.equal(error.hresult, E_FAIL);
    assert.equal(error.winerror, 0x80004005);
    assert.equal(error.method, "IFoo::Bar");
    assert.equal(error.details?.description, "boom");
    assert.equal(error.message, "IFoo::Bar: Unspecified error (0x80004005) - boom");
  });

  it("has a text for any HRESULT", () => {
    const error = new COMError(E_FAIL);
    assert.ok(error.text.length > 0);
    assert.equal(error.details, null);
    assert.equal(error.method, null);
  });

  it("checkHRESULT passes success codes through", () => {
    assert.equal(checkHRESULT(S_OK), S_OK);
    assert.equal(checkHRESULT(S_FALSE), S_FALSE);
    assert.throws(
      () => checkHRESULT(E_NOINTERFACE, "Test"),
      (error: unknown) => error instanceof COMError && error.hresult === E_NOINTERFACE && error.method === "Test",
    );
  });
});
