import assert from "node:assert/strict";
import { existsSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { after, describe, it } from "node:test";
import * as ctypes from "node-ctypes";

import { CLSID_ShellLink, COMError, CoCreateInstance, IPersistFile, IShellLinkW, STGM_READ, S_OK, apartment } from "@node-ctypes/comtypes";

const skip = process.platform !== "win32" && "COM requires Windows";
const directory = mkdtempSync(join(tmpdir(), "node-comtypes-"));
after(() => rmSync(directory, { recursive: true, force: true }));

describe("IShellLinkW + IPersistFile", { skip }, () => {
  it("creates and reads back a shortcut", async () => {
    using _ = apartment();
    const path = join(directory, "notepad.lnk");
    {
      using link = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
      assert.equal(await link.SetPath("C:\\Windows\\notepad.exe"), S_OK);
      await link.SetDescription("Notepad — node-comtypes");
      await link.SetWorkingDirectory("C:\\Windows");
      await link.SetArguments("/A readme.txt");
      await link.SetIconLocation("C:\\Windows\\notepad.exe", 1);
      await link.SetShowCmd(3);
      using file = link.QueryInterface(IPersistFile);
      await file.Save(path, 1);
      assert.equal((await file.GetCurFile())?.toLowerCase(), path.toLowerCase());
    }
    assert.ok(existsSync(path));

    using link = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
    using file = link.QueryInterface(IPersistFile);
    await file.Load(path, STGM_READ);
    const text = ctypes.create_unicode_buffer(1024);
    await link.GetPath(text, 1024, null, 0);
    assert.equal(text.value.toLowerCase(), "c:\\windows\\notepad.exe");
    await link.GetDescription(text, 1024);
    assert.equal(text.value, "Notepad — node-comtypes");
    await link.GetWorkingDirectory(text, 1024);
    assert.equal(text.value, "C:\\Windows");
    await link.GetArguments(text, 1024);
    assert.equal(text.value, "/A readme.txt");
    const icon = new ctypes.c_int32();
    await link.GetIconLocation(text, 1024, ctypes.byref(icon));
    assert.equal(icon.value, 1);
    const show = new ctypes.c_int32();
    await link.GetShowCmd(ctypes.byref(show));
    assert.equal(show.value, 3);
  });

  it("raises COMError with the failing method", async () => {
    using _ = apartment();
    using link = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
    using file = link.QueryInterface(IPersistFile);
    await assert.rejects(file.Load(join(directory, "missing.lnk"), STGM_READ), (error: unknown) => error instanceof COMError && error.method === "IPersistFile::Load" && error.hresult < 0 && error.message.startsWith("IPersistFile::Load: "));
  });
});
