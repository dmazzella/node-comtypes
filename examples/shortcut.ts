// Create a Windows shortcut with IShellLinkW + IPersistFile and read it back.
//
//   npm run build && node examples/shortcut.ts

import { tmpdir } from "node:os";
import { join } from "node:path";
import { rmSync } from "node:fs";
import { create_unicode_buffer } from "node-ctypes";
import { CLSID_ShellLink, CoCreateInstance, IPersistFile, IShellLinkW, STGM_READ, apartment } from "@node-ctypes/comtypes";

const path = join(tmpdir(), "node-comtypes-notepad.lnk");

using _ = apartment();

{
  using link = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
  link.SetPath("C:\\Windows\\notepad.exe");
  link.SetDescription("Notepad — created by node-comtypes");
  link.SetWorkingDirectory("C:\\Windows");
  using file = link.QueryInterface(IPersistFile);
  file.Save(path, 1);
  console.log(`Saved ${path}`);
}

{
  using link = CoCreateInstance(CLSID_ShellLink, IShellLinkW);
  using file = link.QueryInterface(IPersistFile);
  file.Load(path, STGM_READ);
  const text = create_unicode_buffer(1024);
  link.GetPath(text, 1024, null, 0);
  console.log(`Target:      ${text.value}`);
  link.GetDescription(text, 1024);
  console.log(`Description: ${text.value}`);
}

rmSync(path, { force: true });
