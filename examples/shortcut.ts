// Create a Windows shortcut with IShellLinkW + IPersistFile and read it back.
// The shell link is created on the COM executor: its methods run there.
//
//   npm run build && node examples/shortcut.ts

import { tmpdir } from "node:os";
import { join } from "node:path";
import { rmSync } from "node:fs";
import { create_unicode_buffer } from "node-ctypes";
import { CreateObject, IPersistFile, STGM_READ, ShellLink, apartment } from "@node-ctypes/comtypes";

const path = join(tmpdir(), "node-comtypes-notepad.lnk");

using _ = apartment();

{
  using link = await CreateObject(ShellLink); // IShellLinkW
  await link.SetPath("C:\\Windows\\notepad.exe");
  await link.SetDescription("Notepad — created by node-comtypes");
  await link.SetWorkingDirectory("C:\\Windows");
  using file = link.QueryInterface(IPersistFile);
  await file.Save(path, 1);
  console.log(`Saved ${path}`);
}

{
  using link = await CreateObject(ShellLink);
  using file = link.QueryInterface(IPersistFile);
  await file.Load(path, STGM_READ);
  const text = create_unicode_buffer(1024);
  await link.GetPath(text, 1024, null, 0);
  console.log(`Target:      ${text.value}`);
  await link.GetDescription(text, 1024);
  console.log(`Description: ${text.value}`);
}

rmSync(path, { force: true });
