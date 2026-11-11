// Active Directory search through ADSI IDirectorySearch, the native interface
// behind ADO and System.DirectoryServices.
//
//   npm run build
//   node examples/active-directory.ts
//   node examples/active-directory.ts --filter "(objectClass=computer)" --attributes cn,operatingSystem,lastLogonTimestamp
//   node examples/active-directory.ts --base "LDAP://OU=Staff,DC=contoso,DC=com" --scope onelevel --tree
//   node examples/active-directory.ts --user CONTOSO\admin --pass secret
//
// What node-comtypes provides:
// - binding without native calls: GetObject("LDAP://...") (ADSI monikers) and,
//   with credentials, GetObject("LDAP:").OpenDSObject(...) (= ADsOpenObject);
// - ADSVALUE (with its 27-member union), ADS_SEARCHPREF_INFO,
//   ADS_SEARCH_COLUMN and the ADS_* enums from the Active DS type library;
// - IDirectorySearch declared with COMMETHOD: outputs are returned, failures throw.

import { parseArgs } from "node:util";
import * as ctypes from "node-ctypes";
import { COINIT_MULTITHREADED, COMInterface, COMMETHOD, GUID, GetModule, GetObject, HRESULT, IUnknown, LPOLESTR, STDMETHOD, apartment, dispatchPointer } from "@node-ctypes/comtypes";
import type { DispatchObject } from "@node-ctypes/comtypes";
import { readMemory } from "@node-ctypes/comtypes/low-level";

function parseOptions() {
  try {
    return parseArgs({
      options: {
        base: { type: "string" }, // default: the domain of this computer (RootDSE defaultNamingContext)
        filter: { type: "string", default: "(&(objectClass=user)(objectCategory=person))" },
        attributes: { type: "string", default: "distinguishedName,cn,displayName,userPrincipalName,givenName,sn,mail,objectSid,objectGUID" }, // "*": all
        scope: { type: "string", default: "subtree" }, // base | onelevel | subtree
        "page-size": { type: "string", default: "1000" },
        user: { type: "string" },
        pass: { type: "string" },
        limit: { type: "string" }, // print at most this many objects (default: all)
        tree: { type: "boolean", default: false }, // group the objects by container (OU/CN) path
        debug: { type: "boolean", default: false },
      },
    }).values;
  } catch (error) {
    // Unknown options, missing values ("--limit -1" is read as an option: use --limit=-1).
    console.error((error as Error).message);
    process.exit(2);
  }
}

const options = parseOptions();

// A multithreaded apartment: directory objects obtained on the COM executor are
// in the same apartment, so IDirectorySearch (a [local] interface, never marshaled)
// is usable here and its calls run on the executor. In a single-threaded
// apartment (e.g. Electron's main thread) use GetObject(path, { home: true }).
using _ = apartment(COINIT_MULTITHREADED);

// The Active DS type library describes the ADSI structures and enums.
const Ads = (await GetModule("activeds.tlb")) as any;
const { ADSTYPEENUM: AdsType, ADS_SEARCHPREF_ENUM: Pref, ADS_SCOPEENUM: Scope, ADS_CHASE_REFERRALS_ENUM: Referrals, ADS_AUTHENTICATION_ENUM: Auth } = Ads;
// MIDL names the anonymous union of ADSVALUE: read its name from the record.
const UNION: string = Ads._adsvalue._fields_[1][0];
const ADSVALUE_SIZE = ctypes.sizeof(Ads._adsvalue);
const SEARCHPREF_SIZE = ctypes.sizeof(Ads.ads_searchpref_info);

const S_ADS_NOMORE_ROWS = 0x5012;
const S_ADS_NOMORE_COLUMNS = 0x5013;

// The type library also defines IDirectorySearch; declaring it here gives static
// types and keeps GetNextColumnName raw: its name is freed with FreeADsMem (not
// CoTaskMemFree) and S_ADS_NOMORE_COLUMNS is a success code.
const IDirectorySearch = COMInterface({
  name: "IDirectorySearch",
  iid: "{109BA8EC-92F0-11D0-A790-00C04FD8D5A8}",
  extends: IUnknown,
  methods: [
    COMMETHOD([], HRESULT, "SetSearchPreference", [["in"], ctypes.c_void_p, "pSearchPrefs"], [["in"], ctypes.c_uint32, "dwNumPrefs"]),
    COMMETHOD([], HRESULT, "ExecuteSearch", [["in"], LPOLESTR, "pszSearchFilter"], [["in"], ctypes.POINTER(ctypes.c_wchar_p), "pAttributeNames"], [["in"], ctypes.c_uint32, "dwNumberAttributes"], [["out"], ctypes.c_void_p, "phSearchResult"]),
    COMMETHOD([], HRESULT, "AbandonSearch", [["in"], ctypes.c_void_p, "hSearchResult"]),
    COMMETHOD([], HRESULT, "GetFirstRow", [["in"], ctypes.c_void_p, "hSearchResult"]),
    COMMETHOD([], HRESULT, "GetNextRow", [["in"], ctypes.c_void_p, "hSearchResult"]),
    COMMETHOD([], HRESULT, "GetPreviousRow", [["in"], ctypes.c_void_p, "hSearchResult"]),
    STDMETHOD(HRESULT, "GetNextColumnName", [ctypes.c_void_p, ctypes.c_void_p]),
    COMMETHOD([], HRESULT, "GetColumn", [["in"], ctypes.c_void_p, "hSearchResult"], [["in"], ctypes.c_void_p, "szColumnName"], [["out"], Ads.ads_search_column as ctypes.StructureClass, "pSearchColumn"]),
    COMMETHOD([], HRESULT, "FreeColumn", [["in"], ctypes.c_void_p, "pSearchColumn"]),
    COMMETHOD([], HRESULT, "CloseSearchHandle", [["in"], ctypes.c_void_p, "hSearchResult"]),
  ],
});

const FreeADsMem = new ctypes.WinDLL("activeds.dll").func("FreeADsMem", ctypes.c_int32, [ctypes.c_void_p]);
const ConvertSidToStringSidW = new ctypes.WinDLL("advapi32.dll").func("ConvertSidToStringSidW", ctypes.c_int32, [ctypes.c_void_p, ctypes.c_void_p]);
const LocalFree = new ctypes.WinDLL("kernel32.dll").func("LocalFree", ctypes.c_void_p, [ctypes.c_void_p]);

/** Bind to a directory object, with explicit credentials when given (Python: ADsOpenObject). */
async function bind(path: string): Promise<DispatchObject> {
  if (!options.user) return await GetObject(path, { dynamic: true });
  using namespace = await GetObject("LDAP:", { dynamic: true });
  return await namespace.OpenDSObject(path, options.user, options.pass ?? "", Auth.ADS_SECURE_AUTHENTICATION);
}

/** An ADS_SEARCHPREF_INFO array: [preference, value] pairs (booleans as ADSTYPE_BOOLEAN). */
function preferences(entries: readonly (readonly [number, number | boolean])[]): Buffer {
  const buffer = Buffer.alloc(SEARCHPREF_SIZE * entries.length);
  entries.forEach(([preference, value], index) => {
    const info = new Ads.ads_searchpref_info(buffer.subarray(index * SEARCHPREF_SIZE, (index + 1) * SEARCHPREF_SIZE));
    info.dwSearchPref = preference;
    info.vValue.dwType = typeof value === "boolean" ? AdsType.ADSTYPE_BOOLEAN : AdsType.ADSTYPE_INTEGER;
    info.vValue[UNION].Integer = Number(value);
  });
  return buffer;
}

// ─── Value conversion ────────────────────────────────────────────────

/** Large integers holding a FILETIME (100 ns ticks since 1601). */
/** Large integers converted from FILETIME (100 ns ticks since 1601) to ISO 8601. */
const FILETIMES = new Set(["badPasswordTime", "lastLogoff", "lastLogon", "lastLogonTimestamp", "pwdLastSet"]);

/** A binary SID in native memory as "S-1-5-21-...", formatted by Windows. */
function sidString(address: bigint): string | null {
  const text = new ctypes.c_void_p();
  if (!ConvertSidToStringSidW(address, ctypes.byref(text)) || !text.value) return null;
  try {
    return ctypes.wstring_at(text.value);
  } finally {
    LocalFree(text.value);
  }
}

function convert(type: number, value: any, attribute: string): unknown {
  switch (type) {
    case AdsType.ADSTYPE_DN_STRING:
    case AdsType.ADSTYPE_CASE_EXACT_STRING:
    case AdsType.ADSTYPE_CASE_IGNORE_STRING:
    case AdsType.ADSTYPE_PRINTABLE_STRING:
    case AdsType.ADSTYPE_NUMERIC_STRING:
    case AdsType.ADSTYPE_OBJECT_CLASS:
      // The string members of the union share the same LPWSTR.
      return value.CaseIgnoreString ? ctypes.wstring_at(value.CaseIgnoreString) : null;
    case AdsType.ADSTYPE_BOOLEAN:
      return value.Boolean !== 0;
    case AdsType.ADSTYPE_INTEGER:
      return value.Integer;
    case AdsType.ADSTYPE_LARGE_INTEGER: {
      const ticks = BigInt(value.LargeInteger.QuadPart);
      if (!FILETIMES.has(attribute)) return ticks;
      return new Date(Number(ticks) / 10_000 - 11_644_473_600_000).toISOString();
    }
    case AdsType.ADSTYPE_UTC_TIME: {
      const time = value.UTCTime;
      return new Date(Date.UTC(time.wYear, time.wMonth - 1, time.wDay, time.wHour, time.wMinute, time.wSecond, time.wMilliseconds)).toISOString();
    }
    case AdsType.ADSTYPE_OCTET_STRING: {
      const { dwLength, lpValue } = value.OctetString;
      // Empty values are skipped (filtered out by readColumn).
      if (!lpValue || !dwLength) return undefined;
      switch (attribute.toLowerCase()) {
        case "objectsid":
          return sidString(lpValue);
        case "objectguid": {
          const bytes = readMemory(lpValue, dwLength);
          return bytes.length === 16 ? new GUID(new Uint8Array(bytes)).toString() : bytes.toString("hex");
        }
        default:
          return readMemory(lpValue, dwLength);
      }
    }
    default:
      return `<unsupported type ${type}>`;
  }
}

/** The values of a column (always an array, also for single-valued attributes). */
function readColumn(column: any, attribute: string): unknown[] {
  const values = Array.from({ length: column.dwNumValues }, (_, index) => {
    const value = Ads._adsvalue.from_address(BigInt(column.pADsValues) + BigInt(index * ADSVALUE_SIZE));
    return convert(column.dwADsType, value[UNION], attribute);
  });
  return values.filter((value) => value !== undefined);
}

// ─── Search ──────────────────────────────────────────────────────────

type Row = Record<string, unknown>;

/** First value of an attribute. */
const first = (values: unknown) => (Array.isArray(values) ? values[0] : values);

/**
 * Rows of a search, read page by page while iterating (`for await`). Every
 * IDirectorySearch call runs on the COM executor: GetNextRow may wait for the
 * server (the next page), and the event loop stays free meanwhile.
 */
async function* search(directory: DispatchObject, filter: string, attributes: readonly string[]): AsyncGenerator<Row> {
  using searcher = dispatchPointer(directory).QueryInterface(IDirectorySearch);
  const scope = { base: Scope.ADS_SCOPE_BASE, onelevel: Scope.ADS_SCOPE_ONELEVEL, subtree: Scope.ADS_SCOPE_SUBTREE }[options.scope.toLowerCase()];
  if (scope === undefined) throw new TypeError(`Unknown scope ${options.scope}`);
  const prefs = preferences([
    [Pref.ADS_SEARCHPREF_PAGESIZE, Number(options["page-size"])],
    [Pref.ADS_SEARCHPREF_SEARCH_SCOPE, scope],
    [Pref.ADS_SEARCHPREF_ASYNCHRONOUS, true],
    [Pref.ADS_SEARCHPREF_CHASE_REFERRALS, Referrals.ADS_CHASE_REFERRALS_ALWAYS],
    [Pref.ADS_SEARCHPREF_CACHE_RESULTS, false],
  ]);
  await searcher.SetSearchPreference(prefs, prefs.length / SEARCHPREF_SIZE);

  // LPWSTR array of the attribute names; -1 (0xFFFFFFFF) for all attributes.
  const names = attributes.length ? ctypes.array(ctypes.c_wchar_p, attributes.length)(...attributes) : null;
  const handle = await searcher.ExecuteSearch(filter, names, attributes.length || 0xffffffff);
  let completed = false;
  try {
    const name = new ctypes.c_void_p();
    while ((await searcher.GetNextRow(handle)) !== S_ADS_NOMORE_ROWS) {
      const row: Row = {};
      while ((await searcher.GetNextColumnName(handle, ctypes.byref(name))) !== S_ADS_NOMORE_COLUMNS) {
        try {
          const attribute = ctypes.wstring_at(name.value!) ?? "";
          const column = await searcher.GetColumn(handle, name.value);
          try {
            row[attribute] = readColumn(column, attribute);
          } finally {
            await searcher.FreeColumn(column._buffer);
          }
        } finally {
          FreeADsMem(name.value);
        }
      }
      yield row;
    }
    completed = true;
  } finally {
    // Stopping early (break) asks the server to stop sending pages. Best effort:
    // asynchronous paged searches may refuse it (0x8007001F); CloseSearchHandle
    // releases the search in any case.
    if (!completed) {
      try {
        await searcher.AbandonSearch(handle);
      } catch {
        // Closed below.
      }
    }
    await searcher.CloseSearchHandle(handle);
    // The attribute names (and their strings) must outlive the search: ADSI reads them while paging.
    void names;
  }
}

// ─── Tree output ─────────────────────────────────────────────────────
//
// For tree UI components: every node has `id`, `title` and `_path` (its
// container names, root first). Containers also have `children`; objects keep
// their attributes.

interface TreeNode {
  id: number;
  title: string;
  _path: string[];
}

interface Container extends TreeNode {
  children: (Container | TreeObject)[];
}

type TreeObject = Row & TreeNode;

/** Containers of a DN, root first: "CN=John,OU=Staff,OU=Paris,DC=contoso,DC=com" gives ["Paris", "Staff"]. */
function containers(dn: string): string[] {
  return dn
    .split(/(?<!\\),/) // commas escaped as \, belong to the value
    .slice(1) // the object itself
    .filter((rdn) => !/^DC=/i.test(rdn))
    .map((rdn) => rdn.slice(rdn.indexOf("=") + 1).replace(/\\(.)/g, "$1"))
    .reverse();
}

/** 32-bit string hash for node ids (as the node-ctypes LDAP demo). */
function hashString(text: string): number {
  let hash = 0;
  for (let index = 0; index < text.length; index++) hash = ((hash << 5) - hash + text.charCodeAt(index)) | 0;
  return Math.abs(hash);
}

/**
 * Group objects by container path. Objects get `title` (their `title`
 * attribute, else the DN) and `id` (hash of the DN); containers get the hash of
 * their whole path, so that equally named containers in different places keep
 * distinct ids. Each level lists its containers first, then its objects, by name.
 */
function toTree(rows: readonly Row[]): (Container | TreeObject)[] {
  const root: (Container | TreeObject)[] = [];
  const nodes = new Map<string, Container>();
  for (const row of rows) {
    const dn = String(first(row.distinguishedName) ?? first(row.ADsPath) ?? "");
    const path = containers(dn);
    let siblings = root;
    path.forEach((name, depth) => {
      const key = path.slice(0, depth + 1).join("/");
      let node = nodes.get(key);
      if (!node) {
        node = { id: hashString(key), title: name, _path: path.slice(0, depth + 1), children: [] };
        nodes.set(key, node);
        siblings.push(node);
      }
      siblings = node.children;
    });
    const title = first(row.title);
    siblings.push({ ...row, _path: path, title: typeof title === "string" && title ? title : dn, id: hashString(dn) });
  }
  const isContainer = (node: Container | TreeObject): node is Container => "children" in node && Array.isArray(node.children);
  const key = (node: Container | TreeObject) => (isContainer(node) ? `0${node.title}` : `1${String(first(node.distinguishedName) ?? node.title)}`);
  const sort = (nodes: (Container | TreeObject)[]) => {
    nodes.sort((a, b) => key(a).localeCompare(key(b)));
    for (const node of nodes) if (isContainer(node)) sort(node.children);
  };
  sort(root);
  return root;
}

// ─── Main ────────────────────────────────────────────────────────────

try {
  const limit = options.limit === undefined ? Infinity : Number(options.limit);
  if (limit !== Infinity && (!Number.isInteger(limit) || limit < 0)) throw new TypeError(`--limit must be a non-negative integer, got ${options.limit}`);

  let base = options.base;
  if (!base) {
    using rootDSE = await bind("LDAP://RootDSE");
    base = `LDAP://${await rootDSE.Get("defaultNamingContext")}`;
  }
  const attributes = options.attributes === "*" ? [] : options.attributes.split(",").map((attribute) => attribute.trim());
  // The tree is built from the distinguished names.
  if (options.tree && attributes.length && !attributes.some((attribute) => attribute.toLowerCase() === "distinguishedname")) attributes.push("distinguishedName");

  console.log(`Searching ${base} for ${options.filter}`);
  const started = performance.now();
  using directory = await bind(base);
  const rows: Row[] = [];
  for await (const row of search(directory, options.filter, attributes)) rows.push(row);
  const elapsed = performance.now() - started;
  rows.sort((a, b) => String(first(a.distinguishedName) ?? "").localeCompare(String(first(b.distinguishedName) ?? "")));

  if (options.debug) {
    // console.dir shows 100 array items by default: print every object, or `limit` of them.
    const shown = rows.slice(0, limit);
    console.dir(options.tree ? toTree(shown) : shown, { depth: null, maxArrayLength: null, maxStringLength: null });
    if (rows.length > limit) console.log(`... ${rows.length - limit} more (--limit)`);
  }
  console.log(`${rows.length} object(s) in ${elapsed.toFixed(0)} ms`);
} catch (error) {
  // COMError messages include the method, the HRESULT and the ADSI description.
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
}
