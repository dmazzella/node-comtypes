---
title: Building, testing and internals
---

# Building, testing and internals

## Scripts

```sh
npm run build        # compile TypeScript to dist/
npm run typecheck    # strict typecheck of src/, tests/, tests/types/ and examples/
npm run test:types   # declaration and API type tests
npm run test:unit    # OS-independent unit tests (GUID, HRESULT, layouts, DECIMAL, DATE, VARIANT scalars, typelib build/codegen)
npm test             # full build + type tests + all COM tests (COM tests skipped outside Windows)
npm run test:mirror  # full suite under NODE_CTYPES_NO_EXTERNAL_BUFFERS=1 (Electron / mirror mode)
npm run docs         # generate TypeDoc HTML documentation in docs/
npm run docs:serve   # generate docs and serve locally
npm pack --dry-run   # inspect npm package contents
```

## Electron verification

To verify real Electron main-process behavior (STA UI thread, Chromium message pump, disabled external ArrayBuffers):

```sh
npm run build
electron tests/electron/main.mjs
```

## Internal architecture

Detailed implementation flows (in Italian) are documented in [FLOWS.md](../FLOWS.md).

## CI and publishing

- The [CI workflow](../.github/workflows/ci.yml) runs `npm test` and `npm pack --dry-run` on Node.js 24 and 26 across Windows x64 (`windows-latest`) and Windows ARM64 (`windows-11-arm`).
- Pushing a `v<version>` tag that matches `package.json` publishes to npm via Trusted Publishing (OIDC) once the entire matrix passes (prerelease versions containing `-` publish under the `next` dist-tag).
- The [Docs workflow](../.github/workflows/docs.yml) builds the TypeDoc site (`npm run docs`) and deploys it to GitHub Pages on release tags or manual dispatch.
