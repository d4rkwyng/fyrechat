# Notices — bundled third-party software

FyreChat's own code is MIT-licensed (see [LICENSE](LICENSE.md)). It also
bundles a small number of third-party binaries that are **not** covered by
that license. This file records what they are, where they came from, and on
what basis they ship.

## bncsutil.dll (LGPL 2.1)

`src/bncsutil.dll`

[BNCSutil](https://github.com/BNETDocs/bncsutil) ("Battle.Net Chat Service
Utility"), originally by Eric Naeseth, now maintained by the open Battle.net
community. FyreChat uses two of its exports, `hashPassword` and
`doubleHashPassword`, to compute the classic Battle.net "OLS" (old-auth,
Broken SHA-1) password hash locally instead of sending the plaintext account
password to a BNLS relay server for that step — see `src/class
modules/clsOldAuth.cls` and its call sites in `src/modules/modBNCS.bas`.

- Version: `v1.4.5`, 32-bit (x86) Windows build, from the project's
  [GitHub Releases](https://github.com/BNETDocs/bncsutil/releases/tag/v1.4.5)
  (`bncsutil_v1.4.5_x86_dll.zip`).
- SHA-256 of the bundled `bncsutil.dll`:
  `515a0e13af8df9a44f71467d6ba7867583d8768a2245c5a9eb13e27845e3a545`
- License: GNU Lesser General Public License v2.1 (see the project's
  [`COPYING`](https://github.com/BNETDocs/bncsutil/blob/master/COPYING) file).
  FyreChat loads it as a dynamically-linked DLL at runtime (never statically
  linked, never modified), which is the LGPL's own dynamic-linking
  accommodation for works that merely *use* the library — no part of
  FyreChat itself is LGPL-licensed as a result. The library can be replaced
  in place with any compatible build of `bncsutil.dll`, and its full source
  is available at the link above.

## BCEL.dll, Storm.dll and other bundled runtime files

`src/BCEL.dll` and the runtime files under `assets/required_files.zip` /
`assets/vc7_crt.zip` predate this notice and were inherited from the
FyReBoT source FyreChat was developed from; their individual provenance
isn't tracked here. This section covers only `bncsutil.dll`, added
alongside the local OLS password hashing change.
