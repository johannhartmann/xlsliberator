# LibreOffice source integration

XLSLiberator pins its sole target to LibreOffice full build `26.2.4.2`. Source
fetching, configuration, compilation, testing, and PyUNO probing run only in the
pinned Docker source-build environment. The host wrapper invokes Docker; it must
never invoke Python, LibreOffice, UNO, PyUNO, or `soffice` directly.

## Reproducible workflow

```bash
docker compose --profile office-source build office-source-fetch
./tools/office fetch libreoffice
./tools/office build libreoffice --variant stock --jobs 4
./tools/office test libreoffice --variant stock
./tools/office build libreoffice --variant patched --jobs 4
./tools/office test libreoffice --variant patched
./tools/office worktree libreoffice --name issue-name --with-patches
```

The manifest at `office/libreoffice/manifest.json` pins the upstream tag and
commit, source archive checksum, Debian snapshot and base-image digest, build
options, and patch queue. `tools/office.py` verifies those inputs before use.
Build caches may improve speed but cannot change the checked source, patch, or
output identities. The build deliberately uses LibreOffice's fetched NSS source
instead of Debian bookworm's older system NSS, because the pinned bundled
`xmlsec1` requires a newer crypto backend than that snapshot provides.

Each build writes an identity manifest containing compiler and package versions,
configuration flags, source and patch hashes, architecture, and source-built
binary hashes. The wrapper then builds a selectable runtime image, resolves its
immutable image ID, probes it with the bundled LibreOffice Python/PyUNO, and
merges the runtime, binary, Python, PyUNO, and package identities into the same
evidence. Stock-source and patched images have different mandatory variant
labels; runtime selection fails closed if either image tries to masquerade as
the other.

## Patch and upstream contribution workflow

1. Preserve a minimized redistributable regression fixture and prove the pinned
   stock source build fails it.
2. Create an isolated source worktree with `./tools/office worktree`.
3. Add the smallest appropriate upstream unit or integration test and patch the
   responsible LibreOffice subsystem.
4. Build and test the patched variant, then run the XLSLiberator regression
   subset against its explicitly selected image.
5. Export the commit as a numbered patch under
   `office/libreoffice/patches/`, update `series` and the manifest hashes, and
   preserve the stock/patched differential evidence.
6. Submit the change through LibreOffice Gerrit with its upstream issue and
   tests. Record the Gerrit change, upstream commit, and eventual release in the
   manifest.

Backports remain in the versioned queue only while the pinned `26.2.4.2`
baseline needs them. A dependency update must first prove the new stock build
passes the regression; only then may it remove an upstreamed patch. Unrelated
patches are never bundled into a compatibility backport.

## Licensing and distribution

LibreOffice source is dual-licensed under MPL 2.0 and LGPL 3.0-or-later; the
exact upstream notices (`COPYING`, `COPYING.MPL`, and `COPYING.LGPL`) accompany
the fetched source. Patch entries record their license and upstream provenance.
XLSLiberator itself is GPL-3.0-or-later, so distributing the combined product
must also satisfy the GPL source, notice, and corresponding-source obligations
documented in the repository `LICENSE`.
Anyone distributing source-built runtime images must provide the corresponding
source, applied patch queue, build scripts, license notices, and any relinking or
source-offer material required by the selected license. Container convenience
does not remove those obligations. Before public distribution, release owners
must review the complete image contents and dependency licenses; generated
identity evidence is provenance, not legal advice.
