# Attribution

This package is a redistribution of
[`@huanlin/dsh-plugin-better-sidebar-plugin-office`](https://github.com/HuanLinOTO/dsh-plugin-better-sidebar-plugin-office)
by Huanlin, which is licensed under the GNU Affero General Public License v3.0.
This redistribution stays under the same licence; the full text is in `LICENSE`.

## Origin

- Upstream: https://github.com/HuanLinOTO/dsh-plugin-better-sidebar-plugin-office
- Base: the `adapt/v0.1.2-alpha.1` line, version 0.2.0
- Copyright (C) 2026 Huanlin — retained in full

## Changes in this redistribution

1. **The package is renamed** to `@sparkelf/dsh-plugin-better-sidebar-office` so it can
   be published and pinned independently of the upstream release cycle.
2. **`fix(docx)`: the page fits the pane.** docx-preview writes every page at the
   width the document declares in its own `pgSz` — A4 arrives as `595.25pt` — and
   centres the pages in a flex wrapper, so a page wider than the sidebar overran the
   grey surround on both sides. The page now keeps its document size and the wrapper
   scales it to the space it actually has.
3. **`feat(docx)`: previews draw with the document's own fonts.** The government
   document faces (`方正小标宋简体`, `FangSong_GB2312`, `KaiTi_GB2312`, `SimHei`) are not
   web-safe, so the preview fell back to a default sans and the line breaks moved
   with it. The families the runtime publishes on
   `window.__DSH_OFFICE_VIEWER_FONTS__` are now registered before rendering.

Both changes were also offered upstream as pull requests, so the upstream package
may carry equivalent fixes in a later release.

## Source

Corresponding source for this build is published at
https://github.com/SparkElf/dsh-plugin-better-sidebar-plugin-office, which satisfies
the AGPL requirement that a network-distributed derivative offer its complete
corresponding source.
