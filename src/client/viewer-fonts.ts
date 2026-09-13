/**
 * Fonts the surrounding runtime publishes for Office previews.
 *
 * The document preview has to draw with the fonts the document declares: a
 * government notice set in 仿宋_GB2312 shows the wrong face, and often the wrong
 * line breaks, when the browser falls back to a default sans. Those fonts are
 * not web-safe, so the runtime that ships them publishes their URLs on a global
 * for any viewer to pick up; nothing here hard-codes a font path or a vendor.
 */

/** One family the runtime offers, with an alias that may be a document's name for it. */
interface ViewerFontResource {
  readonly family: string
  readonly source: string
}

declare global {
  interface Window {
    /** Read-only list the font provider plugin installs; absent when it is not loaded. */
    __DSH_OFFICE_VIEWER_FONTS__?: readonly ViewerFontResource[]
  }
}

/**
 * Narrow an unknown global to the font list the provider promises.
 *
 * The global crosses a plugin boundary, so it is validated rather than trusted:
 * a malformed entry would otherwise throw inside `FontFace` construction and take
 * the whole preview down with it.
 *
 * @param value - whatever the global currently holds.
 * @returns the usable entries, or an empty list when the shape does not match.
 */
export function readViewerFonts(value: unknown): ViewerFontResource[] {
  if (!Array.isArray(value)) return []
  const fonts: ViewerFontResource[] = []
  for (const entry of value) {
    if (entry === null || typeof entry !== 'object') continue
    const { family, source } = entry as { family?: unknown; source?: unknown }
    if (typeof family !== 'string' || family === '') continue
    if (typeof source !== 'string' || source === '') continue
    fonts.push({ family, source })
  }
  return fonts
}

/** One loaded face per family, so repeated previews reuse the same fetch. */
const registered = new Map<string, Promise<FontFace | null>>()

/**
 * Register one family with the document.
 *
 * A face that fails to load resolves to null instead of rejecting: a preview
 * that cannot fetch 仿宋 is still a readable preview, and the document falls back
 * to its next font exactly as it would in an editor without the font installed.
 *
 * @param font - the family name and the URL that serves it.
 * @returns the added face, or null when it could not be loaded.
 */
function registerFont(font: ViewerFontResource): Promise<FontFace | null> {
  const existing = registered.get(font.family)
  if (existing !== undefined) return existing
  if (typeof FontFace !== 'function') return Promise.resolve(null)
  const pending = (async (): Promise<FontFace | null> => {
    try {
      const face = await new FontFace(font.family, `url("${font.source}")`, {
        style: 'normal',
        weight: '400',
      }).load()
      document.fonts.add(face)
      return face
    } catch {
      // The preview keeps working with whatever the host does have.
      return null
    }
  })()
  registered.set(font.family, pending)
  return pending
}

/**
 * Load every published Office font and wait until the document can use them.
 *
 * Awaiting the faces matters for docx-preview: it measures text to place runs and
 * to lay out tables, so a face that arrives after that measurement reflows the
 * page. Resolving only once the faces are registered and `document.fonts.ready`
 * has settled means the first render already sees the real metrics.
 *
 * @returns the number of families this call newly loaded or found loaded.
 */
export async function installViewerFonts(): Promise<number> {
  const fonts = readViewerFonts(typeof window === 'undefined' ? undefined : window.__DSH_OFFICE_VIEWER_FONTS__)
  if (fonts.length === 0) return 0
  const results = await Promise.all(fonts.map(font => registerFont(font)))
  // Faces already loaded on an earlier preview are counted here too, which keeps
  // the return value meaningful for a caller that renders many documents.
  if (typeof document !== 'undefined' && document.fonts !== undefined) await document.fonts.ready
  return new Set(results.map((face, index) => face === null ? null : fonts[index]?.family).filter(Boolean)).size
}
