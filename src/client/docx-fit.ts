/**
 * Page fitting for the docx-preview surface.
 *
 * docx-preview writes every page at the width the document declares in its own
 * `pgSz` (A4 is roughly 794px) and centres the pages inside a flex wrapper. A
 * sidebar is narrower than that, so the page overruns the grey surround. The
 * library exposes no scale option: `ignoreWidth` drops the declared width and
 * reflows the text to the container, which loses the margins, line breaks and
 * pagination that a preview exists to show.
 *
 * So the page box keeps its document size and the pane scales it down instead.
 * Both helpers here are pure, which keeps the arithmetic testable without a DOM.
 */

/**
 * Read a CSS pixel length, rejecting anything that cannot size a page.
 *
 * @param value - a computed or inline length such as `"794px"`.
 * @returns the positive pixel count, or null when the value is absent, zero or
 * unparseable.
 */
export function parsePixelLength(value: string | null | undefined): number | null {
  if (value === null || value === undefined || value === '') return null
  const parsed = Number.parseFloat(value)
  if (!Number.isFinite(parsed) || parsed <= 0) return null
  return parsed
}

/**
 * Scale that fits one document page into the width the pane can give it.
 *
 * Never enlarges: a page that already fits keeps its size, so the type scale of
 * a narrow document stays comparable to the paper rather than stretching to
 * fill the pane. A missing measurement also yields 1, which leaves the page at
 * its document size instead of guessing.
 *
 * @param pageWidth - the page width the document declares, in CSS pixels.
 * @param availableWidth - the width the pane can give the page, after padding.
 * @returns a factor in (0, 1].
 */
export function docxFitScale(pageWidth: number | null, availableWidth: number | null): number {
  if (pageWidth === null || availableWidth === null) return 1
  if (!Number.isFinite(pageWidth) || !Number.isFinite(availableWidth)) return 1
  if (pageWidth <= 0 || availableWidth <= 0) return 1
  return Math.min(1, availableWidth / pageWidth)
}
