/**
 * Page fitting for the docx-preview surface.
 *
 * docx-preview writes every page at the width the document declares in its own
 * `pgSz` — A4 is roughly 794px — and centres the pages inside a flex wrapper. A
 * sidebar is narrower than that, so the page overruns the grey surround on both
 * sides at once, because the wrapper centres it.
 *
 * The width that decides the fit is the **wrapper's content box**, not the pane.
 * The library paints its own padding inside the wrapper (30px a side), so a
 * factor computed against the pane is 60px too generous and still overflows.
 *
 * The library exposes no scale option: `ignoreWidth` drops the declared width
 * and reflows the text to the container, which loses the margins, line breaks
 * and pagination a preview exists to show. So the page keeps its document size
 * and the wrapper scales it.
 */

/**
 * Read a CSS length in device-independent pixels.
 *
 * docx-preview writes the page width from the document's own `pgSz`, which is
 * expressed in points — `"595.25pt"` for A4. Parsing that as a bare number
 * treats it as pixels and understates the page by a third (595 against a real
 * 794), so the scale it produces is too large and the page still overruns. The
 * unit is read here and converted, since only pixels compare with the space the
 * pane offers.
 *
 * @param value - a CSS length such as `"794px"` or `"595.25pt"`.
 * @returns the width in pixels, or null when it cannot size a page.
 */
export function parsePixelLength(value: string | null | undefined): number | null {
  if (value === null || value === undefined || value === '') return null
  const parsed = Number.parseFloat(value)
  if (!Number.isFinite(parsed) || parsed <= 0) return null
  const unit = /[a-z%]+/iu.exec(value.trim())?.[0].toLowerCase() ?? 'px'
  // point is 1/72in and a CSS pixel is 1/96in, so a point is 96/72 of a pixel.
  if (unit === 'pt') return parsed * (96 / 72)
  // Anything else (px included) is already a length the comparison can use.
  return parsed
}

/**
 * Scale that fits one document page into the width available to it.
 *
 * Never enlarges: a page that already fits keeps its size, so the type scale of
 * a narrow document stays comparable to the paper rather than stretching to fill
 * the space. A missing measurement also yields 1, which leaves the page at its
 * document size instead of guessing.
 *
 * @param pageWidth - the page width the document declares, in CSS pixels.
 * @param availableWidth - the width actually available to the page: the
 * wrapper's content box, which already excludes the padding the library paints.
 * @returns a factor in (0, 1].
 */
export function docxFitScale(pageWidth: number | null, availableWidth: number | null): number {
  if (pageWidth === null || availableWidth === null) return 1
  if (!Number.isFinite(pageWidth) || !Number.isFinite(availableWidth)) return 1
  if (pageWidth <= 0 || availableWidth <= 0) return 1
  return Math.min(1, availableWidth / pageWidth)
}

/**
 * Width the page may occupy inside the wrapper that holds it.
 *
 * The wrapper pads its own content, so `clientWidth` still includes that
 * padding: sizing the page against it would overrun by exactly that much.
 *
 * @param wrapper - the `.docx-wrapper` element docx-preview created.
 * @returns the content-box width in CSS pixels, or null when it cannot be read.
 */
export function wrapperContentWidth(wrapper: Element | null): number | null {
  // Structural check rather than an HTMLElement one: the DOM global is absent in
  // the unit environment, and the caller only ever passes a real element.
  if (wrapper === null || typeof getComputedStyle !== 'function') return null
  const style = getComputedStyle(wrapper)
  const left = Number.parseFloat(style.paddingLeft)
  const right = Number.parseFloat(style.paddingRight)
  const padding = (Number.isFinite(left) ? left : 0) + (Number.isFinite(right) ? right : 0)
  const width = wrapper.clientWidth - padding
  return width > 0 ? width : null
}
