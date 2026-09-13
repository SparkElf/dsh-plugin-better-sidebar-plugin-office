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
 * Read a CSS pixel length, rejecting anything that cannot size a page.
 *
 * @param value - a computed or inline length such as `"794px"`.
 * @returns the positive pixel count, or null when the value is absent, zero or
 * unparseable.
 */
export declare function parsePixelLength(value: string | null | undefined): number | null;
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
export declare function docxFitScale(pageWidth: number | null, availableWidth: number | null): number;
/**
 * Width the page may occupy inside the wrapper that holds it.
 *
 * The wrapper pads its own content, so `clientWidth` still includes that
 * padding: sizing the page against it would overrun by exactly that much.
 *
 * @param wrapper - the `.docx-wrapper` element docx-preview created.
 * @returns the content-box width in CSS pixels, or null when it cannot be read.
 */
export declare function wrapperContentWidth(wrapper: Element | null): number | null;
