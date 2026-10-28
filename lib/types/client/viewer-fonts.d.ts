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
    readonly family: string;
    readonly source: string;
}
declare global {
    interface Window {
        /** Read-only list the font provider plugin installs; absent when it is not loaded. */
        __DSH_OFFICE_VIEWER_FONTS__?: readonly ViewerFontResource[];
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
export declare function readViewerFonts(value: unknown): ViewerFontResource[];
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
export declare function installViewerFonts(): Promise<number>;
export {};
