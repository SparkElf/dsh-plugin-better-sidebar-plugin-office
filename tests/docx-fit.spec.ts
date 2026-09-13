/**
 * docx page-fitting arithmetic. The scale helpers are pure, so they are testable
 * without a rendered document; `wrapperContentWidth` reads the wrapper's padding
 * through `getComputedStyle`, which the environment provides.
 */
import { describe, expect, it, vi } from 'vitest'
import { docxFitScale, parsePixelLength, wrapperContentWidth } from '../src/client/docx-fit.ts'

describe('parsePixelLength', () => {
  it('reads a computed pixel length', () => {
    expect(parsePixelLength('794px')).toBe(794)
    expect(parsePixelLength('612.5px')).toBe(612.5)
  })

  it('rejects values that cannot size a page', () => {
    expect(parsePixelLength(null)).toBeNull()
    expect(parsePixelLength(undefined)).toBeNull()
    expect(parsePixelLength('')).toBeNull()
    expect(parsePixelLength('auto')).toBeNull()
    expect(parsePixelLength('0px')).toBeNull()
    expect(parsePixelLength('-12px')).toBeNull()
  })
})

describe('docxFitScale', () => {
  it('shrinks a page wider than the pane', () => {
    // A4 at 96dpi is ~794px, while a sidebar pane is narrower; without this the
    // page overruns the grey surround.
    expect(docxFitScale(794, 400)).toBeCloseTo(400 / 794, 5)
  })

  it('never enlarges a page that already fits', () => {
    // A narrow document keeps its true type scale instead of stretching, so the
    // preview still shows the real page proportions.
    expect(docxFitScale(400, 794)).toBe(1)
    expect(docxFitScale(794, 794)).toBe(1)
  })

  it('leaves the page alone when a measurement is missing', () => {
    // Guessing here would misreport the page size; the document's own width is
    // the honest fallback.
    expect(docxFitScale(null, 400)).toBe(1)
    expect(docxFitScale(794, null)).toBe(1)
  })

  it('survives the degenerate widths a hidden wrapper reports', () => {
    expect(docxFitScale(0, 400)).toBe(1)
    expect(docxFitScale(794, 0)).toBe(1)
    expect(docxFitScale(Number.NaN, 400)).toBe(1)
    expect(docxFitScale(794, Number.POSITIVE_INFINITY)).toBe(1)
    expect(docxFitScale(794, Number.NEGATIVE_INFINITY)).toBe(1)
  })
})

describe('wrapperContentWidth', () => {
  /** A stand-in for the wrapper docx-preview creates, with its own padding. */
  const wrapper = (clientWidth: number, paddingLeft: string, paddingRight: string): Element => ({
    clientWidth,
  }) as unknown as Element

  it('subtracts the padding the library paints inside the wrapper', () => {
    // This is the measurement the first attempt got wrong: docx-preview pads its
    // own wrapper 30px a side, so sizing the page against clientWidth alone
    // overruns by exactly that padding.
    vi.stubGlobal('getComputedStyle', () => ({ paddingLeft: '30px', paddingRight: '30px' }))
    expect(wrapperContentWidth(wrapper(432, '30px', '30px'))).toBe(372)
    vi.unstubAllGlobals()
  })

  it('reports nothing for a wrapper it cannot measure', () => {
    vi.stubGlobal('getComputedStyle', () => ({ paddingLeft: '30px', paddingRight: '30px' }))
    expect(wrapperContentWidth(null)).toBeNull()
    // A pane that is collapsed or hidden leaves no width to fit into.
    expect(wrapperContentWidth(wrapper(0, '30px', '30px'))).toBeNull()
    expect(wrapperContentWidth(wrapper(30, '30px', '0px'))).toBeNull()
    vi.unstubAllGlobals()
  })
})
