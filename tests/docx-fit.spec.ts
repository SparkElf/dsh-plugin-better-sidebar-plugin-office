/**
 * docx page-fitting arithmetic. Both helpers are pure, so they are testable
 * without a DOM or a rendered document.
 */
import { describe, expect, it } from 'vitest'
import { docxFitScale, parsePixelLength } from '../src/client/docx-fit.ts'

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

  it('survives the degenerate widths a hidden pane reports', () => {
    expect(docxFitScale(0, 400)).toBe(1)
    expect(docxFitScale(794, 0)).toBe(1)
    expect(docxFitScale(Number.NaN, 400)).toBe(1)
    expect(docxFitScale(794, Number.POSITIVE_INFINITY)).toBe(1)
    expect(docxFitScale(794, Number.NEGATIVE_INFINITY)).toBe(1)
  })
})
