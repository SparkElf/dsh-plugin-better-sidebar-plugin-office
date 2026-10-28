/**
 * Viewer font registration. The provider global and `FontFace` are both supplied
 * by the runtime, so the tests stand in for both and assert the behaviour a
 * preview depends on: the right families load, a bad entry is survivable, and a
 * failing font does not take the document down with it.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { installViewerFonts, readViewerFonts } from '../src/client/viewer-fonts.ts'

afterEach(() => {
  vi.unstubAllGlobals()
  delete (globalThis as { window?: unknown }).window
})

describe('readViewerFonts', () => {
  it('keeps every usable family the provider publishes', () => {
    // These are the names the government-document fonts are declared under, so a
    // document asking for 仿宋_GB2312 or FangSong_GB2312 finds the same face.
    const fonts = readViewerFonts([
      { family: '方正小标宋简体', source: '/office-viewer-fonts/FZXiaoBiaoSong.ttf?v=0.1.2' },
      { family: 'FangSong_GB2312', source: '/office-viewer-fonts/FangSongGB2312.ttf?v=0.1.2' },
      { family: '仿宋_GB2312', source: '/office-viewer-fonts/FangSongGB2312.ttf?v=0.1.2' },
    ])
    expect(fonts).toHaveLength(3)
    expect(fonts[1]).toEqual({
      family: 'FangSong_GB2312',
      source: '/office-viewer-fonts/FangSongGB2312.ttf?v=0.1.2',
    })
  })

  it('ignores entries that could not build a face', () => {
    expect(readViewerFonts(undefined)).toEqual([])
    expect(readViewerFonts(null)).toEqual([])
    expect(readViewerFonts('fonts')).toEqual([])
    expect(readViewerFonts([null, 7, {}, { family: '' }, { family: 'a', source: '' }])).toEqual([])
    expect(readViewerFonts([{ family: '仿宋', source: '/f.ttf' }])).toHaveLength(1)
  })
})

describe('installViewerFonts', () => {
  /** A FontFace stand-in that records what the module asked for. */
  const stubFonts = (outcome: 'load' | 'reject') => {
    const added: string[] = []
    const loaded: string[] = []
    class FakeFontFace {
      constructor(public family: string, public source: string) {}
      async load(): Promise<FakeFontFace> {
        loaded.push(this.family)
        if (outcome === 'reject') throw new Error('font unavailable')
        return this
      }
    }
    vi.stubGlobal('FontFace', FakeFontFace)
    vi.stubGlobal('document', {
      fonts: { add: (face: { family: string }) => added.push(face.family), ready: Promise.resolve() },
    })
    return { added, loaded }
  }

  it('registers each published family before returning', async () => {
    // docx-preview measures text as it renders, so the caller awaits this.
    const { added, loaded } = stubFonts('load')
    vi.stubGlobal('window', {
      __DSH_OFFICE_VIEWER_FONTS__: [
        { family: 'FangSong_GB2312', source: '/office-viewer-fonts/FangSongGB2312.ttf' },
        { family: 'SimHei', source: '/office-viewer-fonts/SimHei.ttf' },
      ],
    })
    await installViewerFonts()
    expect(loaded).toEqual(['FangSong_GB2312', 'SimHei'])
    expect(added).toEqual(['FangSong_GB2312', 'SimHei'])
  })

  it('leaves the preview usable when a font cannot be fetched', async () => {
    // A missing 仿宋 must degrade to the fallback, not fail the document.
    const { added } = stubFonts('reject')
    vi.stubGlobal('window', {
      __DSH_OFFICE_VIEWER_FONTS__: [{ family: 'KaiTi_GB2312', source: '/missing.ttf' }],
    })
    await expect(installViewerFonts()).resolves.toBe(0)
    expect(added).toEqual([])
  })

  it('does nothing when the font provider is not installed', async () => {
    const { loaded } = stubFonts('load')
    vi.stubGlobal('window', {})
    await expect(installViewerFonts()).resolves.toBe(0)
    expect(loaded).toEqual([])
  })
})
