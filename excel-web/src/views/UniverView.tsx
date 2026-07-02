import { useEffect, useRef } from 'react'
import { createUniver, LocaleType } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import '@univerjs/preset-sheets-core/lib/index.css'
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing'
import UniverPresetSheetsDrawingEnUS from '@univerjs/preset-sheets-drawing/locales/en-US'
import '@univerjs/preset-sheets-drawing/lib/index.css'
import type { Picture, WorkbookModel } from '../model'
import { toUniverSnapshot } from '../adapters/univerSnapshot'

interface Props {
  workbook: WorkbookModel
}

// Minimal shape of the over-grid-image facade. The drawing preset registers these
// methods on FWorksheet at runtime (its ESM entry imports the sheets-drawing facade
// as a side effect); we type them locally so we don't depend on the facade's
// `declare module` augmentation being pulled onto the TS module graph.
interface ImageBuilder {
  setSource(src: string): ImageBuilder
  setColumn(c: number): ImageBuilder
  setRow(r: number): ImageBuilder
  setColumnOffset(px: number): ImageBuilder
  setRowOffset(px: number): ImageBuilder
  setWidth(px: number): ImageBuilder
  setHeight(px: number): ImageBuilder
  buildAsync(): Promise<unknown>
}
interface DrawingWorksheet {
  newOverGridImage(): ImageBuilder
  insertImages(images: unknown[]): void
}

async function buildImages(fSheet: DrawingWorksheet, pictures: Picture[]): Promise<unknown[]> {
  const images: unknown[] = []
  for (const p of pictures) {
    const image = await fSheet
      .newOverGridImage()
      .setSource(p.src) // data: URL renders directly — no network, no external service
      .setColumn(p.fromCol)
      .setRow(p.fromRow)
      .setColumnOffset(p.offsetX)
      .setRowOffset(p.offsetY)
      .setWidth(p.width)
      .setHeight(p.height)
      .buildAsync()
    images.push(image)
  }
  return images
}

/**
 * Approach C — render with the Univer spreadsheet SDK (Apache-2.0), canvas-based
 * with native Excel-like selection/navigation. The workbook is loaded from a
 * client-side snapshot and set read-only; editing UI is hidden. Floating pictures
 * are inserted via the OSS drawing preset's over-grid-image facade.
 */
export default function UniverView({ workbook }: Props) {
  const containerRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const container = containerRef.current
    if (!container) return

    const { univerAPI } = createUniver({
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: { ...UniverPresetSheetsCoreEnUS, ...UniverPresetSheetsDrawingEnUS },
      },
      presets: [
        UniverSheetsCorePreset({
          container,
          header: false, // hide the editing toolbar/ribbon — this is a read-only preview
          footer: { sheetBar: true, statisticBar: true, zoomSlider: true }, // keep sheet tabs + status bar
          formulaBar: true, // keep the Excel-like formula bar
          disableAutoFocus: true, // don't grab focus / enter the cell editor
        }),
        UniverSheetsDrawingPreset(), // OSS image/drawing support
      ],
    })

    const fWorkbook = univerAPI.createWorkbook(toUniverSnapshot(workbook))

    // Insert floating pictures, then lock the workbook read-only. Images must be
    // added while the workbook is still editable, so this runs before setEditable(false).
    let disposed = false
    void (async () => {
      try {
        for (const sheet of workbook.sheets) {
          if (!sheet.pictures?.length || disposed) continue
          const fSheet = (
            fWorkbook as unknown as { getSheetByName(name: string): DrawingWorksheet | null }
          ).getSheetByName(sheet.name)
          if (!fSheet) continue
          const images = await buildImages(fSheet, sheet.pictures)
          if (!disposed) fSheet.insertImages(images)
        }
      } catch {
        /* best-effort: a drawing failure must never break the preview */
      } finally {
        if (!disposed) fWorkbook.setEditable(false) // read-only
      }
    })()

    return () => {
      disposed = true
      univerAPI.dispose()
      if (container) container.innerHTML = '' // safety net for React StrictMode remounts
    }
  }, [workbook])

  return <div className="univer-host" ref={containerRef} />
}
