import { useEffect, useRef } from 'react'
import { createUniver, LocaleType } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import '@univerjs/preset-sheets-core/lib/index.css'
import type { WorkbookModel } from '../model'
import { toUniverSnapshot } from '../adapters/univerSnapshot'

interface Props {
  workbook: WorkbookModel
}

/**
 * Approach C — render with the Univer spreadsheet SDK (Apache-2.0), canvas-based
 * with native Excel-like selection/navigation. The workbook is loaded from a
 * client-side snapshot and set read-only; editing UI is hidden.
 */
export default function UniverView({ workbook }: Props) {
  const containerRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const container = containerRef.current
    if (!container) return

    const { univerAPI } = createUniver({
      locale: LocaleType.EN_US,
      locales: { [LocaleType.EN_US]: UniverPresetSheetsCoreEnUS },
      presets: [
        UniverSheetsCorePreset({
          container,
          header: false, // hide the editing toolbar/ribbon — this is a read-only preview
          footer: { sheetBar: true, statisticBar: true, zoomSlider: true }, // keep sheet tabs + status bar
          formulaBar: true, // keep the Excel-like formula bar
          disableAutoFocus: true, // don't grab focus / enter the cell editor
        }),
      ],
    })

    const fWorkbook = univerAPI.createWorkbook(toUniverSnapshot(workbook))
    fWorkbook.setEditable(false) // read-only

    return () => {
      univerAPI.dispose()
      if (container) container.innerHTML = '' // safety net for React StrictMode remounts
    }
  }, [workbook])

  return <div className="univer-host" ref={containerRef} />
}
