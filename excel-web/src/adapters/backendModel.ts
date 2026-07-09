import type { WorkbookModel } from '../model'

export interface LoadResult {
  model: WorkbookModel
  /** Wall-clock time to fetch + deserialize, in ms. */
  ms: number
}

/**
 * Approach A — the ASP.NET backend parsed the file with ClosedXML and returns
 * the shared WorkbookModel as JSON. The browser only deserializes it.
 */
export async function loadFromBackend(
    selectedFile: File
): Promise<LoadResult> {
  const t0 = performance.now()

  const formData = new FormData()
  formData.append('file', selectedFile)

  const res = await fetch('/api/workbook', {
    method: 'POST',
    body: formData,
  })

  if (!res.ok) {
    throw new Error(
        `Backend /api/workbook responded with ${res.status}`
    )
  }

  const model = (await res.json()) as WorkbookModel

  return {
    model,
    ms: performance.now() - t0,
  }
}
