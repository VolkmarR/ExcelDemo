import { lazy, Suspense, useEffect, useState } from 'react'
import './App.css'
import SheetGrid from './grid/SheetGrid'
import { loadFromBackend, type LoadResult } from './adapters/backendModel'
import { loadFromExcelJs } from './adapters/exceljsModel'

// Univer is large; load it as its own chunk so the other two tabs don't pay for it.
const UniverView = lazy(() => import('./views/UniverView'))

type Approach = 'backend' | 'exceljs' | 'univer'

const TABS: { id: Approach; label: string; desc: string }[] = [
  { id: 'backend', label: 'Backend · ClosedXML', desc: 'Parsed on the .NET server (ClosedXML) → JSON model → custom virtualized grid.' },
  { id: 'exceljs', label: 'Frontend · ExcelJS', desc: 'Raw .xlsx bytes parsed in the browser (ExcelJS) → same custom grid. Nothing leaves the browser.' },
  { id: 'univer', label: 'Univer SDK', desc: 'Same model rendered by the Univer spreadsheet engine (canvas), read-only, client-side.' },
]

function stats(res: LoadResult): string {
  const cells = res.model.sheets.reduce((n, s) => n + s.cells.length, 0)
  const sheets = res.model.sheets.length
  return `${Math.round(res.ms)} ms · ${sheets} sheet${sheets === 1 ? '' : 's'} · ${cells.toLocaleString()} cells`
}

function App() {
  const [approach, setApproach] = useState<Approach>('backend')
  const [backend, setBackend] = useState<LoadResult | null>(null)
  const [exceljs, setExceljs] = useState<LoadResult | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)

  // Univer renders the backend-parsed model, so both those tabs share it.
  const needBackend = approach === 'backend' || approach === 'univer'
  const active = needBackend ? backend : exceljs

  useEffect(() => {
    if (active) return
    let cancelled = false
    setLoading(true)
    setError(null)
    const task = needBackend ? loadFromBackend() : loadFromExcelJs()
    task
      .then((res) => {
        if (cancelled) return
        if (needBackend) setBackend(res)
        else setExceljs(res)
      })
      .catch((e: unknown) => {
        if (!cancelled) setError(e instanceof Error ? e.message : String(e))
      })
      .finally(() => {
        if (!cancelled) setLoading(false)
      })
    return () => {
      cancelled = true
    }
  }, [approach, active, needBackend])

  const tab = TABS.find((t) => t.id === approach)!

  return (
    <div className="app">
      <header className="topbar">
        <div className="titles">
          <span className="brand">Excel Preview POC</span>
          <span className="desc">{tab.desc}</span>
        </div>
        <nav className="approach-tabs" role="tablist">
          {TABS.map((t) => (
            <button
              key={t.id}
              type="button"
              role="tab"
              aria-selected={t.id === approach}
              className={t.id === approach ? 'active' : ''}
              onClick={() => setApproach(t.id)}
            >
              {t.label}
            </button>
          ))}
        </nav>
        <div className="status">{active && !loading ? stats(active) : ' '}</div>
      </header>

      <main className="viewport">
        {loading && <div className="msg">Loading DemoData.xlsx…</div>}
        {error && (
          <div className="msg error">
            {error}
            <br />
            Make sure the ExcelApi backend is running on http://localhost:5269.
          </div>
        )}
        {!loading && !error && active && approach === 'backend' && <SheetGrid key="backend" workbook={active.model} />}
        {!loading && !error && active && approach === 'exceljs' && <SheetGrid key="exceljs" workbook={active.model} />}
        {!loading && !error && active && approach === 'univer' && (
          <Suspense fallback={<div className="msg">Loading Univer…</div>}>
            <UniverView key="univer" workbook={active.model} />
          </Suspense>
        )}
      </main>
    </div>
  )
}

export default App
