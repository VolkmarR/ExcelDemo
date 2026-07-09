import { lazy, Suspense, useEffect, useState } from 'react'
import './App.css'

import SheetGrid from './grid/SheetGrid'
import { loadFromBackend, type LoadResult } from './adapters/backendModel'
import { loadFromExcelJs } from './adapters/exceljsModel'
import SyncfusionView from './views/SyncfusionView'
import Home from "./views/Home.tsx";

const UniverView = lazy(() => import('./views/UniverView'))
const ReactDataGridView = lazy(() => import('./views/ReactDataGridView'))
const JspreadsheetView = lazy(() => import('./views/JspreadsheetView'))

type Approach =
    | 'home'
    | 'backend'
    | 'exceljs'
    | 'univer'
    | 'rdg'
    | 'jss'
    | 'syncfusion'

const TABS: { id: Approach; label: string; desc: string }[] = [
    {
        id: 'home',
        label: 'Home',
        desc: 'Select an Excel workbook'
    },
    {
        id: 'backend',
        label: 'Backend · ClosedXML',
        desc: 'Parsed on the .NET server (ClosedXML) → JSON model → custom virtualized grid.'
    },
    {
        id: 'exceljs',
        label: 'Frontend · ExcelJS',
        desc: 'Raw .xlsx bytes parsed in the browser (ExcelJS) → same custom grid.'
    },
    {
        id: 'univer',
        label: 'Univer SDK',
        desc: 'Same model rendered by the Univer spreadsheet engine.'
    },
    {
        id: 'rdg',
        label: 'react-data-grid · OSS',
        desc: 'Same ClosedXML JSON model rendered by react-data-grid.'
    },
    {
        id: 'jss',
        label: 'Jspreadsheet CE · OSS',
        desc: 'Same ClosedXML JSON model rendered by Jspreadsheet CE.'
    },
    {
        id: 'syncfusion',
        label: 'Syncfusion',
        desc: 'Syncfusion Viewer'
    }
]

function stats(res: LoadResult): string {
    const cells = res.model.sheets.reduce((n, s) => n + s.cells.length, 0)
    const sheets = res.model.sheets.length

    return `${Math.round(res.ms)} ms · ${sheets} sheet${
        sheets === 1 ? '' : 's'
    } · ${cells.toLocaleString()} cells`
}

function App() {
    const [approach, setApproach] = useState<Approach>('home')
    const [selectedFile, setSelectedFile] = useState<File | null>(null)

    const [backend, setBackend] = useState<LoadResult | null>(null)
    const [exceljs, setExceljs] = useState<LoadResult | null>(null)

    const [loading, setLoading] = useState(false)
    const [error, setError] = useState<string | null>(null)

    const needBackend =
        approach === 'backend' ||
        approach === 'univer' ||
        approach === 'rdg' ||
        approach === 'jss'

    const active = needBackend ? backend : exceljs

    useEffect(() => {
        if (approach === 'home') {
            return
        }

        if (!selectedFile) {
            setError('Please select an Excel file first.')
            return
        }

        if (approach === 'syncfusion') {
            return
        }

        if (active) {
            return
        }

        let cancelled = false

        setLoading(true)
        setError(null)

        const task = needBackend
            ? loadFromBackend(selectedFile)
            : loadFromExcelJs(selectedFile)

        task
            .then((res) => {
                if (cancelled) return

                if (needBackend) {
                    setBackend(res)
                } else {
                    setExceljs(res)
                }
            })
            .catch((e: unknown) => {
                if (!cancelled) {
                    setError(e instanceof Error ? e.message : String(e))
                }
            })
            .finally(() => {
                if (!cancelled) {
                    setLoading(false)
                }
            })

        return () => {
            cancelled = true
        }
    }, [approach, selectedFile, active, needBackend])

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

                <div className="status">
                    {active && !loading ? stats(active) : '\u00A0'}
                </div>
            </header>

            <main className="viewport">
                {approach === 'home' && (
                    <Home
                        selectedFile={selectedFile}
                        onFileSelected={(file) => {
                            setSelectedFile(file)

                            setBackend(null)
                            setExceljs(null)
                            setError(null)

                            if (file) {
                                setApproach('backend')
                            }
                        }}
                    />
                )}

                {loading && (
                    <div className="msg">
                        Loading {selectedFile?.name ?? 'workbook'}...
                    </div>
                )}

                {error && (
                    <div className="msg error">
                        {error}
                    </div>
                )}

                {!loading &&
                    !error &&
                    active &&
                    approach === 'backend' && (
                        <SheetGrid key="backend" workbook={active.model} />
                    )}

                {!loading &&
                    !error &&
                    active &&
                    approach === 'exceljs' && (
                        <SheetGrid key="exceljs" workbook={active.model} />
                    )}

                {!loading &&
                    !error &&
                    approach === 'syncfusion' && (
                        <SyncfusionView key="syncfusion" />
                    )}

                {!loading &&
                    !error &&
                    active &&
                    approach === 'univer' && (
                        <Suspense fallback={<div className="msg">Loading Univer…</div>}>
                            <UniverView key="univer" workbook={active.model} />
                        </Suspense>
                    )}

                {!loading &&
                    !error &&
                    active &&
                    approach === 'rdg' && (
                        <Suspense fallback={<div className="msg">Loading grid…</div>}>
                            <ReactDataGridView
                                key="rdg"
                                workbook={active.model}
                            />
                        </Suspense>
                    )}

                {!loading &&
                    !error &&
                    active &&
                    approach === 'jss' && (
                        <Suspense fallback={<div className="msg">Loading Jspreadsheet…</div>}>
                            <JspreadsheetView
                                key="jss"
                                workbook={active.model}
                            />
                        </Suspense>
                    )}
            </main>
        </div>
    )
}

export default App