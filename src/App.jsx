import { useEffect, useMemo, useState } from 'react'
import { onAuthStateChanged, signInWithPopup, signOut } from 'firebase/auth'
import {
  addDoc,
  collection,
  onSnapshot,
  orderBy,
  query,
  serverTimestamp,
} from 'firebase/firestore'
import { auth, db, googleProvider, hasFirebaseConfig } from './lib/firebase'

const DATA_STRUCTURE_VERSION = 1
const DEFAULT_TARGET_EMAIL = 'care-team@example.com'

const emptyForm = {
  date: new Date().toISOString().slice(0, 10),
  feedingMl: '0',
  sleepHours: '0',
  diaperCount: '0',
  notes: '',
}

function normalizeRecord(entry) {
  return {
    date: entry.date || '',
    feedingMl: Number(entry.feedingMl || 0),
    sleepHours: Number(entry.sleepHours || 0),
    diaperCount: Number(entry.diaperCount || 0),
    notes: entry.notes || '',
    schemaVersion: Number(entry.schemaVersion || 1),
  }
}

async function createWorkbookBuffer(records) {
  const { default: ExcelJS } = await import('exceljs')
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('Baby Daily')

  worksheet.columns = [
    { header: 'Date', key: 'date', width: 14 },
    { header: 'Feeding (ml)', key: 'feedingMl', width: 14 },
    { header: 'Sleep (hours)', key: 'sleepHours', width: 14 },
    { header: 'Diaper Count', key: 'diaperCount', width: 14 },
    { header: 'Notes', key: 'notes', width: 36 },
    { header: 'Schema Version', key: 'schemaVersion', width: 14 },
  ]

  records.forEach((record) => {
    worksheet.addRow(normalizeRecord(record))
  })

  return workbook.xlsx.writeBuffer()
}

function downloadBuffer(buffer, filename) {
  const file = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
  const url = URL.createObjectURL(file)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  link.click()
  URL.revokeObjectURL(url)
  return file
}

function App() {
  const [user, setUser] = useState(null)
  const [logs, setLogs] = useState([])
  const [form, setForm] = useState(emptyForm)
  const [message, setMessage] = useState('')
  const [isSaving, setIsSaving] = useState(false)
  const [isExporting, setIsExporting] = useState(false)

  const targetEmail = import.meta.env.VITE_REPORT_EMAIL || DEFAULT_TARGET_EMAIL

  useEffect(() => {
    if (!auth) {
      return undefined
    }

    const unsubscribe = onAuthStateChanged(auth, (nextUser) => {
      setUser(nextUser)
      if (!nextUser) {
        setLogs([])
      }
    })
    return unsubscribe
  }, [])

  useEffect(() => {
    if (!db || !user) {
      return undefined
    }

    const logsRef = collection(db, 'users', user.uid, 'dailyLogs')
    const logsQuery = query(logsRef, orderBy('date', 'desc'))

    const unsubscribe = onSnapshot(logsQuery, (snapshot) => {
      setLogs(snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() })))
    })

    return unsubscribe
  }, [user])

  const sortedLogs = useMemo(
    () => [...logs].sort((a, b) => a.date.localeCompare(b.date)),
    [logs],
  )

  const handleLogin = async () => {
    if (!auth || !googleProvider) {
      return
    }

    setMessage('')
    await signInWithPopup(auth, googleProvider)
  }

  const handleLogout = async () => {
    if (!auth) {
      return
    }

    await signOut(auth)
  }

  const handleSave = async (event) => {
    event.preventDefault()

    if (!db || !user) {
      return
    }

    setMessage('')
    setIsSaving(true)

    try {
      await addDoc(collection(db, 'users', user.uid, 'dailyLogs'), {
        ...normalizeRecord(form),
        schemaVersion: DATA_STRUCTURE_VERSION,
        createdAt: serverTimestamp(),
      })
      setForm(emptyForm)
      setMessage('Saved daily record.')
    } finally {
      setIsSaving(false)
    }
  }

  const handleDownload = async () => {
    setMessage('')

    if (!sortedLogs.length) {
      setMessage('No records to export.')
      return
    }

    setIsExporting(true)

    try {
      const buffer = await createWorkbookBuffer(sortedLogs)
      const filename = `baby-daily-v${DATA_STRUCTURE_VERSION}.xlsx`
      downloadBuffer(buffer, filename)
      setMessage('XLSX exported successfully.')
    } finally {
      setIsExporting(false)
    }
  }

  const handleEmail = async () => {
    setMessage('')

    if (!sortedLogs.length) {
      setMessage('No records to email.')
      return
    }

    setIsExporting(true)

    try {
      const buffer = await createWorkbookBuffer(sortedLogs)
      const filename = `baby-daily-v${DATA_STRUCTURE_VERSION}.xlsx`
      const file = downloadBuffer(buffer, filename)

      if (
        navigator.canShare &&
        navigator.canShare({ files: [new File([file], filename, { type: file.type })] })
      ) {
        await navigator.share({
          files: [new File([file], filename, { type: file.type })],
          title: 'Baby Daily Report',
          text: `Please send this report to ${targetEmail}`,
        })
      }

      const subject = encodeURIComponent('Baby Daily XLSX Report')
      const body = encodeURIComponent(
        `Hi,\n\nPlease send the attached report file (${filename}) to this address.\nData structure version: ${DATA_STRUCTURE_VERSION}.`,
      )
      window.open(`mailto:${targetEmail}?subject=${subject}&body=${body}`, '_blank')
      setMessage(`Draft email opened for ${targetEmail}. Attach the downloaded XLSX file if needed.`)
    } finally {
      setIsExporting(false)
    }
  }

  return (
    <main className="min-h-screen bg-slate-100 p-4 text-slate-900 sm:p-8">
      <div className="mx-auto flex w-full max-w-4xl flex-col gap-6">
        <header className="rounded-2xl bg-white p-6 shadow-sm">
          <h1 className="text-2xl font-bold">Baby Daily 2026</h1>
          <p className="mt-1 text-sm text-slate-600">
            Firebase auth + Firestore logbook with XLSX export and versioned data structure.
          </p>
          <p className="mt-1 text-sm text-slate-600">
            Current data structure version: <strong>{DATA_STRUCTURE_VERSION}</strong>
          </p>
        </header>

        {!hasFirebaseConfig && (
          <section className="rounded-2xl border border-amber-300 bg-amber-50 p-4 text-sm text-amber-900">
            Missing Firebase config. Please set VITE_FIREBASE_* values in your environment.
          </section>
        )}

        <section className="rounded-2xl bg-white p-6 shadow-sm">
          {!user ? (
            <button
              type="button"
              onClick={handleLogin}
              disabled={!hasFirebaseConfig}
              className="rounded-lg bg-slate-900 px-4 py-2 font-medium text-white disabled:cursor-not-allowed disabled:opacity-50"
            >
              Sign in with Google
            </button>
          ) : (
            <div className="flex flex-wrap items-center justify-between gap-3">
              <p className="text-sm">
                Signed in as <strong>{user.email}</strong>
              </p>
              <button
                type="button"
                onClick={handleLogout}
                className="rounded-lg border border-slate-300 px-4 py-2 text-sm"
              >
                Sign out
              </button>
            </div>
          )}
        </section>

        <section className="rounded-2xl bg-white p-6 shadow-sm">
          <h2 className="mb-4 text-lg font-semibold">Daily log entry</h2>
          <form className="grid gap-3 sm:grid-cols-2" onSubmit={handleSave}>
            <label className="flex flex-col gap-1 text-sm">
              Date
              <input
                required
                type="date"
                value={form.date}
                onChange={(event) => setForm({ ...form, date: event.target.value })}
                className="rounded-lg border border-slate-300 px-3 py-2"
              />
            </label>
            <label className="flex flex-col gap-1 text-sm">
              Feeding (ml)
              <input
                required
                min="0"
                type="number"
                value={form.feedingMl}
                onChange={(event) => setForm({ ...form, feedingMl: event.target.value })}
                className="rounded-lg border border-slate-300 px-3 py-2"
              />
            </label>
            <label className="flex flex-col gap-1 text-sm">
              Sleep (hours)
              <input
                required
                min="0"
                step="0.5"
                type="number"
                value={form.sleepHours}
                onChange={(event) => setForm({ ...form, sleepHours: event.target.value })}
                className="rounded-lg border border-slate-300 px-3 py-2"
              />
            </label>
            <label className="flex flex-col gap-1 text-sm">
              Diaper count
              <input
                required
                min="0"
                type="number"
                value={form.diaperCount}
                onChange={(event) => setForm({ ...form, diaperCount: event.target.value })}
                className="rounded-lg border border-slate-300 px-3 py-2"
              />
            </label>
            <label className="flex flex-col gap-1 text-sm sm:col-span-2">
              Notes
              <textarea
                value={form.notes}
                onChange={(event) => setForm({ ...form, notes: event.target.value })}
                className="min-h-24 rounded-lg border border-slate-300 px-3 py-2"
              />
            </label>
            <div className="sm:col-span-2">
              <button
                type="submit"
                disabled={!user || isSaving || !hasFirebaseConfig}
                className="rounded-lg bg-sky-600 px-4 py-2 font-medium text-white disabled:cursor-not-allowed disabled:opacity-50"
              >
                {isSaving ? 'Saving...' : 'Save entry'}
              </button>
            </div>
          </form>
        </section>

        <section className="rounded-2xl bg-white p-6 shadow-sm">
          <div className="mb-4 flex flex-wrap items-center justify-between gap-2">
            <h2 className="text-lg font-semibold">Saved entries ({logs.length})</h2>
            <div className="flex flex-wrap gap-2">
              <button
                type="button"
                onClick={handleDownload}
                disabled={isExporting}
                className="rounded-lg border border-slate-300 px-3 py-2 text-sm disabled:cursor-not-allowed disabled:opacity-50"
              >
                Export XLSX
              </button>
              <button
                type="button"
                onClick={handleEmail}
                disabled={isExporting}
                className="rounded-lg bg-emerald-600 px-3 py-2 text-sm font-medium text-white disabled:cursor-not-allowed disabled:opacity-50"
              >
                Email XLSX to {targetEmail}
              </button>
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full min-w-[640px] border-collapse text-left text-sm">
              <thead>
                <tr className="border-b border-slate-200 text-slate-500">
                  <th className="py-2">Date</th>
                  <th className="py-2">Feeding</th>
                  <th className="py-2">Sleep</th>
                  <th className="py-2">Diaper</th>
                  <th className="py-2">Notes</th>
                  <th className="py-2">v</th>
                </tr>
              </thead>
              <tbody>
                {logs.map((entry) => (
                  <tr className="border-b border-slate-100" key={entry.id}>
                    <td className="py-2">{entry.date}</td>
                    <td className="py-2">{entry.feedingMl} ml</td>
                    <td className="py-2">{entry.sleepHours} h</td>
                    <td className="py-2">{entry.diaperCount}</td>
                    <td className="py-2">{entry.notes || '-'}</td>
                    <td className="py-2">{entry.schemaVersion || 1}</td>
                  </tr>
                ))}
                {!logs.length && (
                  <tr>
                    <td className="py-4 text-slate-500" colSpan={6}>
                      No entries yet.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </section>

        {message && <p className="text-sm text-slate-700">{message}</p>}
      </div>
    </main>
  )
}

export default App
