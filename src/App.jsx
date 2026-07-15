import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import accountsFileUrl from './assets/accounts.xlsx?url'

const BASE_URL = 'https://api.h5r1xc.xyz/xxapi/buyitoken/waitpayerpaymentslip'
const REQUEST_LIMIT = 200
const TOKEN_STORAGE_KEY = 'tivrapay-indiatoken'

const OUTPUT_HEADERS = [
  'rptNo',
  'orderNo',
  'amount',
  'acctNo',
  'acctCode',
  'acctName',
  'matchAccNo',
  'matchIfsc',
  'matchName',
]

const OUTPUT_LABELS = {
  rptNo: 'Rpt No',
  orderNo: 'Order No',
  amount: 'Amount',
  acctNo: 'Account No',
  acctCode: 'Acct Code',
  acctName: 'Acct Name',
  matchAccNo: 'Match Acc No',
  matchIfsc: 'Match IFSC',
  matchName: 'Match Name',
}

function extractLastFour(value) {
  const digits = String(value ?? '').replace(/\D/g, '')
  if (!digits) {
    return ''
  }
  return digits.slice(-4)
}

function sleep(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms)
  })
}

async function fetchWithRetry(url, options, maxRetries = 3) {
  for (let attempt = 1; attempt <= maxRetries; attempt += 1) {
    const response = await fetch(url, options)
    const text = await response.text()

    if (response.status === 200) {
      return JSON.parse(text)
    }

    if ([520, 522, 524, 503].includes(response.status) && attempt < maxRetries) {
      await sleep(3000)
      continue
    }

    throw new Error(`API Error ${response.status}: ${text}`)
  }

  throw new Error('Max retry reached')
}

async function fetchTivraPayData(indiaToken) {
  const options = {
    method: 'GET',
headers: {
  Accept: 'application/json',
  indiatoken: indiaToken,
  'x-rs-cfg-tivpayreqgate': 'A7K9X2M8Q4P1Z'
},
  }

  let page = 1
  let total = 0
  let allList = []

  while (true) {
    const params = new URLSearchParams({
      page: String(page),
      limit: String(REQUEST_LIMIT),
      if_asc: 'false',
      min_amount: '5000',
      max_amount: '100000',
      method: '1',
      date_asc: '1',
    })

    const url = `${BASE_URL}?${params.toString()}`
    const json = await fetchWithRetry(url, options)

    if (!json?.data?.list) {
      throw new Error(`data.list not found. Full response: ${JSON.stringify(json)}`)
    }

    total = json.data.total || 0
    const list = json.data.list
    allList = allList.concat(list)

    if (allList.length >= total || list.length === 0) {
      break
    }

    page += 1
  }

  return { total, allList }
}

function normalizeHeaderName(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '')
}

async function loadAccountsFromXlsx() {
  const response = await fetch(accountsFileUrl)
  if (!response.ok) {
    throw new Error(`Failed to load accounts file: ${response.status}`)
  }

  const arrayBuffer = await response.arrayBuffer()
  const workbook = XLSX.read(arrayBuffer, { type: 'array' })
  const sheetName = workbook.SheetNames[0]
  if (!sheetName) {
    throw new Error('No sheet found in accounts.xlsx')
  }

  const sheet = workbook.Sheets[sheetName]
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false })
  if (!rows.length) {
    throw new Error('accounts.xlsx is empty')
  }

  const headerRow = rows[0].map((cell) => normalizeHeaderName(cell))
  const nameIndex = headerRow.findIndex((cell) => cell === 'name')
  const ifscIndex = headerRow.findIndex((cell) => cell === 'ifscbank')
  const accountIndex = headerRow.findIndex((cell) => cell === 'accountnumber')

  if (nameIndex === -1 || ifscIndex === -1 || accountIndex === -1) {
    throw new Error('accounts.xlsx must contain: Name, IFSC / Bank, Account Number')
  }

  const accountMap = new Map()

  rows.slice(1).forEach((row) => {
    const accountNumber = extractLastFour(row[accountIndex])
    if (accountNumber.length !== 4) {
      return
    }

    const entry = {
      matchAccNo: accountNumber,
      matchIfsc: String(row[ifscIndex] || ''),
      matchName: String(row[nameIndex] || ''),
    }

    if (!accountMap.has(accountNumber)) {
      accountMap.set(accountNumber, [])
    }
    accountMap.get(accountNumber).push(entry)
  })

  return accountMap
}

function App() {
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [result, setResult] = useState(null)
  const [token, setToken] = useState(() => localStorage.getItem(TOKEN_STORAGE_KEY) || '')
  const [isTokenDialogOpen, setIsTokenDialogOpen] = useState(false)
  const [draftToken, setDraftToken] = useState(token)

  const summary = useMemo(() => {
    if (!result) {
      return null
    }
    return {
      referenceAccounts: result.accountMapSize,
      apiTotal: result.apiTotal,
      apiFetched: result.apiRows.length,
      matched: result.matchedRows.length,
    }
  }, [result])

  const onMatchClick = async () => {
    const cleanToken = token.trim()
    if (!cleanToken) {
      setError('Please enter indiatoken.')
      return
    }

    localStorage.setItem(TOKEN_STORAGE_KEY, cleanToken)

    setLoading(true)
    setError('')
    setResult(null)

    try {
      const accountMap = await loadAccountsFromXlsx()
      if (accountMap.size === 0) {
        throw new Error('No valid 4-digit account numbers in accounts.xlsx')
      }

      const { total, allList } = await fetchTivraPayData(cleanToken)

      const normalizedApi = allList.map((item) => {
        const apiLast4 = extractLastFour(item.acctNo)
        const matches = accountMap.get(apiLast4) || []
        const firstMatch = matches[0]

        return {
          rptNo: item.rptNo ?? '',
          orderNo: item.orderNo ?? '',
          amount: item.amount ?? '',
          acctNo: item.acctNo ?? '',
          acctCode: item.acctCode ?? '',
          acctName: item.acctName ?? '',
          matchAccNo: firstMatch?.matchAccNo ?? '',
          matchIfsc: firstMatch?.matchIfsc ?? '',
          matchName: firstMatch?.matchName ?? '',
          isMatched: matches.length > 0,
        }
      })

      const matchedRows = normalizedApi.filter((item) => item.isMatched)

      setResult({
        apiTotal: total,
        apiRows: normalizedApi,
        matchedRows,
        accountMapSize: accountMap.size,
      })
    } catch (err) {
      setError(err?.message || 'Something went wrong while matching data.')
    } finally {
      setLoading(false)
    }
  }

  const openTokenDialog = () => {
    setDraftToken(token)
    setIsTokenDialogOpen(true)
  }

  const closeTokenDialog = () => {
    setIsTokenDialogOpen(false)
  }

  const saveTokenFromDialog = () => {
    const cleanToken = draftToken.trim()
    setToken(cleanToken)
    localStorage.setItem(TOKEN_STORAGE_KEY, cleanToken)
    setIsTokenDialogOpen(false)
  }

  return (
    <main className="mx-auto w-full max-w-7xl p-3 sm:p-5 lg:p-7">
      <div className="mb-2 flex justify-end">
        <button
          type="button"
          className="rounded-xl border border-slate-300 bg-white px-4 py-2 font-semibold text-slate-700 shadow-sm transition hover:-translate-y-0.5 hover:shadow"
          onClick={openTokenDialog}
        >
          Save Token
        </button>
      </div>

      <h1 className="mb-4 text-2xl font-bold tracking-tight text-slate-900 sm:text-3xl">
        Matcher
      </h1>

      <section className="mb-3 flex items-center gap-3">
        <button
          type="button"
          className="rounded-xl bg-blue-600 px-4 py-2 font-semibold text-white shadow-md transition hover:-translate-y-0.5 hover:bg-blue-700 disabled:cursor-not-allowed disabled:opacity-60 disabled:hover:translate-y-0"
          onClick={onMatchClick}
          disabled={loading}
        >
          {loading ? 'Matching...' : 'Fetch & Match'}
        </button>
      </section>

      {error && (
        <p className="mb-4 rounded-xl border border-rose-200 bg-rose-50 px-3 py-2 font-semibold text-rose-700">
          {error}
        </p>
      )}

      {summary && (
        <section className="mb-5 flex flex-wrap gap-2 rounded-xl border border-slate-200 bg-slate-50 p-3">
          <div className="rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm text-slate-700">
            Assets unique last 4: {summary.referenceAccounts}
          </div>
          <div className="rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm text-slate-700">
            API total: {summary.apiTotal}
          </div>
          <div className="rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm text-slate-700">
            API fetched: {summary.apiFetched}
          </div>
          <div className="rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm text-slate-700">
            Matched: {summary.matched}
          </div>
        </section>
      )}

      {result && (
        <>
          <h2 className="mb-3 text-lg font-semibold text-slate-900">Matched Rows</h2>
          <DataTable rows={result.matchedRows} />
        </>
      )}

      {isTokenDialogOpen && (
        <div
          className="fixed inset-0 z-50 grid place-items-center bg-slate-900/50 p-4"
          role="presentation"
          onClick={closeTokenDialog}
        >
          <div
            className="w-full max-w-xl rounded-xl border border-slate-200 bg-white p-4 shadow-2xl"
            role="dialog"
            aria-modal="true"
            onClick={(e) => e.stopPropagation()}
          >
            <h3 className="mb-3 text-lg font-semibold text-slate-900">Save API Token</h3>
            <input
              className="w-full rounded-xl border border-slate-300 px-3 py-2 text-sm text-slate-800 outline-none ring-blue-200 placeholder:text-slate-400 focus:border-blue-500 focus:ring-4"
              type="text"
              placeholder="Enter indiatoken"
              value={draftToken}
              onChange={(event) => setDraftToken(event.target.value)}
            />
            <div className="mt-4 flex justify-end gap-2">
              <button
                type="button"
                className="rounded-xl border border-slate-300 bg-white px-4 py-2 font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50"
                onClick={closeTokenDialog}
              >
                Cancel
              </button>
              <button
                type="button"
                className="rounded-xl bg-blue-600 px-4 py-2 font-semibold text-white shadow-md transition hover:bg-blue-700"
                onClick={saveTokenFromDialog}
              >
                Save
              </button>
            </div>
          </div>
        </div>
      )}
    </main>
  )
}

function DataTable({ rows }) {
  if (rows.length === 0) {
    return (
      <p className="mb-6 rounded-xl border border-dashed border-slate-300 bg-slate-50 p-3 text-slate-600">
        No rows found.
      </p>
    )
  }

  return (
    <div>
      <div className="mb-6 hidden overflow-x-auto rounded-xl border border-slate-200 bg-white shadow-sm md:block">
        <table>
          <thead className="bg-slate-100 text-slate-700">
            <tr>
              {OUTPUT_HEADERS.map((header) => (
                <th className="px-3 py-2 text-left text-xs font-semibold uppercase" key={header}>
                  {header}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((item, idx) => (
              <tr className="border-t border-slate-100 odd:bg-white even:bg-slate-50" key={`${item.orderNo || item.rptNo || 'row'}-${idx}`}>
                {OUTPUT_HEADERS.map((header) => (
                  <td className="max-w-[240px] wrap-break-word px-3 py-2 text-sm text-slate-700" key={`${idx}-${header}`}>
                    {String(item[header] ?? '')}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="mb-6 grid gap-2.5 md:hidden">
        {rows.map((item, idx) => (
          <article
            className="rounded-lg border border-slate-200 bg-white p-2.5 shadow-sm"
            key={`${item.orderNo || item.rptNo || 'card'}-${idx}`}
          >
            <div className="mb-2 flex items-center justify-between rounded-md bg-slate-50 px-2 py-1.5">
              <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                Rpt: {String(item.rptNo ?? '-')}
              </span>
              <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                Order: {String(item.orderNo ?? '-')}
              </span>
            </div>
            {OUTPUT_HEADERS.map((header) => (
              <div
                className="grid grid-cols-[90px,1fr] items-start gap-2 border-b border-slate-100 py-1 last:border-b-0"
                key={`${idx}-card-${header}`}
              >
                <span className="text-[11px] font-medium text-slate-500">
                  {OUTPUT_LABELS[header]}
                </span>
                <span className="wrap-break-word text-sm leading-5 text-slate-800">
                  {String(item[header] ?? '')}
                </span>
              </div>
            ))}
          </article>
        ))}
      </div>
    </div>
  )
}

export default App
