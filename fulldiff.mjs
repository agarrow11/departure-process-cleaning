import { readFile } from "node:fs/promises"

const PORT = process.env.DEV_PORT || "3000"
const BASE = `http://localhost:${PORT}`

// Proper record-aware CSV parser (handles quoted fields with embedded commas + newlines)
function parseCsv(text) {
  const records = []
  let row = []
  let cur = ""
  let inQ = false
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (inQ) {
      if (ch === '"') {
        if (text[i + 1] === '"') { cur += '"'; i++ } else inQ = false
      } else cur += ch
    } else {
      if (ch === '"') inQ = true
      else if (ch === ",") { row.push(cur); cur = "" }
      else if (ch === "\n") { row.push(cur); records.push(row); row = []; cur = "" }
      else if (ch === "\r") { /* skip */ }
      else cur += ch
    }
  }
  if (cur !== "" || row.length > 0) { row.push(cur); records.push(row) }
  return records
}

const files = {
  oldSurvey: "data/Old-Survey_052826-597993.xlsx",
  newSurvey: "data/New-Survey_060226-1a7f5b.xlsx",
  exitInterview: "data/Exit-Interview_052826-623ea3.xlsx",
  beyondBain: "data/Beyond-Bain-Extract_060226-9e98f9.xlsx",
  deptHierarchy: "data/dept_hierarchy_20260601-aa7426.xlsx",
  geoHierarchy: "data/geographic_hierarchy_20260501-0691f7.xlsx",
  pdGrade: "data/PD-Grade-Mapping_20260403.csv",
}

const fd = new FormData()
for (const [field, path] of Object.entries(files)) {
  const buf = await readFile(path)
  fd.append(field, new Blob([buf]), path.split("/").pop())
}

const res = await fetch(`${BASE}/api/pipeline`, { method: "POST", body: fd })
const json = await res.json()
if (!json.success) { console.log("PIPELINE ERROR:", json.error); process.exit(1) }

const producedCsv = Buffer.from(json.csv, "base64").toString("utf8")
const prod = parseCsv(producedCsv)
const refRecs = parseCsv(await readFile("data/reference_output.csv", "utf8"))

const refHeader = refRecs[0]
const prodHeader = prod[0]

console.log("=== ROW COUNTS (data rows) ===")
console.log("reference:", refRecs.length - 1, "| produced:", prod.length - 1)
console.log("\n=== COLUMN COUNTS ===")
console.log("reference:", refHeader.length, "| produced:", prodHeader.length)

const refSet = new Set(refHeader)
const prodSet = new Set(prodHeader)
console.log("\n=== IN REFERENCE, MISSING FROM PRODUCED ===")
refHeader.filter((c) => !prodSet.has(c)).forEach((c) => console.log("  -", JSON.stringify(c)))
console.log("\n=== EXTRA IN PRODUCED, NOT IN REFERENCE ===")
prodHeader.filter((c) => !refSet.has(c)).forEach((c) => console.log("  +", JSON.stringify(c)))

// Column ORDER comparison (only meaningful when sets match)
console.log("\n=== COLUMN ORDER ===")
let orderOk = refHeader.length === prodHeader.length
let firstOrderDiff = -1
for (let i = 0; i < Math.min(refHeader.length, prodHeader.length); i++) {
  if (refHeader[i] !== prodHeader[i]) { orderOk = false; if (firstOrderDiff < 0) firstOrderDiff = i; }
}
console.log(orderOk ? "  IDENTICAL ORDER" : `  first order diff at index ${firstOrderDiff}: ref=${JSON.stringify(refHeader[firstOrderDiff])} prod=${JSON.stringify(prodHeader[firstOrderDiff])}`)

// VALUE comparison: build maps keyed by Ecode, compare shared columns
const ecodeIdxRef = refHeader.indexOf("Ecode")
const ecodeIdxProd = prodHeader.indexOf("Ecode")
const sharedCols = refHeader.filter((c) => prodSet.has(c))
function indexByEcode(recs, header, ecodeIdx) {
  const m = new Map()
  for (let i = 1; i < recs.length; i++) {
    const k = recs[i][ecodeIdx]
    if (k && !m.has(k)) m.set(k, recs[i])
  }
  return m
}
const refByE = indexByEcode(refRecs, refHeader, ecodeIdxRef)
const prodByE = indexByEcode(prod, prodHeader, ecodeIdxProd)
const refColIdx = Object.fromEntries(refHeader.map((c, i) => [c, i]))
const prodColIdx = Object.fromEntries(prodHeader.map((c, i) => [c, i]))
let cellsCompared = 0, cellsDiff = 0
const diffByCol = {}
let checkedEcodes = 0
for (const [ecode, rRow] of refByE) {
  const pRow = prodByE.get(ecode)
  if (!pRow) continue
  checkedEcodes++
  for (const c of sharedCols) {
    if (c === "Ecode") continue
    const rv = (rRow[refColIdx[c]] ?? "").trim()
    const pv = (pRow[prodColIdx[c]] ?? "").trim()
    cellsCompared++
    if (rv !== pv) { cellsDiff++; diffByCol[c] = (diffByCol[c] || 0) + 1 }
  }
}
console.log("\n=== VALUE COMPARISON (matched Ecodes, shared cols) ===")
console.log("  ecodes compared:", checkedEcodes, "| cells compared:", cellsCompared, "| cells differing:", cellsDiff)
const topDiffs = Object.entries(diffByCol).sort((a, b) => b[1] - a[1]).slice(0, 20)
if (topDiffs.length) {
  console.log("  top differing columns:")
  topDiffs.forEach(([c, n]) => console.log(`    ${n.toString().padStart(6)}  ${JSON.stringify(c)}`))
}

// Sample concrete value pairs for the most-differing columns
const sampleCols = ["Start Date", "End Date", "Finished", "Mapped Employee Level", "Recorded Date", "BB_Bain Departure Year (ZID7_163)"]
console.log("\n=== SAMPLE VALUE PAIRS (ref | prod) ===")
for (const c of sampleCols) {
  if (!prodSet.has(c) || !refSet.has(c)) { console.log(`\n  [${c}] (not in both)`); continue }
  console.log(`\n  [${c}]`)
  let shown = 0
  for (const [ecode, rRow] of refByE) {
    const pRow = prodByE.get(ecode); if (!pRow) continue
    const rv = (rRow[refColIdx[c]] ?? "").trim()
    const pv = (pRow[prodColIdx[c]] ?? "").trim()
    if (rv !== pv) { console.log(`    ${ecode}: ${JSON.stringify(rv)} | ${JSON.stringify(pv)}`); if (++shown >= 4) break }
  }
}
