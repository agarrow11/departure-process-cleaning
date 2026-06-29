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

// Column ORDER check (only over the common columns, in sequence)
console.log("\n=== COLUMN ORDER (first mismatch over common cols) ===")
let orderOk = true
const minLen = Math.min(refHeader.length, prodHeader.length)
for (let i = 0; i < minLen; i++) {
  if (refHeader[i] !== prodHeader[i]) {
    console.log(`  position ${i}: ref=${JSON.stringify(refHeader[i])} prod=${JSON.stringify(prodHeader[i])}`)
    orderOk = false
    break
  }
}
if (orderOk) console.log("  exact order match over", minLen, "columns")

// Value-level comparison on matched Ecodes over common columns.
const norm = (v) => (v == null ? "" : String(v).trim())
const toMap = (recs, header) => {
  const idx = header.indexOf("Ecode")
  const m = new Map()
  for (let r = 1; r < recs.length; r++) {
    const obj = {}
    for (let c = 0; c < header.length; c++) obj[header[c]] = recs[r][c]
    const key = norm(recs[r][idx])
    if (key && !m.has(key)) m.set(key, obj)
  }
  return m
}
const refMap = toMap(refRecs, refHeader)
const prodMap = toMap(prod, prodHeader)
const commonCols = refHeader.filter((c) => prodSet.has(c))
let comparedRows = 0
const colMismatch = new Map()
for (const [ecode, robj] of refMap) {
  const pobj = prodMap.get(ecode)
  if (!pobj) continue
  comparedRows++
  for (const c of commonCols) {
    if (norm(robj[c]) !== norm(pobj[c])) colMismatch.set(c, (colMismatch.get(c) || 0) + 1)
  }
}
console.log("\n=== VALUE COMPARISON (matched Ecodes:", comparedRows, ") ===")
const sorted = [...colMismatch.entries()].sort((a, b) => b[1] - a[1])
if (sorted.length === 0) console.log("  all common-column values match on matched Ecodes")
else { console.log("  columns with differing values (count):"); sorted.slice(0, 25).forEach(([c, n]) => console.log(`    ${n}  ${JSON.stringify(c)}`)) }

// Show concrete examples for the top mismatching columns
console.log("\n=== EXAMPLES (top differing columns) ===")
for (const [col] of sorted.slice(0, 8)) {
  let shown = 0
  for (const [ecode, robj] of refMap) {
    const pobj = prodMap.get(ecode)
    if (pobj && norm(robj[col]) !== norm(pobj[col])) {
      console.log(`  ${JSON.stringify(col)} Ecode ${ecode}: ref=${JSON.stringify(robj[col])} prod=${JSON.stringify(pobj[col])}`)
      if (++shown >= 2) break
    }
  }
}
