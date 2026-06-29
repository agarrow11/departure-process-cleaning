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
