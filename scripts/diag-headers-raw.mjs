import { readFileSync } from "node:fs"
import * as XLSX from "xlsx"

// The real pipeline pins its header to matrix row index 1 (Excel row 2) for
// input files (INPUT_FILE_HEADER_ROW). Row 0 is Qualtrics' internal variable
// name row (e.g. "Q128"); row 1 is the human-readable question text actually
// used as the output column name.
function headersOf(path) {
  const wb = XLSX.read(readFileSync(path))
  const out = {}
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null })
    out[name] = rows[1] || []
  }
  return out
}

function findMatches(headersBySheet, needleRegex) {
  const hits = []
  for (const [sheet, headers] of Object.entries(headersBySheet)) {
    headers.forEach((h, i) => {
      if (h != null && needleRegex.test(String(h))) hits.push({ sheet, col: i, header: h })
    })
  }
  return hits
}

const oldSurvey = headersOf("data/Old-Survey_052826-597993.xlsx")
const newSurveyOld = headersOf("data/New-Survey_060226-1a7f5b.xlsx")
const newSurveyUpdated = headersOf("data/New-Survey_070126-Copy-34c6c1.xlsx")

console.log("=== 'Balanced, energizing work life' across all three files ===")
console.log("Old Survey:", JSON.stringify(findMatches(oldSurvey, /Balanced.*energizing/i), null, 2))
console.log("New Survey (previous):", JSON.stringify(findMatches(newSurveyOld, /Balanced.*energizing/i), null, 2))
console.log("New Survey (updated):", JSON.stringify(findMatches(newSurveyUpdated, /Balanced.*energizing/i), null, 2))

console.log("\n=== 'learn about your new job' across all three files ===")
console.log("Old Survey:", JSON.stringify(findMatches(oldSurvey, /learn about your new job/i), null, 2))
console.log("New Survey (previous):", JSON.stringify(findMatches(newSurveyOld, /learn about your new job/i), null, 2))
console.log("New Survey (updated):", JSON.stringify(findMatches(newSurveyUpdated, /learn about your new job/i), null, 2))

console.log("\n=== Q128 headers, previous vs updated New Survey ===")
console.log("previous:", JSON.stringify(findMatches(newSurveyOld, /^Q128/i), null, 2))
console.log("updated:", JSON.stringify(findMatches(newSurveyUpdated, /^Q128/i), null, 2))

// Sample values under the new "Selected Choice" column in the updated file.
for (const [sheet] of Object.entries(newSurveyUpdated)) {
  const wb = XLSX.read(readFileSync("data/New-Survey_070126-Copy-34c6c1.xlsx"))
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { defval: null })
  const key = Object.keys(rows[0] || {}).find((k) => /learn about your new job/i.test(k))
  if (key) {
    const vals = rows.map((r) => r[key]).filter((v) => v != null && String(v).trim() !== "")
    console.log(`\nSample non-blank values for "${key}" in sheet "${sheet}" (${vals.length} total):`)
    console.log(JSON.stringify([...new Set(vals)].slice(0, 15), null, 2))
  }
}
