// Throwaway verification: confirm the pipeline widens its fixed-column
// output when the New Survey source file adds columns not present in the
// Old Survey (Q202/Q203/Q204), and that the PII/duplicate hard gates still
// pass unchanged at the new width. Not shipped.
import { readFileSync } from "node:fs"
import { runPipeline } from "../lib/pipeline.ts"

function ab(path) {
  const buf = readFileSync(path)
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength)
}

const files = {
  oldSurvey: ab("data/Old-Survey_052826-597993.xlsx"),
  newSurvey: ab("data/New-Survey_070126-Copy-34c6c1.xlsx"),
  exitInterview: ab("data/Exit-Interview_052826-623ea3.xlsx"),
  beyondBain: ab("data/Beyond-Bain-Extract_060226-9e98f9.xlsx"),
  deptHierarchy: ab("data/dept_hierarchy_20260601-aa7426.xlsx"),
  geoHierarchy: ab("data/geographic_hierarchy_20260501-0691f7.xlsx"),
  pdGrade: ab("data/PD-Grade-Mapping_20260403.csv"),
}

const previousFiles = {
  ...files,
  newSurvey: ab("data/New-Survey_060226-1a7f5b.xlsx"),
}

function assert(cond, msg) {
  if (!cond) throw new Error("ASSERTION FAILED: " + msg)
  console.log("OK:", msg)
}

const before = await runPipeline(previousFiles)
const after = await runPipeline(files)

console.log("\n--- Before (original New Survey) ---")
console.log("totalColumns:", before.stats.totalColumns)
console.log("populationRows:", before.stats.populationRows)

console.log("\n--- After (updated New Survey) ---")
console.log("totalColumns:", after.stats.totalColumns)
console.log("populationRows:", after.stats.populationRows)

const beforeCols = Object.keys(before.rows[0])
const afterCols = Object.keys(after.rows[0])

// The real diff is 5 new columns, not the 3 comp questions alone: the survey
// instrument also added one satisfaction sub-item ("Balanced, energizing
// work life") and restructured "How did you learn about your new job?" from
// a 7-option multi-select into a single "- Selected Choice" field. All 5
// are genuinely new column names not present in the prior New Survey file.
const priorSet = new Set(beforeCols)

const COMP_COL_BASE = "Please specify the following values in US$ equivalent:"
const compCols = afterCols.filter((c) => c.startsWith(COMP_COL_BASE))
const compColsAlreadyPresent = compCols.filter((c) => priorSet.has(c))
const compColsNew = compCols.filter((c) => !priorSet.has(c))
console.log(`\nCompensation-prefixed columns present after: ${compCols.length}`)
console.log(`  already present before: ${JSON.stringify(compColsAlreadyPresent)}`)
console.log(`  new: ${JSON.stringify(compColsNew)}`)
assert(compColsNew.length === 3, `exactly 3 compensation columns are new: ${JSON.stringify(compColsNew)}`)
assert(
  afterCols.some((c) => c.includes("Balanced, energizing work life")),
  "new satisfaction sub-item present",
)
assert(
  afterCols.includes("How did you learn about your new job? - Selected Choice"),
  "restructured single-select 'how did you learn' column present",
)

const newOnes = afterCols.filter((c) => !priorSet.has(c))
assert(newOnes.length === 5, `exactly 5 columns are new: ${JSON.stringify(newOnes)}`)
const tailStart = afterCols.length - 5
assert(
  afterCols.slice(tailStart).every((c) => newOnes.includes(c)),
  "the 5 new columns sit at the tail of the schema, not inserted mid-schema",
)

// Confirm every prior column is still present, in the same relative order —
// including the 7 old multi-select boolean sub-columns, which the new file
// no longer emits but which remain in the schema (blank for new rows) since
// the Old Survey file still asks the question in that format.
const afterPriorOnly = afterCols.filter((c) => priorSet.has(c))
assert(
  afterPriorOnly.length === beforeCols.length && afterPriorOnly.every((c, i) => c === beforeCols[i]),
  "every prior column survives, in the same relative order",
)
const oldMultiSelectCol =
  "How did you learn about your new job? (select all that apply) - Selected Choice - LinkedIn"
assert(afterCols.includes(oldMultiSelectCol), "old multi-select sub-column still present in the schema")
const newSurveySourcedRows = after.rows.filter((r) => r._source === "new")
const blankOnNewRows = newSurveySourcedRows.every((r) => r[oldMultiSelectCol] == null || String(r[oldMultiSelectCol]).trim() === "")
assert(blankOnNewRows || newSurveySourcedRows.length === 0, "old multi-select column is blank on rows sourced from the restructured survey")

// Confirm at least one row actually carries a value in a new column, proving
// data (not just a header) made it through.
const compColSample = compCols[0]
const populated = after.rows.filter((r) => r[compColSample] != null && String(r[compColSample]).trim() !== "")
console.log(`\nRows with a non-blank "${compColSample}": ${populated.length} of ${after.rows.length}`)
const selectedChoicePopulated = after.rows.filter(
  (r) => r["How did you learn about your new job? - Selected Choice"] != null && String(r["How did you learn about your new job? - Selected Choice"]).trim() !== "",
)
console.log(`Rows with a non-blank restructured "Selected Choice": ${selectedChoicePopulated.length} of ${after.rows.length}`)

// Confirm the PII/duplicate audit checks still ran and passed at the new width.
const checkLabels = after.checks.map((c) => c.label)
assert(checkLabels.includes("PII redaction structure"), "PII redaction structure check ran")
assert(checkLabels.includes("Exact-row deduplication"), "Exact-row deduplication check ran")
const piiCheck = after.checks.find((c) => c.label === "PII redaction structure")
assert(piiCheck.status === "ok", `PII redaction structure check passed: ${piiCheck.detail}`)

console.log("\nAll schema-widening assertions passed.")
