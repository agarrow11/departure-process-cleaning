import fs from "node:fs"
import ExcelJS from "exceljs"
import {
  analyzeDuplicateValues,
  deduplicateExactRows,
  exactRowSignature,
  rowsToCSV,
  rowsToXLSX,
  runPipeline,
  stripPIIColumns,
  validateNoExactDuplicates,
} from "../lib/pipeline.ts"

const ab = (path) => { const b = fs.readFileSync(path); return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) }
const files = {
  oldSurvey: ab("data/Old-Survey_052826-597993.xlsx"),
  newSurvey: ab("data/New-Survey_060226-1a7f5b.xlsx"),
  exitInterview: ab("data/Exit-Interview_052826-623ea3.xlsx"),
  beyondBain: ab("data/Beyond-Bain-Extract_060226-9e98f9.xlsx"),
  deptHierarchy: ab("data/dept_hierarchy_20260601-aa7426.xlsx"),
  geoHierarchy: ab("data/geographic_hierarchy_20260501-0691f7.xlsx"),
  pdGrade: fs.readFileSync("data/PD-Grade-Mapping_20260403.csv", "utf8"),
}
const assert = (condition, message) => { if (!condition) throw new Error(message) }
const result = await runPipeline(files)
const redacted = stripPIIColumns(result.rows)
const fullColumns = Object.keys(result.rows[0])
const redactedColumns = Object.keys(redacted[0])
assert(fullColumns.length === 183, `full schema changed: ${fullColumns.length}`)
assert(redactedColumns.length === 175, `no-PII schema changed: ${redactedColumns.length}`)
assert(new Set(result.rows.map((row) => exactRowSignature(row, fullColumns))).size === result.rows.length, "exact duplicate survived")

const response = analyzeDuplicateValues(result.rows, "Response ID")
const ecode = analyzeDuplicateValues(result.rows, "Ecode")
const AMBER = "FFF4B183"
async function verifyWorkbook(rows, expectedColumns, label) {
  const buffer = await rowsToXLSX(rows)
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buffer)
  const ws = wb.worksheets[0]
  assert(ws.columnCount === expectedColumns.length, `${label}: column count ${ws.columnCount}`)
  const headers = expectedColumns.map((_, index) => String(ws.getRow(1).getCell(index + 1).value ?? ""))
  const normalizeHeader = (value) => value.replace(/\r\n/g, "\n")
  const firstHeaderMismatch = headers.findIndex((header, index) => normalizeHeader(header) !== normalizeHeader(expectedColumns[index]))
  assert(firstHeaderMismatch === -1, `${label}: header order changed at ${firstHeaderMismatch + 1}: ${JSON.stringify(headers[firstHeaderMismatch])} vs ${JSON.stringify(expectedColumns[firstHeaderMismatch])}`)
  for (const [column, duplicates] of [["Response ID", response.values], ["Ecode", ecode.values]]) {
    const columnIndex = headers.indexOf(column) + 1
    assert(columnIndex > 0, `${label}: missing ${column}`)
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
      const value = String(rows[rowIndex][column] ?? "").trim()
      const expectedAmber = Boolean(value && duplicates.has(value))
      const actualAmber = ws.getRow(rowIndex + 2).getCell(columnIndex).fill?.fgColor?.argb === AMBER
      assert(actualAmber === expectedAmber, `${label}: ${column} fill mismatch at data row ${rowIndex + 1}`)
    }
  }
}
await verifyWorkbook(result.rows, fullColumns, "full XLSX")
await verifyWorkbook(redacted, redactedColumns, "no-PII XLSX")
const csvRows = rowsToCSV(redacted).trimEnd().split(/\r?\n/)
assert(csvRows.length === redacted.length + 1, `CSV row count mismatch: ${csvRows.length}`)

const planted = [{ Ecode: "A", "Response ID": "R1", x: 1 }, { Ecode: "A", "Response ID": "R1", x: 1 }]
const plantedDedup = deduplicateExactRows(planted, ["Ecode", "Response ID", "x"])
assert(plantedDedup.rows.length === 1 && plantedDedup.removed === 1, "planted exact duplicate was not removed")
const nonIdentical = [{ Ecode: "12345", "Response ID": "R1", x: 1 }, { Ecode: "12345", "Response ID": "R1", x: 2 }]
assert(analyzeDuplicateValues(nonIdentical, "Response ID").affectedRows === 2, "Response ID duplicate review missed planted rows")
assert(analyzeDuplicateValues(nonIdentical, "Ecode").affectedRows === 2, "Ecode duplicate review missed planted rows")
let mustFailCaught = false
try { validateNoExactDuplicates(planted, ["Ecode", "Response ID", "x"]) } catch { mustFailCaught = true }
assert(mustFailCaught, "must-fail control did not reject an exact duplicate")

console.log(JSON.stringify({
  rowsAfterDeduplication: result.rows.length,
  exactRowsRemoved: result.stats.exactDuplicateRowsRemoved,
  duplicateResponseIdGroups: response.groups,
  duplicateResponseIdRows: response.affectedRows,
  duplicateEcodeGroups: ecode.groups,
  duplicateEcodeRows: ecode.affectedRows,
  fullColumns: fullColumns.length,
  noPiiColumns: redactedColumns.length,
  xlsxHighlightsVerified: true,
  csvRowsVerified: true,
  plantedExactDuplicateRemoved: true,
  plantedNonIdenticalDuplicatesRetained: true,
  mustFailCaught,
}, null, 2))
