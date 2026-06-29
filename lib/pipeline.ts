/**
 * Departure Data Pipeline — TypeScript/Node.js port
 * ===================================================
 * Direct port of departure_data_pipeline.py
 * Uses SheetJS (xlsx) for Excel/CSV processing.
 *
 * Join sequence:
 *   1. Stack Old Survey + New Survey
 *   2. Stack Exit Interview files
 *   3. Apply 3 mapping lookups (Dept, Geo, PD Grade)
 *   4. Outer join Survey + EI on Ecode
 *   5. Left join Beyond Bain on Ecode
 *   6. Output cleaned CSV + XLSX
 *
 * NOTE: Sheet names and header-row positions are detected automatically by
 * scanning each workbook for the anchor columns each file is expected to
 * contain. Month-to-month changes to date-stamped tab names
 * (e.g. "Dept Hierarchy_260206") or shifted header rows therefore no longer
 * break the pipeline — as long as the underlying column names stay the same.
 */

import * as XLSX from "xlsx"
import ExcelJS from "exceljs"

// =============================================================================
// TYPES
// =============================================================================

export type Row = Record<string, string | number | null | undefined>

export interface PipelineFiles {
  oldSurvey: ArrayBuffer
  newSurvey: ArrayBuffer
  exitInterview: ArrayBuffer
  beyondBain: ArrayBuffer
  deptHierarchy: ArrayBuffer
  geoHierarchy: ArrayBuffer
  pdGrade: ArrayBuffer | string // CSV can come as string or buffer
}

export interface PipelineResult {
  rows: Row[]
  stats: {
    oldSurveyRows: number
    newSurveyRows: number
    combinedSurveyRows: number
    exitInterviewRows: number
    populationRows: number
    bothSourcesCount: number
    surveyOnlyCount: number
    eiOnlyCount: number
    beyondBainMatched: number
    totalColumns: number
  }
  warnings: string[]
  audit: AuditReport
}

/** Points at one specific impacted record so the analyst can locate it. */
export interface AuditRowRef {
  /** The employee code, or "" if the row has no Ecode. */
  ecode: string
  /** Which source file the row came from, e.g. "Survey 1", "Exit Interview". */
  source: string
  /** 1-based row number within that source file's data (null when not applicable). */
  rowNumber: number | null
}

/** One offending value found by an audit check, e.g. an unmatched code. */
export interface AuditIssue {
  /** The offending value (e.g. an unmatched department code, or a blank-key placeholder). */
  value: string
  /** How many rows were affected by this value. */
  rowCount: number
  /** The exact rows impacted — row number + source file + Ecode. */
  rows: AuditRowRef[]
}

/** A single data-quality / integrity check performed while running the pipeline. */
export interface AuditCheck {
  /** Short human-readable name, e.g. "Department mapping". */
  label: string
  /** ok = clean, warning = ran but some rows need attention, error = check failed. */
  status: "ok" | "warning" | "error"
  /** Plain-English explanation of the result. */
  detail: string
  /** Optional concrete examples (e.g. the offending codes) to make fixing easy. */
  samples?: string[]
  /** Header for the "value" column when drilling into the full issue list. */
  issueLabel?: string
  /** The COMPLETE list of offending values, for drill-down. Not truncated. */
  issues?: AuditIssue[]
}

export interface AuditReport {
  checks: AuditCheck[]
  okCount: number
  warningCount: number
  errorCount: number
}

// =============================================================================
// COLUMN CONFIG — Update here if source file column names change
// =============================================================================

const COLS = {
  // Shared join key
  ECODE: "Ecode",

  // Exit Interview
  EI_CLEANED_ECODE: "Cleaned Ecode",

  // Beyond Bain
  BB_ECODE: "E-code (ZID11_48)",
  BB_FUNCTION: "Function (ZID4_117)",
  BB_MAPPED_FUNCTION: "Mapped Role Function", // created by Logic #1

  // Dept Hierarchy mapping
  DEPT_CODE: "dept_code", // lookup key
  DEPT_DEPT: "Department", // → Department in input files
  DEPT_SUBFUNC: "Sub-Function", // → Sub-Function in input files
  DEPT_FUNC: "People Function", // → Function in input files
  DEPT_INPUT_KEY: "Department_ID", // key field in input files

  // Geo Hierarchy mapping
  GEO_CODE: "office_code", // lookup key
  GEO_OFFICE: "people_office_display", // → ExternalDataReference
  GEO_CLUSTER: "people_cluster_display", // → Cluster
  GEO_REGION: "people_region_display", // → Region
  GEO_INPUT_KEY: "Office_Code", // key field in input files

  // PD Grade mapping
  PD_CODE: "PD Grade", // lookup key (same name in input files)
  PD_SENIORITY: "Seniority", // → Mapped Seniority
  PD_LEVEL: "Employee Level", // → Mapped Employee Level
  PD_POSITION: "CPH_3", // → Mapped Position
}

/**
 * Anchor columns used to auto-locate the correct sheet AND header row in each
 * workbook. Every `required` column MUST appear in a sheet's header row for that
 * sheet to be accepted; `optional` columns raise the confidence score when
 * choosing between candidate rows/sheets.
 */
const SHEET_ANCHORS = {
  survey: {
    required: [COLS.ECODE],
    optional: [COLS.DEPT_INPUT_KEY, COLS.GEO_INPUT_KEY, COLS.PD_CODE],
  },
  ei: {
    required: [COLS.EI_CLEANED_ECODE],
    optional: [COLS.DEPT_INPUT_KEY, COLS.GEO_INPUT_KEY, COLS.PD_CODE],
  },
  dept: {
    required: [COLS.DEPT_CODE],
    optional: [COLS.DEPT_DEPT, COLS.DEPT_SUBFUNC, COLS.DEPT_FUNC],
  },
  geo: {
    required: [COLS.GEO_CODE],
    optional: [COLS.GEO_OFFICE, COLS.GEO_CLUSTER, COLS.GEO_REGION],
  },
  bb: {
    required: [COLS.BB_ECODE],
    optional: [COLS.BB_FUNCTION],
  },
  pd: {
    required: [COLS.PD_CODE],
    optional: [COLS.PD_SENIORITY, COLS.PD_LEVEL, COLS.PD_POSITION],
  },
}

// How many rows to scan from the top of a sheet when hunting for the header row.
const MAX_HEADER_SCAN = 20

// =============================================================================
// HELPERS
// =============================================================================

function standardizeEcode(val: unknown): string {
  if (val == null) return ""
  return String(val).trim().toUpperCase()
}

function normalizeHeader(v: unknown): string {
  return String(v ?? "").trim().toLowerCase()
}

function readWorkbook(buffer: ArrayBuffer | string): XLSX.WorkBook {
  return typeof buffer === "string"
    ? XLSX.read(buffer, { type: "string" })
    : XLSX.read(buffer, { type: "array" })
}

/** Convert a worksheet to a raw matrix (array of rows), preserving row indices. */
function sheetToMatrix(ws: XLSX.WorkSheet): unknown[][] {
  return XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: null,
    raw: false,
    blankrows: true,
  }) as unknown[][]
}

/**
 * Scan the first MAX_HEADER_SCAN rows of a matrix and return the row index whose
 * cells best match the expected column names (case-insensitive). Returns row -1
 * if no scanned row matches any expected column.
 */
function detectHeaderRow(matrix: unknown[][], expectedCols: string[]): { row: number; score: number } {
  const expected = expectedCols.map(normalizeHeader)
  let best = { row: -1, score: 0 }
  const limit = Math.min(MAX_HEADER_SCAN, matrix.length)
  for (let i = 0; i < limit; i++) {
    const cells = new Set((matrix[i] ?? []).map(normalizeHeader))
    let score = 0
    for (const e of expected) if (e && cells.has(e)) score++
    if (score > best.score) best = { row: i, score }
  }
  return best
}

/**
 * Locate the correct sheet + header row in a workbook by content, then parse it.
 * Picks the sheet whose header row contains the most `required` anchors (then
 * the highest overall score). Throws a descriptive error if none qualifies.
 */
function locateAndRead(
  buffer: ArrayBuffer | string,
  anchors: { required: string[]; optional: string[] },
  label: string,
  warnings: string[],
  /**
   * When set, the header is NOT content-detected but pinned to this 0-based row
   * index. Used for the survey/EI input files, whose layout is fixed (row 1 is
   * ignored, row 2 = headers). This prevents header-like text on row 1 from
   * being mistaken for the header, which would otherwise pull the row-3 metadata
   * line in as a data record.
   */
  fixedHeaderRow?: number,
): { rows: Row[]; sheetName: string; headerRow: number; matrix: unknown[][] } {
  const wb = readWorkbook(buffer)
  const allAnchors = [...anchors.required, ...anchors.optional]

  let best = { name: "", headerRow: -1, score: -1, requiredHits: -1 }

  for (const name of wb.SheetNames) {
    const matrix = sheetToMatrix(wb.Sheets[name])
    if (matrix.length === 0) continue

    // Header row: pinned (input files) or detected by content (mapping files).
    let row: number
    let score: number
    if (fixedHeaderRow != null) {
      if (matrix.length <= fixedHeaderRow) continue
      row = fixedHeaderRow
      const cells = new Set((matrix[row] ?? []).map(normalizeHeader))
      score = allAnchors.reduce((acc, a) => acc + (a && cells.has(normalizeHeader(a)) ? 1 : 0), 0)
    } else {
      const detected = detectHeaderRow(matrix, allAnchors)
      row = detected.row
      score = detected.score
    }
    if (row < 0) continue

    const headerCells = new Set((matrix[row] ?? []).map(normalizeHeader))
    const requiredHits = anchors.required.filter((k) => headerCells.has(normalizeHeader(k))).length
    if (
      requiredHits > best.requiredHits ||
      (requiredHits === best.requiredHits && score > best.score)
    ) {
      best = { name, headerRow: row, score, requiredHits }
    }
  }

  if (best.headerRow < 0 || best.requiredHits < anchors.required.length) {
    throw new Error(
      `Could not locate the ${label} sheet. Expected a sheet containing column(s) ` +
        `[${anchors.required.join(", ")}]${fixedHeaderRow != null ? ` on row ${fixedHeaderRow + 1}` : ""}. ` +
        `Available sheets: ${wb.SheetNames.join(", ")}.`,
    )
  }

  warnings.push(`${label}: detected sheet "${best.name}", header on row ${best.headerRow + 1}.`)

  const rows: Row[] = XLSX.utils.sheet_to_json(wb.Sheets[best.name], {
    defval: null,
    raw: false,
    range: best.headerRow,
  })
  // Raw matrix (blank rows preserved) so callers can map records to true Excel row numbers.
  const matrix = sheetToMatrix(wb.Sheets[best.name])
  return { rows, sheetName: best.name, headerRow: best.headerRow, matrix }
}

/**
 * Build data records from the raw matrix while preserving each record's TRUE
 * spreadsheet row number (1-based). Per the source-file structure:
 *   • the detected header sits on its row (normally Excel row 2),
 *   • the single row directly beneath it (normally Excel row 3) is always
 *     metadata and is dropped — it is never a data record, and
 *   • real data begins on the next row (normally Excel row 4).
 *
 * This mirrors the reference pandas pipeline (`read_excel(header=1)` then
 * `df.iloc[1:]`), including two behaviors that matter for an exact match:
 *   1. Duplicate header names are de-collided pandas-style — the first keeps its
 *      name, the next becomes "Name.1", then "Name.2", and so on.
 *   2. Interior blank rows are KEPT (pandas does not drop them), so the row
 *      count matches the reference output exactly.
 * Returns the ordered list of (de-collided) column names alongside the rows so
 * the caller can reproduce pandas' column ordering downstream.
 */
function buildPositionedRows(
  matrix: unknown[][],
  headerRow: number,
  label: string,
  warnings: string[],
): { rows: Array<{ row: Row; excelRow: number }>; columns: string[] } {
  const rawHeaders = (matrix[headerRow] ?? []).map((h) => String(h ?? "").trim())

  // pandas-style duplicate-name mangling. Empty header cells are skipped
  // (treated as unnamed and excluded), which the reference output also does.
  const seen = new Map<string, number>()
  const colNames: (string | null)[] = rawHeaders.map((h) => {
    if (!h) return null
    const n = seen.get(h) ?? 0
    seen.set(h, n + 1)
    return n === 0 ? h : `${h}.${n}`
  })
  const columns = colNames.filter((c): c is string => c !== null)

  const metaRowIdx = headerRow + 1 // the "row 3" metadata row — always dropped
  const firstDataIdx = headerRow + 2 // the "row 4" first real data record

  if (matrix.length > metaRowIdx) {
    warnings.push(`${label}: dropped row ${metaRowIdx + 1} (metadata); data linked from row ${firstDataIdx + 1}.`)
  }

  // Trim only fully-blank TRAILING rows (an Excel artifact); interior blanks are
  // preserved to mirror pandas and keep the row count aligned with the reference.
  let lastIdx = matrix.length - 1
  while (lastIdx >= firstDataIdx) {
    const cells = matrix[lastIdx] ?? []
    if (cells.every((c) => c === null || c === undefined || String(c).trim() === "")) lastIdx--
    else break
  }

  const out: Array<{ row: Row; excelRow: number }> = []
  for (let mi = firstDataIdx; mi <= lastIdx; mi++) {
    const cells = matrix[mi] ?? []
    const row: Row = {}
    for (let ci = 0; ci < colNames.length; ci++) {
      const key = colNames[ci]
      if (key === null) continue
      row[key] = cells[ci] ?? null
    }
    out.push({ row, excelRow: mi + 1 }) // mi is 0-based; +1 = true 1-based Excel row
  }
  return { rows: out, columns }
}

/**
 * Fixed 0-based header row index for the survey/EI input files. Their layout is
 * deterministic: row 1 (index 0) is ignored, row 2 (index 1) is the header,
 * row 3 (index 2) is metadata/blank, and real data begins on row 4 (index 3).
 */
const INPUT_FILE_HEADER_ROW = 1

/**
 * Load survey file.
 * - Sheet detected automatically (handles date-stamped tab names).
 * - Header pinned to row 2; the row-3 metadata/blank line is always skipped.
 */
function loadSurvey(
  buffer: ArrayBuffer,
  sourceTag: string,
  refLabel: string,
  warnings: string[],
): { rows: Row[]; columns: string[] } {
  const { headerRow, matrix } = locateAndRead(
    buffer,
    SHEET_ANCHORS.survey,
    "Survey",
    warnings,
    INPUT_FILE_HEADER_ROW,
  )
  const { rows: positioned, columns } = buildPositionedRows(matrix, headerRow, refLabel, warnings)
  const rows = positioned.map(({ row, excelRow }) => ({
    ...row,
    [COLS.ECODE]: standardizeEcode(row[COLS.ECODE]),
    _source: sourceTag,
    // Source file + TRUE spreadsheet row number, to locate the record in the audit drill-down.
    _srcFile: refLabel,
    _srcRow: excelRow,
  }))
  return { rows, columns }
}

/**
 * Load Exit Interview file.
 * - Sheet detected automatically via the "Cleaned Ecode" anchor.
 * - Header pinned to row 2; the row-3 metadata/blank line is always skipped.
 * - Primary key "Cleaned Ecode" is renamed to "Ecode" for joining.
 */
function loadExitInterview(buffer: ArrayBuffer, warnings: string[]): { rows: Row[]; columns: string[] } {
  const { headerRow, matrix } = locateAndRead(
    buffer,
    SHEET_ANCHORS.ei,
    "Exit Interview",
    warnings,
    INPUT_FILE_HEADER_ROW,
  )
  const { rows: positioned, columns: rawCols } = buildPositionedRows(matrix, headerRow, "Exit Interview", warnings)

  // Rename "Cleaned Ecode" → "Ecode" IN PLACE (pandas rename preserves order).
  const columns = rawCols.map((c) => (c === COLS.EI_CLEANED_ECODE ? COLS.ECODE : c))

  const rows = positioned.map(({ row, excelRow }) => {
    const cleaned = { ...row }
    if (COLS.EI_CLEANED_ECODE in cleaned) {
      cleaned[COLS.ECODE] = standardizeEcode(cleaned[COLS.EI_CLEANED_ECODE])
      delete cleaned[COLS.EI_CLEANED_ECODE]
    }
    cleaned._source = "exit_interview"
    cleaned._srcFile = "Exit Interview"
    cleaned._srcRow = excelRow
    return cleaned
  })
  return { rows, columns }
}

// =============================================================================
// MAPPING LOADERS
// =============================================================================

function loadDeptMapping(
  buffer: ArrayBuffer,
  warnings: string[],
): Map<string, { dept: string; subFunc: string; func: string }> {
  const { rows } = locateAndRead(buffer, SHEET_ANCHORS.dept, "Dept Hierarchy", warnings)

  const map = new Map<string, { dept: string; subFunc: string; func: string }>()
  for (const r of rows) {
    const code = String(r[COLS.DEPT_CODE] ?? "").trim().toUpperCase()
    if (!code || map.has(code)) continue
    map.set(code, {
      dept: String(r[COLS.DEPT_DEPT] ?? ""),
      subFunc: String(r[COLS.DEPT_SUBFUNC] ?? ""),
      func: String(r[COLS.DEPT_FUNC] ?? ""),
    })
  }
  return map
}

function loadGeoMapping(
  buffer: ArrayBuffer,
  warnings: string[],
): Map<number, { office: string; cluster: string; region: string }> {
  const { rows } = locateAndRead(buffer, SHEET_ANCHORS.geo, "Geo Hierarchy", warnings)

  const map = new Map<number, { office: string; cluster: string; region: string }>()
  for (const r of rows) {
    const rawCode = r[COLS.GEO_CODE]
    const code = rawCode != null ? Number(rawCode) : NaN
    if (isNaN(code) || map.has(code)) continue
    map.set(code, {
      office: String(r[COLS.GEO_OFFICE] ?? ""),
      cluster: String(r[COLS.GEO_CLUSTER] ?? ""),
      region: String(r[COLS.GEO_REGION] ?? ""),
    })
  }
  return map
}

function loadPDMapping(
  csvBuffer: ArrayBuffer | string,
  warnings: string[],
): Map<string, { seniority: string; level: string; position: string }> {
  const { rows } = locateAndRead(csvBuffer, SHEET_ANCHORS.pd, "PD Grade", warnings)

  const map = new Map<string, { seniority: string; level: string; position: string }>()
  for (const r of rows) {
    const code = String(r[COLS.PD_CODE] ?? "").trim().toUpperCase()
    if (!code || map.has(code)) continue
    map.set(code, {
      seniority: String(r[COLS.PD_SENIORITY] ?? ""),
      level: String(r[COLS.PD_LEVEL] ?? ""),
      position: String(r[COLS.PD_POSITION] ?? ""),
    })
  }
  return map
}

// =============================================================================
// LOOKUP AUDIT — tracks how many input rows matched each mapping
// =============================================================================

/** Running tally of lookup outcomes for one mapping (Dept/Geo/PD/Beyond Bain). */
export interface LookupAudit {
  total: number // rows that had a non-blank lookup key
  matched: number // rows whose key was found in the mapping
  blankKey: number // rows with no lookup key at all
  unmatched: Map<string, AuditRowRef[]> // unmatched key -> affected rows
}

function newLookupAudit(): LookupAudit {
  return { total: 0, matched: 0, blankKey: 0, unmatched: new Map() }
}

/**
 * Build a structured locator (row number + source file + Ecode) for a row, so
 * the audit can point the analyst at the exact record to triage. The source
 * file and row number are stamped onto each row during load.
 */
function rowLocator(r: Row): AuditRowRef {
  return {
    ecode: String(r[COLS.ECODE] ?? "").trim(),
    source: String((r as Row)._srcFile ?? "").trim(),
    rowNumber: typeof (r as Row)._srcRow === "number" ? ((r as Row)._srcRow as number) : null,
  }
}

/** Record one lookup attempt against an audit tally, remembering the affected row. */
function recordLookup(audit: LookupAudit, rawKey: unknown, matched: boolean, loc: AuditRowRef): void {
  const key = String(rawKey ?? "").trim()
  if (!key) {
    audit.blankKey++
    return
  }
  audit.total++
  if (matched) audit.matched++
  else {
    const rows = audit.unmatched.get(key) ?? []
    rows.push(loc)
    audit.unmatched.set(key, rows)
  }
}

// =============================================================================
// MAPPING APPLIERS
// =============================================================================

function applyDeptMapping(
  rows: Row[],
  map: Map<string, { dept: string; subFunc: string; func: string }>,
  audit: LookupAudit,
): Row[] {
  return rows.map((r) => {
    const code = String(r[COLS.DEPT_INPUT_KEY] ?? "").trim().toUpperCase()
    const lookup = map.get(code)
    recordLookup(audit, r[COLS.DEPT_INPUT_KEY], !!lookup, rowLocator(r))
    // pandas `key.map(...)` OVERWRITES: a missing code yields null (NaN), it does
    // not fall back to any pre-existing value in the column.
    return {
      ...r,
      Department: lookup?.dept ?? null,
      "Sub-Function": lookup?.subFunc ?? null,
      Function: lookup?.func ?? null,
    }
  })
}

function applyGeoMapping(
  rows: Row[],
  map: Map<number, { office: string; cluster: string; region: string }>,
  audit: LookupAudit,
): Row[] {
  return rows.map((r) => {
    const code = Number(r[COLS.GEO_INPUT_KEY])
    const lookup = isNaN(code) ? undefined : map.get(code)
    recordLookup(audit, r[COLS.GEO_INPUT_KEY], !!lookup, rowLocator(r))
    return {
      ...r,
      ExternalDataReference: lookup?.office ?? null,
      Cluster: lookup?.cluster ?? null,
      Region: lookup?.region ?? null,
    }
  })
}

function applyPDMapping(
  rows: Row[],
  map: Map<string, { seniority: string; level: string; position: string }>,
  audit: LookupAudit,
): Row[] {
  return rows.map((r) => {
    const code = String(r[COLS.PD_CODE] ?? "").trim().toUpperCase()
    const lookup = map.get(code)
    recordLookup(audit, r[COLS.PD_CODE], !!lookup, rowLocator(r))
    return {
      ...r,
      "Mapped Seniority": lookup?.seniority ?? null,
      "Mapped Employee Level": lookup?.level ?? null,
      "Mapped Position": lookup?.position ?? null,
    }
  })
}

/** Audit tallies for the three mapping lookups, shared across survey + EI rows. */
export interface MappingAudits {
  dept: LookupAudit
  geo: LookupAudit
  pd: LookupAudit
}

function applyAllMappings(
  rows: Row[],
  deptMap: Map<string, { dept: string; subFunc: string; func: string }>,
  geoMap: Map<number, { office: string; cluster: string; region: string }>,
  pdMap: Map<string, { seniority: string; level: string; position: string }>,
  audits: MappingAudits,
): Row[] {
  let result = applyDeptMapping(rows, deptMap, audits.dept)
  result = applyGeoMapping(result, geoMap, audits.geo)
  result = applyPDMapping(result, pdMap, audits.pd)
  return result
}

/**
 * Logic #1 — Beyond Bain Function Remap
 * Takes first value before semicolon in Function (ZID4_117).
 * Stores result in new column "Mapped Role Function".
 */
function applyBeyondBainLogic(rows: Row[]): Row[] {
  return rows.map((r) => {
    const raw = String(r[COLS.BB_FUNCTION] ?? "")
    const mapped = raw === "null" || raw === "" ? null : raw.split(";")[0].trim()
    return { ...r, [COLS.BB_MAPPED_FUNCTION]: mapped }
  })
}

// =============================================================================
// OUTER JOIN (Survey + EI on Ecode)
// =============================================================================

/**
 * Outer join survey + EI on Ecode, reproducing the reference pandas merge:
 *   - shared columns (present in both, except the key) are COALESCED with the
 *     survey value winning and the EI value filling only where survey is null;
 *   - EI-only columns are kept under their original names;
 *   - the result has a single fixed column schema = survey columns (in order)
 *     followed by EI-only columns (in order);
 *   - duplicate Ecodes produce a cross product per key (pandas semantics).
 */
function outerJoin(
  surveyRows: Row[],
  surveyCols: string[],
  eiRows: Row[],
  eiCols: string[],
): {
  merged: Row[]
  mergedCols: string[]
  bothCount: number
  surveyOnlyCount: number
  eiOnlyCount: number
} {
  const surveySet = new Set(surveyCols)
  const eiOnly = eiCols.filter((c) => c !== COLS.ECODE && !surveySet.has(c))
  const sharedSet = new Set(eiCols.filter((c) => c !== COLS.ECODE && surveySet.has(c)))
  const mergedCols = [...surveyCols, ...eiOnly]

  // Group rows by standardized Ecode, preserving first-seen order on each side.
  // Blank Ecodes get a UNIQUE sentinel key per row so they never match anything
  // (mirroring pandas, where NaN keys do not join to each other).
  let blankSeq = 0
  const group = (rows: Row[]): { map: Map<string, Row[]>; order: string[] } => {
    const map = new Map<string, Row[]>()
    const order: string[] = []
    for (const r of rows) {
      const raw = String(r[COLS.ECODE] ?? "").trim()
      const k = raw === "" ? `\x00blank${blankSeq++}` : raw
      let list = map.get(k)
      if (!list) {
        list = []
        map.set(k, list)
        order.push(k)
      }
      list.push(r)
    }
    return { map, order }
  }
  const s = group(surveyRows)
  const e = group(eiRows)

  const buildRow = (S?: Row, E?: Row): Row => {
    // Use the real Ecode from whichever side is present (sentinel keys are internal only).
    const ecode = String((S ?? E)?.[COLS.ECODE] ?? "")
    const row: Row = {}
    for (const c of surveyCols) {
      if (c === COLS.ECODE) {
        row[c] = ecode
      } else if (sharedSet.has(c)) {
        // Coalesce: survey value wins; fall back to EI only where survey is null.
        row[c] = S?.[c] ?? E?.[c] ?? null
      } else {
        row[c] = S?.[c] ?? null
      }
    }
    for (const c of eiOnly) row[c] = E?.[c] ?? null
    return row
  }

  const merged: Row[] = []
  let bothCount = 0
  let surveyOnlyCount = 0
  let eiOnlyCount = 0
  const seen = new Set<string>()

  for (const key of s.order) {
    seen.add(key)
    const sList = s.map.get(key)!
    const eList = e.map.get(key)
    if (eList && eList.length) {
      for (const S of sList) for (const E of eList) { merged.push(buildRow(S, E)); bothCount++ }
    } else {
      for (const S of sList) { merged.push(buildRow(S, undefined)); surveyOnlyCount++ }
    }
  }
  for (const key of e.order) {
    if (seen.has(key)) continue
    for (const E of e.map.get(key)!) { merged.push(buildRow(undefined, E)); eiOnlyCount++ }
  }

  return { merged, mergedCols, bothCount, surveyOnlyCount, eiOnlyCount }
}

// =============================================================================
// AUDIT SUMMARY HELPERS
// =============================================================================

const MAX_SAMPLES = 10

/**
 * Turn a raw lookup tally into a human-readable audit check.
 * - error  : the mapping never matched a single row (likely a column/sheet mismatch)
 * - warning: some rows didn't match (codes missing from the mapping file)
 * - ok     : every keyed row matched
 */
function summarizeLookup(label: string, mapSize: number, audit: LookupAudit): AuditCheck {
  const unmatchedRows = [...audit.unmatched.values()].reduce((a, b) => a + b.length, 0)
  // Complete, untruncated list of every unmatched code (most-affected first), for drill-down.
  const issues: AuditIssue[] = [...audit.unmatched.entries()]
    .map(([value, rows]) => ({ value, rowCount: rows.length, rows }))
    .sort((a, b) => b.rowCount - a.rowCount)
  const samples = issues
    .slice(0, MAX_SAMPLES)
    .map(({ value, rows }) => `${value} → ${rows.slice(0, 3).map(formatRowRef).join(", ")}${rows.length > 3 ? "…" : ""}`)
  const issueLabel = "Unmatched code"

  const blankNote = audit.blankKey > 0 ? ` ${fmtInt(audit.blankKey)} row(s) had a blank key.` : ""

  if (mapSize === 0) {
    return {
      label,
      status: "error",
      detail: `The ${label} file loaded 0 lookup rows — check that it contains the expected key column and data.`,
    }
  }
  if (audit.total > 0 && audit.matched === 0) {
    return {
      label,
      status: "error",
      detail:
        `None of the ${fmtInt(audit.total)} keyed rows matched the ${label} file. ` +
        `This usually means the input key column or the mapping key column changed name.${blankNote}`,
      samples,
      issueLabel,
      issues,
    }
  }
  if (audit.unmatched.size > 0) {
    return {
      label,
      status: "warning",
      detail:
        `${fmtInt(audit.matched)} of ${fmtInt(audit.total)} keyed rows matched. ` +
        `${fmtInt(unmatchedRows)} row(s) across ${fmtInt(audit.unmatched.size)} unique code(s) ` +
        `were not found in the ${label} file and need a mapping entry added.${blankNote}`,
      samples,
      issueLabel,
      issues,
    }
  }
  return {
    label,
    status: "ok",
    detail: `All ${fmtInt(audit.matched)} keyed rows matched the ${label} file.${blankNote}`,
  }
}

function fmtInt(n: number): string {
  return n.toLocaleString("en-US")
}

/** Compact human-readable reference to a single impacted row, for chip previews. */
function formatRowRef(loc: AuditRowRef): string {
  const where = loc.source && loc.rowNumber != null ? `${loc.source} row ${loc.rowNumber}` : loc.source
  if (loc.ecode && where) return `${loc.ecode} (${where})`
  return loc.ecode || where || "(unidentified row)"
}

// =============================================================================
// COLUMN-ORDER HELPERS (reproduce pandas concat + assignment ordering)
// =============================================================================

/** The 9 columns created/overwritten by the mapping lookups, in assignment order. */
const MAPPED_COLS = [
  "Department",
  "Sub-Function",
  "Function",
  "ExternalDataReference",
  "Cluster",
  "Region",
  "Mapped Seniority",
  "Mapped Employee Level",
  "Mapped Position",
]

/**
 * Beyond Bain columns kept in the final output, in order (prefixed with BB_).
 * The reference output retains only this subset of the Beyond Bain extract.
 */
const BB_ALLOWLIST = [
  "BB_Company (ZID4_9)",
  "BB_Function (ZID4_117)",
  "BB_Job Level (ZID4_88)",
  "BB_Bain Departure Year (ZID7_163)",
  "BB_Client Priority (ZID11_167)",
  "BB_Company Industry (ZID4_202)",
  "BB_Company Sector (ZID4_203)",
  "BB_Mapped Role Function",
]

/** Union of two ordered column lists: first list, then second-list items not already present. */
function orderedUnion(a: string[], b: string[]): string[] {
  const seen = new Set(a)
  const out = [...a]
  for (const c of b) if (!seen.has(c)) { seen.add(c); out.push(c) }
  return out
}

/** Append any of the 9 mapped columns that are not already present (pandas appends on assignment). */
function ensureMappedCols(cols: string[]): string[] {
  const seen = new Set(cols)
  const out = [...cols]
  for (const m of MAPPED_COLS) if (!seen.has(m)) { seen.add(m); out.push(m) }
  return out
}

// =============================================================================
// MAIN PIPELINE
// =============================================================================

export async function runPipeline(files: PipelineFiles): Promise<PipelineResult> {
  const warnings: string[] = []

  // ── STEP 1: Load mapping files ─────────────────────────────────────────
  const deptMap = loadDeptMapping(files.deptHierarchy, warnings)
  const geoMap = loadGeoMapping(files.geoHierarchy, warnings)
  const pdMap = loadPDMapping(files.pdGrade, warnings)

  // Audit tallies shared across survey + EI rows
  const mappingAudits: MappingAudits = {
    dept: newLookupAudit(),
    geo: newLookupAudit(),
    pd: newLookupAudit(),
  }

  // ── STEP 2: Load and stack surveys ─────────────────────────────────────
  const oldSurvey = loadSurvey(files.oldSurvey, "departure_survey", "Survey 1", warnings)
  const newSurvey = loadSurvey(files.newSurvey, "departure_survey", "Survey 2", warnings)

  const oldSurveyRows = oldSurvey.rows.length
  const newSurveyRows = newSurvey.rows.length

  // Stack survey rows (pandas concat: union of columns in first-seen order).
  const combinedSurvey = [...oldSurvey.rows, ...newSurvey.rows]
  // Column order = old columns, then any new-survey columns not already present.
  const surveyColsBase = orderedUnion(oldSurvey.columns, newSurvey.columns)

  // ── STEP 3: Load exit interview ─────────────────────────────────────────
  const ei = loadExitInterview(files.exitInterview, warnings)
  const combinedEI = ei.rows

  // ── STEP 4: Apply mappings ───────────────────────────────────────────���──
  // Mappings overwrite/create the 9 mapped columns; append any that are new.
  const mappedSurvey = applyAllMappings(combinedSurvey, deptMap, geoMap, pdMap, mappingAudits)
  const mappedEI = applyAllMappings(combinedEI, deptMap, geoMap, pdMap, mappingAudits)
  const surveyCols = ensureMappedCols(surveyColsBase)
  const eiCols = ensureMappedCols(ei.columns)

  // ── STEP 5: Outer join survey + EI ─────────────────────────────────────
  const { merged: population, mergedCols, bothCount, surveyOnlyCount, eiOnlyCount } = outerJoin(
    mappedSurvey,
    surveyCols,
    mappedEI,
    eiCols,
  )

  // ── STEP 6: Enrich with Beyond Bain (left join) ─────────────────────────
  const { rows: bbRaw } = locateAndRead(files.beyondBain, SHEET_ANCHORS.bb, "Beyond Bain", warnings)
  let bbRows: Row[] = bbRaw.map((r) => ({
    ...r,
    [COLS.BB_ECODE]: standardizeEcode(r[COLS.BB_ECODE]),
  }))
  bbRows = applyBeyondBainLogic(bbRows)

  // Index BB by Ecode (BB_-prefixed values restricted to the allowlist for output).
  const bbByEcode = new Map<string, Row>()
  for (const r of bbRows) {
    const key = String(r[COLS.BB_ECODE] ?? "")
    if (key && !bbByEcode.has(key)) bbByEcode.set(key, r)
  }

  let beyondBainMatched = 0
  const bbAudit = newLookupAudit()

  // ── STEP 7: Assemble final rows against a single fixed column schema ─────
  const finalColumns = [...mergedCols, ...BB_ALLOWLIST]
  const finalRows = population.map((r) => {
    const ecode = String(r[COLS.ECODE] ?? "")
    const bbRow = bbByEcode.get(ecode)
    recordLookup(bbAudit, ecode, !!bbRow, { ecode, source: "Merged population", rowNumber: null })
    if (bbRow) beyondBainMatched++
    const out: Row = {}
    for (const c of mergedCols) out[c] = r[c] ?? null
    for (const c of BB_ALLOWLIST) {
      // BB_ prefix → original Beyond Bain column name.
      const src = c.slice(3)
      out[c] = bbRow ? (bbRow[src] ?? null) : null
    }
    return out
  })

  const totalColumns = finalColumns.length

  // ── STEP 8: Assemble the audit report ───────────────────────────────────
  const checks: AuditCheck[] = []

  // Blank join-key checks (rows that can never join because they have no Ecode).
  // The offending rows have no Ecode, so we identify them by source file + row number.
  const surveyBlankLocs = combinedSurvey
    .filter((r) => !String(r[COLS.ECODE] ?? "").trim())
    .map((r) => rowLocator(r))
  const eiBlankLocs = combinedEI
    .filter((r) => !String(r[COLS.ECODE] ?? "").trim())
    .map((r) => rowLocator(r))
  // One issue per affected row; the value marks the problem, the row carries the location.
  const blankIssues = (locs: AuditRowRef[]): AuditIssue[] =>
    locs.map((loc) => ({ value: "(blank Ecode)", rowCount: 1, rows: [loc] }))
  checks.push({
    label: "Survey join key (Ecode)",
    status: surveyBlankLocs.length > 0 ? "warning" : "ok",
    detail:
      surveyBlankLocs.length > 0
        ? `${fmtInt(surveyBlankLocs.length)} of ${fmtInt(combinedSurvey.length)} survey rows have a blank Ecode; they remain in the output as unmatched (survey-only) rows.`
        : `All ${fmtInt(combinedSurvey.length)} survey rows have a valid Ecode.`,
    samples: surveyBlankLocs.slice(0, MAX_SAMPLES).map(formatRowRef),
    issueLabel: surveyBlankLocs.length > 0 ? "Issue" : undefined,
    issues: surveyBlankLocs.length > 0 ? blankIssues(surveyBlankLocs) : undefined,
  })
  checks.push({
    label: "Exit Interview join key (Ecode)",
    status: eiBlankLocs.length > 0 ? "warning" : "ok",
    detail:
      eiBlankLocs.length > 0
        ? `${fmtInt(eiBlankLocs.length)} of ${fmtInt(combinedEI.length)} exit-interview rows have a blank Cleaned Ecode; they remain in the output as unmatched (EI-only) rows.`
        : `All ${fmtInt(combinedEI.length)} exit-interview rows have a valid Ecode.`,
    samples: eiBlankLocs.slice(0, MAX_SAMPLES).map(formatRowRef),
    issueLabel: eiBlankLocs.length > 0 ? "Issue" : undefined,
    issues: eiBlankLocs.length > 0 ? blankIssues(eiBlankLocs) : undefined,
  })

  // Mapping lookup checks
  checks.push(summarizeLookup("Department mapping", deptMap.size, mappingAudits.dept))
  checks.push(summarizeLookup("Geography mapping", geoMap.size, mappingAudits.geo))
  checks.push(summarizeLookup("PD Grade mapping", pdMap.size, mappingAudits.pd))

  // Beyond Bain enrichment check (left join — non-matches are expected, so cap severity at warning)
  const bbIssues: AuditIssue[] = [...bbAudit.unmatched.entries()]
    .map(([value, rows]) => ({ value, rowCount: rows.length, rows }))
    .sort((a, b) => b.rowCount - a.rowCount)
  checks.push({
    label: "Beyond Bain enrichment",
    status: population.length > 0 && beyondBainMatched === 0 ? "warning" : "ok",
    detail:
      population.length > 0 && beyondBainMatched === 0
        ? `No population rows matched a Beyond Bain record — verify the Ecode columns line up.`
        : `${fmtInt(beyondBainMatched)} of ${fmtInt(population.length)} population rows were enriched with Beyond Bain data` +
          (bbIssues.length > 0
            ? `; ${fmtInt(bbIssues.length)} Ecode(s) had no Beyond Bain match.`
            : `.`),
    samples: bbIssues.slice(0, MAX_SAMPLES).map(({ value }) => value),
    issueLabel: bbIssues.length > 0 ? "Unenriched Ecode" : undefined,
    issues: bbIssues.length > 0 ? bbIssues : undefined,
  })

  const audit: AuditReport = {
    checks,
    okCount: checks.filter((c) => c.status === "ok").length,
    warningCount: checks.filter((c) => c.status === "warning").length,
    errorCount: checks.filter((c) => c.status === "error").length,
  }

  return {
    rows: finalRows,
    stats: {
      oldSurveyRows,
      newSurveyRows,
      combinedSurveyRows: combinedSurvey.length,
      exitInterviewRows: combinedEI.length,
      populationRows: population.length,
      bothSourcesCount: bothCount,
      surveyOnlyCount,
      eiOnlyCount,
      beyondBainMatched,
      totalColumns,
    },
    warnings,
    audit,
  }
}

// =============================================================================
// OUTPUT HELPERS
// =============================================================================

/**
 * Mapped output columns. When any of these is blank in the final output it means
 * a lookup didn't resolve — exactly the issue an analyst needs to fix — so we
 * highlight those cells yellow in the exported workbook.
 */
const HIGHLIGHT_COLS = [
  "Department",
  "Sub-Function",
  "Function",
  "ExternalDataReference",
  "Cluster",
  "Region",
  "Mapped Seniority",
  "Mapped Employee Level",
  "Mapped Position",
]

function isBlank(v: unknown): boolean {
  return v === null || v === undefined || String(v).trim() === ""
}

/**
 * Build the cleaned XLSX with ExcelJS so we can apply cell styling (SheetJS's
 * community build cannot write fills). Any blank mapped cell is filled yellow
 * so unresolved lookups are easy to spot and triage in the output.
 */
export async function rowsToXLSX(rows: Row[]): Promise<Buffer> {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet("departure_data_cleaned")

  // Column order = union of keys across all rows, preserving first-seen order.
  const cols: string[] = []
  const seen = new Set<string>()
  for (const r of rows) {
    for (const k of Object.keys(r)) {
      if (!seen.has(k)) {
        seen.add(k)
        cols.push(k)
      }
    }
  }

  ws.columns = cols.map((c) => ({ header: c, key: c }))
  ws.getRow(1).font = { bold: true }
  ws.views = [{ state: "frozen", ySplit: 1 }] // keep header visible while scrolling

  const highlightCols = HIGHLIGHT_COLS.filter((c) => seen.has(c))

  for (const r of rows) {
    const values: Record<string, string | number> = {}
    for (const c of cols) {
      const v = r[c]
      values[c] = isBlank(v) ? "" : (v as string | number)
    }
    const added = ws.addRow(values)
    // Highlight blank mapped cells (unresolved lookups).
    for (const c of highlightCols) {
      if (isBlank(r[c])) {
        added.getCell(c).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFEB3B" }, // yellow
        }
      }
    }
  }

  const buf = await wb.xlsx.writeBuffer()
  return Buffer.from(buf as ArrayBuffer)
}

export function rowsToCSV(rows: Row[]): string {
  const ws = XLSX.utils.json_to_sheet(rows)
  return XLSX.utils.sheet_to_csv(ws)
}
