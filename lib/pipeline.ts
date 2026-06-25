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
): { rows: Row[]; sheetName: string; headerRow: number } {
  const wb = readWorkbook(buffer)
  const allAnchors = [...anchors.required, ...anchors.optional]

  let best = { name: "", headerRow: -1, score: -1, requiredHits: -1 }

  for (const name of wb.SheetNames) {
    const matrix = sheetToMatrix(wb.Sheets[name])
    if (matrix.length === 0) continue
    const { row, score } = detectHeaderRow(matrix, allAnchors)
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
        `[${anchors.required.join(", ")}]. Available sheets: ${wb.SheetNames.join(", ")}.`,
    )
  }

  warnings.push(`${label}: detected sheet "${best.name}", header on row ${best.headerRow + 1}.`)

  const rows: Row[] = XLSX.utils.sheet_to_json(wb.Sheets[best.name], {
    defval: null,
    raw: false,
    range: best.headerRow,
  })
  return { rows, sheetName: best.name, headerRow: best.headerRow }
}

/** Detect a Qualtrics metadata row (the `{"ImportId":...}` row beneath headers). */
function looksLikeQualtricsMeta(r: Row): boolean {
  for (const v of Object.values(r)) {
    const s = String(v ?? "")
    if (s.includes("ImportId") || s.trimStart().startsWith('{"')) return true
  }
  return false
}

/** Drop the leading Qualtrics metadata row only if it is actually present. */
function dropQualtricsMeta(rows: Row[], label: string, warnings: string[]): Row[] {
  if (rows.length > 0 && looksLikeQualtricsMeta(rows[0])) {
    warnings.push(`${label}: dropped Qualtrics metadata row.`)
    return rows.slice(1)
  }
  return rows
}

/**
 * Load survey file.
 * - Sheet + header row detected automatically via the "Ecode" anchor.
 * - The Qualtrics metadata row beneath the header is dropped only if present.
 */
function loadSurvey(buffer: ArrayBuffer, sourceTag: string, warnings: string[]): Row[] {
  const { rows } = locateAndRead(buffer, SHEET_ANCHORS.survey, "Survey", warnings)
  const data = dropQualtricsMeta(rows, "Survey", warnings)
  return data.map((r) => ({
    ...r,
    [COLS.ECODE]: standardizeEcode(r[COLS.ECODE]),
    _source: sourceTag,
  }))
}

/**
 * Load Exit Interview file.
 * - Sheet + header row detected automatically via the "Cleaned Ecode" anchor.
 * - Primary key "Cleaned Ecode" is renamed to "Ecode" for joining.
 */
function loadExitInterview(buffer: ArrayBuffer, warnings: string[]): Row[] {
  const { rows } = locateAndRead(buffer, SHEET_ANCHORS.ei, "Exit Interview", warnings)
  const data = dropQualtricsMeta(rows, "Exit Interview", warnings)

  return data.map((r) => {
    const cleaned = { ...r }
    // Rename "Cleaned Ecode" → "Ecode" for joining
    if (COLS.EI_CLEANED_ECODE in cleaned) {
      cleaned[COLS.ECODE] = standardizeEcode(cleaned[COLS.EI_CLEANED_ECODE])
      delete cleaned[COLS.EI_CLEANED_ECODE]
    }
    cleaned._source = "exit_interview"
    return cleaned
  })
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
// MAPPING APPLIERS
// =============================================================================

function applyDeptMapping(
  rows: Row[],
  map: Map<string, { dept: string; subFunc: string; func: string }>,
): Row[] {
  return rows.map((r) => {
    const code = String(r[COLS.DEPT_INPUT_KEY] ?? "").trim().toUpperCase()
    const lookup = map.get(code)
    return {
      ...r,
      Department: lookup?.dept ?? r["Department"] ?? null,
      "Sub-Function": lookup?.subFunc ?? r["Sub-Function"] ?? null,
      Function: lookup?.func ?? r["Function"] ?? null,
    }
  })
}

function applyGeoMapping(
  rows: Row[],
  map: Map<number, { office: string; cluster: string; region: string }>,
): Row[] {
  return rows.map((r) => {
    const code = Number(r[COLS.GEO_INPUT_KEY])
    const lookup = isNaN(code) ? undefined : map.get(code)
    return {
      ...r,
      ExternalDataReference: lookup?.office ?? r["ExternalDataReference"] ?? null,
      Cluster: lookup?.cluster ?? r["Cluster"] ?? null,
      Region: lookup?.region ?? r["Region"] ?? null,
    }
  })
}

function applyPDMapping(
  rows: Row[],
  map: Map<string, { seniority: string; level: string; position: string }>,
): Row[] {
  return rows.map((r) => {
    const code = String(r[COLS.PD_CODE] ?? "").trim().toUpperCase()
    const lookup = map.get(code)
    return {
      ...r,
      "Mapped Seniority": lookup?.seniority ?? r["Mapped Seniority"] ?? null,
      "Mapped Employee Level": lookup?.level ?? r["Mapped Employee Level"] ?? null,
      "Mapped Position": lookup?.position ?? r["Mapped Position"] ?? null,
    }
  })
}

function applyAllMappings(
  rows: Row[],
  deptMap: Map<string, { dept: string; subFunc: string; func: string }>,
  geoMap: Map<number, { office: string; cluster: string; region: string }>,
  pdMap: Map<string, { seniority: string; level: string; position: string }>,
): Row[] {
  let result = applyDeptMapping(rows, deptMap)
  result = applyGeoMapping(result, geoMap)
  result = applyPDMapping(result, pdMap)
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

function outerJoin(
  surveyRows: Row[],
  eiRows: Row[],
): {
  merged: Row[]
  bothCount: number
  surveyOnlyCount: number
  eiOnlyCount: number
} {
  // Index both by Ecode
  const surveyByEcode = new Map<string, Row>()
  for (const r of surveyRows) {
    const key = String(r[COLS.ECODE] ?? "")
    if (key) surveyByEcode.set(key, r)
  }

  const eiByEcode = new Map<string, Row>()
  for (const r of eiRows) {
    const key = String(r[COLS.ECODE] ?? "")
    if (key) eiByEcode.set(key, r)
  }

  const merged: Row[] = []
  let bothCount = 0
  let surveyOnlyCount = 0
  let eiOnlyCount = 0

  // All ecodes across both sources
  const allEcodes = new Set([...surveyByEcode.keys(), ...eiByEcode.keys()])

  for (const ecode of allEcodes) {
    const sRow = surveyByEcode.get(ecode)
    const eRow = eiByEcode.get(ecode)

    if (sRow && eRow) {
      // Both sources — merge, survey wins on conflicts, EI columns prefixed with EI_
      bothCount++
      const merged_row: Row = { ...eRow }
      for (const [k, v] of Object.entries(sRow)) {
        if (k === COLS.ECODE || k === "_source") continue
        if (k in merged_row) {
          // Shared column: survey value wins; EI value stored as EI_<col>
          merged_row[`EI_${k}`] = merged_row[k]
          merged_row[k] = v ?? merged_row[k] // survey value, fall back to EI if null
        } else {
          merged_row[k] = v
        }
      }
      merged_row[COLS.ECODE] = ecode
      merged_row._source_survey = "departure_survey"
      merged_row._source_ei = "exit_interview"
      merged.push(merged_row)
    } else if (sRow) {
      surveyOnlyCount++
      merged.push({ ...sRow, _source_survey: "departure_survey" })
    } else if (eRow) {
      eiOnlyCount++
      merged.push({ ...eRow, _source_ei: "exit_interview" })
    }
  }

  return { merged, bothCount, surveyOnlyCount, eiOnlyCount }
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

  // ── STEP 2: Load and stack surveys ─────────────────────────────────────
  const oldSurvey = loadSurvey(files.oldSurvey, "departure_survey", warnings)
  const newSurvey = loadSurvey(files.newSurvey, "departure_survey", warnings)

  const oldSurveyRows = oldSurvey.length
  const newSurveyRows = newSurvey.length

  // Stack (union of columns, null where absent)
  const allSurveyCols = new Set([
    ...Object.keys(oldSurvey[0] ?? {}),
    ...Object.keys(newSurvey[0] ?? {}),
  ])
  const normalizeRow = (r: Row, cols: Set<string>): Row => {
    const out: Row = {}
    for (const c of cols) out[c] = r[c] ?? null
    return out
  }
  const combinedSurvey = [
    ...oldSurvey.map((r) => normalizeRow(r, allSurveyCols)),
    ...newSurvey.map((r) => normalizeRow(r, allSurveyCols)),
  ]

  // ── STEP 3: Load exit interview ─────────────────────────────────────────
  const combinedEI = loadExitInterview(files.exitInterview, warnings)

  // ── STEP 4: Apply mappings ──────────────────────────────────────────────
  const mappedSurvey = applyAllMappings(combinedSurvey, deptMap, geoMap, pdMap)
  const mappedEI = applyAllMappings(combinedEI, deptMap, geoMap, pdMap)

  // ── STEP 5: Outer join survey + EI ─────────────────────────────────────
  const { merged: population, bothCount, surveyOnlyCount, eiOnlyCount } = outerJoin(
    mappedSurvey,
    mappedEI,
  )

  // ── STEP 6: Enrich with Beyond Bain (left join) ─────────────────────────
  const { rows: bbRaw } = locateAndRead(files.beyondBain, SHEET_ANCHORS.bb, "Beyond Bain", warnings)
  let bbRows: Row[] = bbRaw.map((r) => ({
    ...r,
    [COLS.BB_ECODE]: standardizeEcode(r[COLS.BB_ECODE]),
  }))
  bbRows = applyBeyondBainLogic(bbRows)

  // Index BB by Ecode
  const bbByEcode = new Map<string, Row>()
  for (const r of bbRows) {
    const key = String(r[COLS.BB_ECODE] ?? "")
    if (key) bbByEcode.set(key, r)
  }

  let beyondBainMatched = 0
  const enriched = population.map((r) => {
    const ecode = String(r[COLS.ECODE] ?? "")
    const bbRow = bbByEcode.get(ecode)
    if (!bbRow) return r
    beyondBainMatched++
    // Prefix all BB columns with BB_ to avoid collisions
    const bbPrefixed: Row = {}
    for (const [k, v] of Object.entries(bbRow)) {
      if (k === COLS.BB_ECODE) continue
      bbPrefixed[`BB_${k}`] = v
    }
    return { ...r, ...bbPrefixed }
  })

  // ── STEP 7: Drop internal pipeline columns ──────────────────────────────
  const finalRows = enriched.map((r) => {
    const { _source, ...rest } = r as Row & { _source?: unknown }
    return rest
  })

  const totalColumns = finalRows.length > 0 ? Object.keys(finalRows[0]).length : 0

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
  }
}

// =============================================================================
// OUTPUT HELPERS
// =============================================================================

export function rowsToXLSX(rows: Row[]): Buffer {
  const ws = XLSX.utils.json_to_sheet(rows)
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, "departure_data_cleaned")
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" }) as Buffer
}

export function rowsToCSV(rows: Row[]): string {
  const ws = XLSX.utils.json_to_sheet(rows)
  return XLSX.utils.sheet_to_csv(ws)
}
