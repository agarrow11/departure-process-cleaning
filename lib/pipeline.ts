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

  // Survey sheets
  SURVEY_SHEET_OLD: "Final File_Part 1_Prepped for Q",
  SURVEY_SHEET_NEW: "Datasheet+and+Departure+Survey+",

  // Exit Interview
  EI_CLEANED_ECODE: "Cleaned Ecode",
  EI_SHEET: "Exit Interview",

  // Beyond Bain
  BB_SHEET: "Proposed Export",
  BB_ECODE: "E-code (ZID11_48)",
  BB_FUNCTION: "Function (ZID4_117)",
  BB_MAPPED_FUNCTION: "Mapped Role Function", // created by Logic #1

  // Dept Hierarchy mapping
  DEPT_SHEET: "Dept Hierarchy_260206",
  DEPT_CODE: "dept_code",       // lookup key
  DEPT_DEPT: "Department",      // → Department in input files
  DEPT_SUBFUNC: "Sub-Function", // → Sub-Function in input files
  DEPT_FUNC: "People Function", // → Function in input files
  DEPT_INPUT_KEY: "Department_ID", // key field in input files

  // Geo Hierarchy mapping
  GEO_SHEET: "New columns and New ordering",
  GEO_CODE: "office_code",                  // lookup key
  GEO_OFFICE: "people_office_display",      // → ExternalDataReference
  GEO_CLUSTER: "people_cluster_display",    // → Cluster
  GEO_REGION: "people_region_display",      // → Region
  GEO_INPUT_KEY: "Office_Code",             // key field in input files

  // PD Grade mapping
  PD_CODE: "PD Grade",        // lookup key (same name in input files)
  PD_SENIORITY: "Seniority",  // → Mapped Seniority
  PD_LEVEL: "Employee Level", // → Mapped Employee Level
  PD_POSITION: "CPH_3",       // → Mapped Position
}

// =============================================================================
// HELPERS
// =============================================================================

function standardizeEcode(val: unknown): string {
  if (val == null) return ""
  return String(val).trim().toUpperCase()
}

function readSheet(buffer: ArrayBuffer, sheetName: string, headerRow = 0): Row[] {
  const wb = XLSX.read(buffer, { type: "array" })
  const ws = wb.Sheets[sheetName]
  if (!ws) throw new Error(`Sheet "${sheetName}" not found. Available: ${wb.SheetNames.join(", ")}`)
  const rows: Row[] = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false, range: headerRow })
  return rows
}

function readFirstSheet(buffer: ArrayBuffer, headerRow = 0): Row[] {
  const wb = XLSX.read(buffer, { type: "array" })
  const ws = wb.Sheets[wb.SheetNames[0]]
  const rows: Row[] = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false, range: headerRow })
  return rows
}

/**
 * Load survey file.
 * - Headers in row 2 (index 1) → headerRow=1 in sheet_to_json
 * - Row immediately after headers is Qualtrics metadata → drop first data row
 */
function loadSurvey(buffer: ArrayBuffer, sheetName: string, sourceTag: string): Row[] {
  const wb = XLSX.read(buffer, { type: "array" })

  // Find sheet — try exact name first, then case-insensitive partial match
  let ws = wb.Sheets[sheetName]
  if (!ws) {
    const match = wb.SheetNames.find(n =>
      n.toLowerCase().includes(sheetName.toLowerCase().slice(0, 10))
    )
    if (match) ws = wb.Sheets[match]
    else throw new Error(`Survey sheet "${sheetName}" not found. Available: ${wb.SheetNames.join(", ")}`)
  }

  // header=1 → row index 1 (0-based) is the header row
  const rows: Row[] = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false, range: 1 })
  // Drop first row (Qualtrics metadata)
  const data = rows.slice(1)
  // Standardize Ecode and tag source
  return data.map(r => ({
    ...r,
    [COLS.ECODE]: standardizeEcode(r[COLS.ECODE]),
    _source: sourceTag,
  }))
}

/**
 * Load Exit Interview file.
 * - Headers in row 2 → same range=1 trick
 * - Primary key is "Cleaned Ecode" → renamed to "Ecode" for joining
 */
function loadExitInterview(buffer: ArrayBuffer): Row[] {
  const wb = XLSX.read(buffer, { type: "array" })
  // Try named sheet first, fall back to first sheet
  const sheetName = wb.SheetNames.includes(COLS.EI_SHEET)
    ? COLS.EI_SHEET
    : wb.SheetNames[0]
  const ws = wb.Sheets[sheetName]
  const rows: Row[] = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false, range: 1 })
  const data = rows.slice(1) // drop Qualtrics metadata row

  return data.map(r => {
    const cleaned = { ...r }
    // Rename "Cleaned Ecode" → "Ecode" for joining; keep raw value too
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

function loadDeptMapping(buffer: ArrayBuffer): Map<string, { dept: string; subFunc: string; func: string }> {
  const wb = XLSX.read(buffer, { type: "array" })
  // Use named sheet, ignore "Sheet1"
  const sheetName = wb.SheetNames.find(n => n !== "Sheet1") ?? wb.SheetNames[0]
  const ws = wb.Sheets[sheetName]
  const rows: Row[] = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false })

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

function loadGeoMapping(buffer: ArrayBuffer): Map<number, { office: string; cluster: string; region: string }> {
  const wb = XLSX.read(buffer, { type: "array" })
  const sheetName = wb.SheetNames.includes(COLS.GEO_SHEET)
    ? COLS.GEO_SHEET
    : wb.SheetNames[0]
  const ws = wb.Sheets[sheetName]
  // Header is in row 3 (index 2)
  const rows: Row[] = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false, range: 2 })

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

function loadPDMapping(csvBuffer: ArrayBuffer | string): Map<string, { seniority: string; level: string; position: string }> {
  let wb: XLSX.WorkBook
  if (typeof csvBuffer === "string") {
    wb = XLSX.read(csvBuffer, { type: "string" })
  } else {
    wb = XLSX.read(csvBuffer, { type: "array" })
  }
  const ws = wb.Sheets[wb.SheetNames[0]]
  const rows: Row[] = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false })

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

function applyDeptMapping(rows: Row[], map: Map<string, { dept: string; subFunc: string; func: string }>): Row[] {
  return rows.map(r => {
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

function applyGeoMapping(rows: Row[], map: Map<number, { office: string; cluster: string; region: string }>): Row[] {
  return rows.map(r => {
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

function applyPDMapping(rows: Row[], map: Map<string, { seniority: string; level: string; position: string }>): Row[] {
  return rows.map(r => {
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
  pdMap: Map<string, { seniority: string; level: string; position: string }>
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
  return rows.map(r => {
    const raw = String(r[COLS.BB_FUNCTION] ?? "")
    const mapped = raw === "null" || raw === "" ? null : raw.split(";")[0].trim()
    return { ...r, [COLS.BB_MAPPED_FUNCTION]: mapped }
  })
}

// =============================================================================
// OUTER JOIN (Survey + EI on Ecode)
// =============================================================================

function outerJoin(surveyRows: Row[], eiRows: Row[]): {
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
  const deptMap = loadDeptMapping(files.deptHierarchy)
  const geoMap = loadGeoMapping(files.geoHierarchy)
  const pdMap = loadPDMapping(files.pdGrade)

  // ── STEP 2: Load and stack surveys ─────────────────────────────────────
  let oldSurvey = loadSurvey(files.oldSurvey, COLS.SURVEY_SHEET_OLD, "departure_survey")
  let newSurvey = loadSurvey(files.newSurvey, COLS.SURVEY_SHEET_NEW, "departure_survey")

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
    ...oldSurvey.map(r => normalizeRow(r, allSurveyCols)),
    ...newSurvey.map(r => normalizeRow(r, allSurveyCols)),
  ]

  // ── STEP 3: Load exit interview ─────────────────────────────────────────
  let combinedEI = loadExitInterview(files.exitInterview)

  // ── STEP 4: Apply mappings ──────────────────────────────────────────────
  const mappedSurvey = applyAllMappings(combinedSurvey, deptMap, geoMap, pdMap)
  const mappedEI = applyAllMappings(combinedEI, deptMap, geoMap, pdMap)

  // ── STEP 5: Outer join survey + EI ─────────────────────────────────────
  const { merged: population, bothCount, surveyOnlyCount, eiOnlyCount } =
    outerJoin(mappedSurvey, mappedEI)

  // ── STEP 6: Enrich with Beyond Bain (left join) ────────���────────────────
  const bbWb = XLSX.read(files.beyondBain, { type: "array" })
  const bbSheet = bbWb.Sheets[COLS.BB_SHEET] ?? bbWb.Sheets[bbWb.SheetNames[0]]
  let bbRows: Row[] = XLSX.utils.sheet_to_json(bbSheet, { defval: null, raw: false })
  bbRows = bbRows.map(r => ({
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
  const enriched = population.map(r => {
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
  const finalRows = enriched.map(r => {
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
