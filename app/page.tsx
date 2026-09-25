"use client"

import Link from "next/link"
import { useState, useCallback, type CSSProperties } from "react"
import { Upload, FileSpreadsheet, CheckCircle, AlertCircle, AlertTriangle, XCircle, Download, Loader2, ChevronDown, ChevronUp, ChevronRight, X, BookOpenText } from "lucide-react"

// ── Types ──────────────────────────────────────────────────────────────────
interface FileSlot {
  key: string
  label: string
  description: string
  accept: string
  required: boolean
}

interface PipelineStats {
  oldSurveyRows: number
  newSurveyRows: number
  combinedSurveyRows: number
  exitInterviewRows: number
  populationRows: number
  bothSourcesCount: number
  surveyOnlyCount: number
  eiOnlyCount: number
  beyondBainMatched: number
  exactDuplicateRowsRemoved: number
  duplicateResponseIdGroups: number
  duplicateResponseIdRows: number
  duplicateEcodeGroups: number
  duplicateEcodeRows: number
  totalColumns: number
}

interface AuditRowRef {
  ecode: string
  source: string
  rowNumber: number | null
}

interface AuditIssue {
  value: string
  rowCount: number
  rows: AuditRowRef[]
}

interface AuditCheck {
  label: string
  status: "ok" | "warning" | "error"
  detail: string
  samples?: string[]
  issueLabel?: string
  issues?: AuditIssue[]
}

interface AuditReport {
  checks: AuditCheck[]
  okCount: number
  warningCount: number
  errorCount: number
}

interface PipelineResponse {
  success: boolean
  stats: PipelineStats
  warnings: string[]
  audit: AuditReport
  xlsx: string
  ecodeMap: string
  xlsxRedacted: string
  csvRedacted: string
  error?: string
}

// ── File slot config ────────────────────────────────────────────────────────
const FILE_SLOTS: FileSlot[] = [
  {
    key: "oldSurvey",
    label: "Old Survey",
    description: "Old_Survey_MMDDYY.xlsx — frozen, header row 2",
    accept: ".xlsx",
    required: true,
  },
  {
    key: "newSurvey",
    label: "New Survey",
    description: "New_Survey_MMDDYY.xlsx — active, header row 2",
    accept: ".xlsx",
    required: true,
  },
  {
    key: "exitInterview",
    label: "Exit Interview",
    description: "Exit_Interview_MMDDYY.xlsx — header row 2, Cleaned Ecode key",
    accept: ".xlsx",
    required: true,
  },
  {
    key: "beyondBain",
    label: "Beyond Bain Extract",
    description: "Beyond_Bain_Extract_MMDDYY.xlsx — enrichment only, header row 1",
    accept: ".xlsx",
    required: true,
  },
  {
    key: "deptHierarchy",
    label: "Department Hierarchy",
    description: "dept_hierarchy_YYYYMMDD.xlsx — dept_code lookup",
    accept: ".xlsx",
    required: true,
  },
  {
    key: "geoHierarchy",
    label: "Geographic Hierarchy",
    description: "geographic_hierarchy_YYYYMMDD.xlsx — office_code lookup",
    accept: ".xlsx",
    required: true,
  },
  {
    key: "pdGrade",
    label: "PD Grade Mapping",
    description: "PD_Grade_Mapping_YYYYMMDD.csv — PD Grade lookup",
    accept: ".csv,.xlsx",
    required: true,
  },
]

// ── Helpers ─────────────────────────────────────────────────────────────────
function base64ToBlob(base64: string, mimeType: string): Blob {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return new Blob([bytes], { type: mimeType })
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement("a")
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}

function fmt(n: number): string {
  return n.toLocaleString()
}

// Builds "departure_data_cleaned_YYYY-MM-DD" using the user's local date.
function exportBaseName(): string {
  const now = new Date()
  const yyyy = now.getFullYear()
  const mm = String(now.getMonth() + 1).padStart(2, "0")
  const dd = String(now.getDate()).padStart(2, "0")
  return `departure_data_cleaned_${yyyy}-${mm}-${dd}`
}

// ── Component ────────────────────────────────────────────────────────────────
export default function PipelinePage() {
  const [files, setFiles] = useState<Record<string, File | null>>({})
  const [status, setStatus] = useState<"idle" | "running" | "done" | "error">("idle")
  const [result, setResult] = useState<PipelineResponse | null>(null)
  const [errorMsg, setErrorMsg] = useState<string>("")
  const [showStats, setShowStats] = useState(true)
  // The audit check the user is drilling into (null = modal closed).
  const [drillCheck, setDrillCheck] = useState<AuditCheck | null>(null)

  const allFilesLoaded = FILE_SLOTS.every(s => !!files[s.key])

  const handleFileChange = useCallback((key: string, file: File | null) => {
    setFiles(prev => ({ ...prev, [key]: file }))
  }, [])

  const handleRun = async () => {
    setStatus("running")
    setResult(null)
    setErrorMsg("")

    try {
      const fd = new FormData()
      for (const slot of FILE_SLOTS) {
        if (files[slot.key]) fd.append(slot.key, files[slot.key]!)
      }

      const res = await fetch("/api/pipeline", { method: "POST", body: fd })
      const data: PipelineResponse = await res.json()

      if (!res.ok || !data.success) {
        setStatus("error")
        setErrorMsg(data.error ?? "Unknown error")
        return
      }

      setResult(data)
      setStatus("done")
    } catch (e: unknown) {
      setStatus("error")
      setErrorMsg(e instanceof Error ? e.message : String(e))
    }
  }

  const handleDownloadXLSX = () => {
    if (!result) return
    const blob = base64ToBlob(result.xlsx, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    downloadBlob(blob, `${exportBaseName()}.xlsx`)
  }

  const handleDownloadMapping = () => {
    if (!result) return
    const blob = base64ToBlob(result.ecodeMap, "text/csv")
    downloadBlob(blob, `${exportBaseName()}_ecode_mapping.csv`)
  }

  const handleDownloadXLSXRedacted = () => {
    if (!result) return
    const blob = base64ToBlob(result.xlsxRedacted, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    downloadBlob(blob, `${exportBaseName()}_no_pii.xlsx`)
  }

  const handleDownloadCSVRedacted = () => {
    if (!result) return
    const blob = base64ToBlob(result.csvRedacted, "text/csv")
    downloadBlob(blob, `${exportBaseName()}_no_pii.csv`)
  }

  return (
    <div style={{ minHeight: "100vh", background: "#f5f5f0", fontFamily: "Arial, Helvetica, sans-serif" }}>
      {/* Header */}
      <div style={{ background: "#1a1a1a", padding: "16px 32px", display: "flex", alignItems: "center", gap: 12 }}>
        <div style={{ width: 6, height: 32, background: "#CC0000", borderRadius: 2 }} />
        <div>
          <div style={{ color: "#fff", fontWeight: 700, fontSize: 18 }}>Departure Data Pipeline</div>
          <div style={{ color: "#aaa", fontSize: 12, marginTop: 2 }}>Data cleaning & merge tool · Bain HR Analytics</div>
        </div>
        <Link
          href="/process-guide"
          style={{ marginLeft: "auto", display: "inline-flex", alignItems: "center", gap: 8, minHeight: 38, padding: "8px 12px", border: "1px solid #555", borderRadius: 5, color: "#fff", fontSize: 13, fontWeight: 700, textDecoration: "none" }}
        >
          <BookOpenText aria-hidden="true" size={16} /> Process guide
        </Link>
      </div>

      <div style={{ maxWidth: 900, margin: "0 auto", padding: "32px 24px" }}>

        {/* Instructions */}
        <div style={{ background: "#fff", borderRadius: 8, padding: "20px 24px", marginBottom: 24, border: "1px solid #e5e5e0" }}>
          <div style={{ fontWeight: 700, fontSize: 15, color: "#1a1a1a", marginBottom: 8 }}>How to use</div>
          <div style={{ fontSize: 13, color: "#555", lineHeight: 1.6 }}>
            Upload all 7 files below (4 input files + 3 mapping files), then click <strong>Run Pipeline</strong>.
            The pipeline will stack the surveys, apply all mapping lookups, merge with the exit interview, enrich with Beyond Bain data,
            and produce a cleaned output file ready for the dashboard.
          </div>
        </div>

        {/* File upload grid */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 24 }}>
          {FILE_SLOTS.map(slot => (
            <FileCard
              key={slot.key}
              slot={slot}
              file={files[slot.key] ?? null}
              onChange={f => handleFileChange(slot.key, f)}
            />
          ))}
        </div>

        {/* Run button */}
        <div style={{ display: "flex", justifyContent: "center", marginBottom: 32 }}>
          <button
            onClick={handleRun}
            disabled={!allFilesLoaded || status === "running"}
            style={{
              background: allFilesLoaded && status !== "running" ? "#CC0000" : "#ccc",
              color: "#fff",
              border: "none",
              borderRadius: 6,
              padding: "14px 48px",
              fontSize: 15,
              fontWeight: 700,
              cursor: allFilesLoaded && status !== "running" ? "pointer" : "not-allowed",
              display: "flex",
              alignItems: "center",
              gap: 10,
              transition: "background 0.2s",
            }}
          >
            {status === "running" ? (
              <><Loader2 size={18} style={{ animation: "spin 1s linear infinite" }} /> Running pipeline…</>
            ) : (
              "Run Pipeline"
            )}
          </button>
        </div>

        {/* Running indicator */}
        {status === "running" && (
          <div style={{ background: "#fff", borderRadius: 8, padding: "24px", textAlign: "center", border: "1px solid #e5e5e0", marginBottom: 24 }}>
            <div style={{ color: "#555", fontSize: 14, marginBottom: 12 }}>Processing files — this may take a minute for large datasets…</div>
            <div style={{ display: "flex", justifyContent: "center", gap: 8 }}>
              {["Loading mappings", "Stacking surveys", "Merging EI", "Enriching BB", "Writing output"].map((step, i) => (
                <div key={i} style={{ background: "#f0f0ec", borderRadius: 4, padding: "4px 10px", fontSize: 11, color: "#888" }}>
                  {step}
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Error */}
        {status === "error" && (
          <div style={{ background: "#FEF2F2", border: "1px solid #FECACA", borderRadius: 8, padding: "16px 20px", marginBottom: 24, display: "flex", gap: 12 }}>
            <AlertCircle size={20} color="#DC2626" style={{ flexShrink: 0, marginTop: 1 }} />
            <div>
              <div style={{ fontWeight: 700, color: "#991B1B", fontSize: 14, marginBottom: 4 }}>Pipeline error</div>
              <div style={{ color: "#B91C1C", fontSize: 13, fontFamily: "monospace" }}>{errorMsg}</div>
            </div>
          </div>
        )}

        {/* Results */}
        {status === "done" && result && (
          <div>
            {/* Warnings */}
            {result.warnings.length > 0 && (
              <div style={{ background: "#FFFBEB", border: "1px solid #FDE68A", borderRadius: 8, padding: "12px 16px", marginBottom: 16 }}>
                {result.warnings.map((w, i) => (
                  <div key={i} style={{ fontSize: 13, color: "#92400E" }}>⚠ {w}</div>
                ))}
              </div>
            )}

            {/* Success banner */}
            <div style={{ background: "#F0FDF4", border: "1px solid #BBF7D0", borderRadius: 8, padding: "16px 20px", marginBottom: 20, display: "flex", gap: 12 }}>
              <CheckCircle size={20} color="#16A34A" style={{ flexShrink: 0, marginTop: 1 }} />
              <div>
                <div style={{ fontWeight: 700, color: "#15803D", fontSize: 14 }}>Pipeline complete</div>
                <div style={{ color: "#166534", fontSize: 13, marginTop: 2 }}>
                  {fmt(result.stats.populationRows)} rows · {fmt(result.stats.totalColumns)} columns
                </div>
              </div>
            </div>

            {/* Stats */}
            <div style={{ background: "#fff", borderRadius: 8, border: "1px solid #e5e5e0", marginBottom: 20 }}>
              <button
                onClick={() => setShowStats(s => !s)}
                style={{ width: "100%", background: "none", border: "none", padding: "14px 20px", display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}
              >
                <span style={{ fontWeight: 700, fontSize: 14, color: "#1a1a1a" }}>Run statistics</span>
                {showStats ? <ChevronUp size={16} color="#888" /> : <ChevronDown size={16} color="#888" />}
              </button>
              {showStats && (
                <div style={{ borderTop: "1px solid #f0f0ec", padding: "16px 20px" }}>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16 }}>
                    <StatTile label="Old Survey rows" value={fmt(result.stats.oldSurveyRows)} />
                    <StatTile label="New Survey rows" value={fmt(result.stats.newSurveyRows)} />
                    <StatTile label="Combined Survey" value={fmt(result.stats.combinedSurveyRows)} />
                    <StatTile label="Exit Interview rows" value={fmt(result.stats.exitInterviewRows)} />
                    <StatTile label="Population (total)" value={fmt(result.stats.populationRows)} highlight />
                    <StatTile label="Total columns" value={fmt(result.stats.totalColumns)} />
                    <StatTile label="Both sources" value={fmt(result.stats.bothSourcesCount)} />
                    <StatTile label="Survey only" value={fmt(result.stats.surveyOnlyCount)} />
                    <StatTile label="EI only" value={fmt(result.stats.eiOnlyCount)} />
                    <StatTile label="Beyond Bain matched" value={fmt(result.stats.beyondBainMatched)} />
                  </div>

                  {/* Audit / error check */}
                  {result.audit && (
                    <div style={{ marginTop: 24, borderTop: "1px solid #f0f0ec", paddingTop: 16 }}>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
                        <span style={{ fontWeight: 700, fontSize: 14, color: "#1a1a1a" }}>Audit &amp; error check</span>
                        <AuditSummaryBadge audit={result.audit} />
                      </div>
                      <div style={{ fontSize: 12, color: "#888", marginBottom: 12, lineHeight: 1.5 }}>
                        Each step below is verified after the merge. Anything marked as a warning or error lists the
                        specific values so they can be corrected in the source files.
                      </div>
                      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                        {result.audit.checks.map((c, i) => (
                          <AuditRow key={i} check={c} onDrillDown={() => setDrillCheck(c)} />
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* Downloads */}
            <div style={{ display: "flex", gap: 12 }}>
              <button
                onClick={handleDownloadXLSX}
                style={{ flex: 1, background: "#1a1a1a", color: "#fff", border: "none", borderRadius: 6, padding: "14px", fontSize: 14, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}
              >
                <Download size={16} /> Download .xlsx
              </button>
            </div>
            {/* PII-redacted downloads */}
            <div style={{ display: "flex", gap: 12, marginTop: 12 }}>
              <button
                onClick={handleDownloadXLSXRedacted}
                style={{ flex: 1, background: "#fff", color: "#1a1a1a", border: "1px solid #1a1a1a", borderRadius: 6, padding: "14px", fontSize: 14, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}
              >
                <Download size={16} /> Download .xlsx (no PII)
              </button>
              <button
                onClick={handleDownloadCSVRedacted}
                style={{ flex: 1, background: "#fff", color: "#1a1a1a", border: "1px solid #1a1a1a", borderRadius: 6, padding: "14px", fontSize: 14, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}
              >
                <Download size={16} /> Download .csv (no PII)
              </button>
            </div>
            <button
              onClick={handleDownloadMapping}
              style={{ marginTop: 12, width: "100%", background: "#fff", color: "#1a1a1a", border: "1px solid #1a1a1a", borderRadius: 6, padding: "14px", fontSize: 14, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}
            >
              <Download size={16} /> Download Ecode mapping (.csv)
            </button>
            <div style={{ marginTop: 8, fontSize: 12, color: "#888", display: "flex", alignItems: "center", gap: 6 }}>
              <span style={{ display: "inline-block", width: 12, height: 12, background: "#FFEB3B", border: "1px solid #e0cf00", borderRadius: 2 }} />
              Cells highlighted yellow in the .xlsx are unresolved lookups (blank mapped values) that need attention. CSV has no highlighting.
            </div>
            <div style={{ marginTop: 6, fontSize: 12, color: "#888" }}>
              The &quot;no PII&quot; files remove 8 columns containing names and email addresses. The Ecode column in every output is anonymized with a unique 5-digit code — use the Ecode mapping file to trace each code back to its original Ecode, and keep that file secure.
            </div>
          </div>
        )}
      </div>

      {/* Drill-down modal: full list of offending values for a single check */}
      {drillCheck && <AuditDrillDownModal check={drillCheck} onClose={() => setDrillCheck(null)} />}

      <style>{`
        @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
      `}</style>
    </div>
  )
}

// ── Sub-components ────────────────────────────────────────────────────────────

function FileCard({ slot, file, onChange }: {
  slot: FileSlot
  file: File | null
  onChange: (f: File | null) => void
}) {
  const loaded = !!file

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault()
    const f = e.dataTransfer.files[0]
    if (f) onChange(f)
  }

  return (
    <div
      onDrop={handleDrop}
      onDragOver={e => e.preventDefault()}
      style={{
        background: loaded ? "#F0FDF4" : "#fff",
        border: `1px solid ${loaded ? "#BBF7D0" : "#e5e5e0"}`,
        borderRadius: 8,
        padding: "16px",
        transition: "all 0.15s",
      }}
    >
      <div style={{ display: "flex", alignItems: "flex-start", gap: 10, marginBottom: 10 }}>
        <FileSpreadsheet size={18} color={loaded ? "#16A34A" : "#CC0000"} style={{ flexShrink: 0, marginTop: 1 }} />
        <div style={{ flex: 1 }}>
          <div style={{ fontWeight: 700, fontSize: 13, color: "#1a1a1a" }}>{slot.label}</div>
          <div style={{ fontSize: 11, color: "#888", marginTop: 2, lineHeight: 1.4 }}>{slot.description}</div>
        </div>
        {loaded && <CheckCircle size={16} color="#16A34A" />}
      </div>

      <label style={{ cursor: "pointer" }}>
        <input
          type="file"
          accept={slot.accept}
          style={{ display: "none" }}
          onChange={e => onChange(e.target.files?.[0] ?? null)}
        />
        <div style={{
          border: `1px dashed ${loaded ? "#86EFAC" : "#ccc"}`,
          borderRadius: 5,
          padding: "8px 12px",
          fontSize: 12,
          color: loaded ? "#15803D" : "#888",
          textAlign: "center",
          background: loaded ? "#DCFCE7" : "#fafaf8",
        }}>
          {loaded ? `✓ ${file!.name}` : "Click to upload or drag & drop"}
        </div>
      </label>
    </div>
  )
}

function StatTile({ label, value, highlight = false }: { label: string; value: string; highlight?: boolean }) {
  return (
    <div style={{
      background: highlight ? "#FFF5F5" : "#fafaf8",
      border: `1px solid ${highlight ? "#FECACA" : "#e5e5e0"}`,
      borderRadius: 6,
      padding: "10px 14px",
    }}>
      <div style={{ fontSize: 11, color: "#888", marginBottom: 4 }}>{label}</div>
      <div style={{ fontSize: 18, fontWeight: 700, color: highlight ? "#CC0000" : "#1a1a1a" }}>{value}</div>
    </div>
  )
}

// Color + icon scheme shared by audit rows and the summary badge.
const AUDIT_STYLE = {
  ok: { color: "#15803D", bg: "#F0FDF4", border: "#BBF7D0", Icon: CheckCircle, word: "OK" },
  warning: { color: "#92400E", bg: "#FFFBEB", border: "#FDE68A", Icon: AlertTriangle, word: "Warning" },
  error: { color: "#991B1B", bg: "#FEF2F2", border: "#FECACA", Icon: XCircle, word: "Error" },
} as const

function AuditSummaryBadge({ audit }: { audit: AuditReport }) {
  const overall: "ok" | "warning" | "error" =
    audit.errorCount > 0 ? "error" : audit.warningCount > 0 ? "warning" : "ok"
  const s = AUDIT_STYLE[overall]
  const label =
    overall === "ok"
      ? "All checks passed"
      : `${audit.errorCount} error${audit.errorCount === 1 ? "" : "s"} · ${audit.warningCount} warning${audit.warningCount === 1 ? "" : "s"}`
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 6,
        background: s.bg,
        border: `1px solid ${s.border}`,
        color: s.color,
        borderRadius: 999,
        padding: "3px 10px",
        fontSize: 12,
        fontWeight: 700,
      }}
    >
      <s.Icon size={13} /> {label}
    </span>
  )
}

function AuditRow({ check, onDrillDown }: { check: AuditCheck; onDrillDown: () => void }) {
  const s = AUDIT_STYLE[check.status]
  const issueCount = check.issues?.length ?? 0
  const hasMore = issueCount > (check.samples?.length ?? 0)
  return (
    <div style={{ background: s.bg, border: `1px solid ${s.border}`, borderRadius: 6, padding: "10px 12px" }}>
      <div style={{ display: "flex", alignItems: "flex-start", gap: 8 }}>
        <s.Icon size={15} color={s.color} style={{ flexShrink: 0, marginTop: 1 }} />
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: s.color }}>{check.label}</div>
          <div style={{ fontSize: 12, color: "#555", marginTop: 2, lineHeight: 1.5 }}>{check.detail}</div>
          {check.samples && check.samples.length > 0 && (
            <div style={{ marginTop: 6, display: "flex", flexWrap: "wrap", gap: 4, alignItems: "center" }}>
              {check.samples.map((sample, i) => (
                <span
                  key={i}
                  style={{
                    fontFamily: "monospace",
                    fontSize: 11,
                    background: "#fff",
                    border: `1px solid ${s.border}`,
                    color: s.color,
                    borderRadius: 4,
                    padding: "1px 6px",
                  }}
                >
                  {sample}
                </span>
              ))}
              {hasMore && (
                <span style={{ fontSize: 11, color: "#888" }}>
                  +{fmt(issueCount - (check.samples?.length ?? 0))} more
                </span>
              )}
            </div>
          )}
          {issueCount > 0 && (
            <button
              onClick={onDrillDown}
              style={{
                marginTop: 8,
                background: "#fff",
                border: `1px solid ${s.border}`,
                color: s.color,
                borderRadius: 5,
                padding: "4px 10px",
                fontSize: 12,
                fontWeight: 700,
                cursor: "pointer",
                display: "inline-flex",
                alignItems: "center",
                gap: 6,
              }}
            >
              View all {fmt(issueCount)} {issueCount === 1 ? "issue" : "issues"} <ChevronRight size={13} />
            </button>
          )}
        </div>
      </div>
    </div>
  )
}

function AuditDrillDownModal({ check, onClose }: { check: AuditCheck; onClose: () => void }) {
  const s = AUDIT_STYLE[check.status]
  const issues = check.issues ?? []
  const valueHeader = check.issueLabel ?? "Value"
  const totalRows = issues.reduce((sum, it) => sum + it.rowCount, 0)

  // Flatten to one entry per impacted record so each row can be triaged individually.
  const flatRows = issues.flatMap((it) => it.rows.map((loc) => ({ value: it.value, loc })))

  const thStyle: CSSProperties = {
    textAlign: "left",
    padding: "8px 20px",
    fontSize: 11,
    color: "#888",
    fontWeight: 700,
    borderBottom: "1px solid #f0f0ec",
  }

  const handleExportCsv = () => {
    const esc = (v: string) => `"${v.replace(/"/g, '""')}"`
    const header = [esc(valueHeader), esc("Row number"), esc("Source file"), esc("Ecode")].join(",")
    const lines = flatRows.map(({ value, loc }) =>
      [esc(value), esc(loc.rowNumber != null ? String(loc.rowNumber) : ""), esc(loc.source), esc(loc.ecode)].join(","),
    )
    const csv = [header, ...lines].join("\r\n")
    const blob = new Blob([csv], { type: "text/csv" })
    const safe = check.label.toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "")
    downloadBlob(blob, `audit_${safe}_${exportBaseName()}.csv`)
  }

  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(0,0,0,0.45)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 24,
        zIndex: 50,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: "#fff",
          borderRadius: 10,
          width: "100%",
          maxWidth: 560,
          maxHeight: "80vh",
          display: "flex",
          flexDirection: "column",
          overflow: "hidden",
          boxShadow: "0 20px 60px rgba(0,0,0,0.3)",
        }}
      >
        {/* Modal header */}
        <div style={{ padding: "16px 20px", borderBottom: "1px solid #f0f0ec", display: "flex", alignItems: "flex-start", gap: 10 }}>
          <s.Icon size={18} color={s.color} style={{ flexShrink: 0, marginTop: 2 }} />
          <div style={{ flex: 1 }}>
            <div style={{ fontWeight: 700, fontSize: 15, color: "#1a1a1a" }}>{check.label}</div>
            <div style={{ fontSize: 12, color: "#888", marginTop: 2 }}>
              {fmt(issues.length)} unique {issues.length === 1 ? "value" : "values"} · {fmt(totalRows)} affected{" "}
              {totalRows === 1 ? "row" : "rows"}
            </div>
          </div>
          <button
            onClick={onClose}
            aria-label="Close"
            style={{ background: "none", border: "none", cursor: "pointer", color: "#888", padding: 4, lineHeight: 0 }}
          >
            <X size={18} />
          </button>
        </div>

        {/* Detail line */}
        <div style={{ padding: "12px 20px", fontSize: 12, color: "#555", background: s.bg, lineHeight: 1.5 }}>
          {check.detail}
        </div>

        {/* Scrollable issue table — one row per impacted record */}
        <div style={{ overflowY: "auto", flex: 1 }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ position: "sticky", top: 0, background: "#fafaf8" }}>
                <th style={thStyle}>{valueHeader}</th>
                <th style={{ ...thStyle, width: 90, textAlign: "right" }}>Row #</th>
                <th style={{ ...thStyle, width: 130 }}>Source file</th>
                <th style={{ ...thStyle, width: 120 }}>Ecode</th>
              </tr>
            </thead>
            <tbody>
              {flatRows.map(({ value, loc }, i) => (
                <tr key={i} style={{ borderBottom: "1px solid #f5f5f0" }}>
                  <td style={{ padding: "8px 20px", fontFamily: "monospace", color: "#1a1a1a" }}>
                    {value || <span style={{ color: "#aaa", fontStyle: "italic" }}>(blank)</span>}
                  </td>
                  <td style={{ padding: "8px 20px", textAlign: "right", fontFamily: "monospace", color: "#1a1a1a" }}>
                    {loc.rowNumber != null ? loc.rowNumber : <span style={{ color: "#aaa" }}>—</span>}
                  </td>
                  <td style={{ padding: "8px 20px", color: "#555" }}>
                    {loc.source || <span style={{ color: "#aaa" }}>—</span>}
                  </td>
                  <td style={{ padding: "8px 20px", fontFamily: "monospace", color: "#555" }}>
                    {loc.ecode || <span style={{ color: "#aaa" }}>—</span>}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Modal footer */}
        <div style={{ padding: "12px 20px", borderTop: "1px solid #f0f0ec", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span style={{ fontSize: 12, color: "#888" }}>Export this list to fix the source files.</span>
          <button
            onClick={handleExportCsv}
            style={{ background: "#1a1a1a", color: "#fff", border: "none", borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 8 }}
          >
            <Download size={14} /> Export list (.csv)
          </button>
        </div>
      </div>
    </div>
  )
}
