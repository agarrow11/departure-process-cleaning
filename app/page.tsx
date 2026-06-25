"use client"

import { useState, useCallback } from "react"
import { Upload, FileSpreadsheet, CheckCircle, AlertCircle, Download, Loader2, ChevronDown, ChevronUp } from "lucide-react"

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
  totalColumns: number
}

interface PipelineResponse {
  success: boolean
  stats: PipelineStats
  warnings: string[]
  xlsx: string
  csv: string
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

  const handleDownloadCSV = () => {
    if (!result) return
    const blob = base64ToBlob(result.csv, "text/csv")
    downloadBlob(blob, `${exportBaseName()}.csv`)
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
              <button
                onClick={handleDownloadCSV}
                style={{ flex: 1, background: "#fff", color: "#1a1a1a", border: "1px solid #1a1a1a", borderRadius: 6, padding: "14px", fontSize: 14, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}
              >
                <Download size={16} /> Download .csv
              </button>
            </div>
          </div>
        )}
      </div>

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
