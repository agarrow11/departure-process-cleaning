import { type NextRequest, NextResponse } from "next/server"
import { runPipeline, rowsToXLSX, rowsToCSV, stripPIIColumns, validateEcodeIntegrity, validatePIIRedaction } from "@/lib/pipeline"
import { cacheResult } from "@/lib/result-cache"
import crypto from "node:crypto"

export const runtime = "nodejs" // Required: pipeline uses Node Buffer APIs
export const maxDuration = 300   // 5 min — large files may take time

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData()

    // ── Validate all required files are present ──────────────────────────
    const required = [
      "oldSurvey", "newSurvey", "exitInterview", "beyondBain",
      "deptHierarchy", "geoHierarchy", "pdGrade",
    ]
    for (const key of required) {
      if (!formData.get(key)) {
        return NextResponse.json(
          { error: `Missing required file: ${key}` },
          { status: 400 }
        )
      }
    }

    // ── Read all files into ArrayBuffers ─────────────────────────────────
    const toBuffer = async (key: string): Promise<ArrayBuffer> => {
      const file = formData.get(key) as File
      return await file.arrayBuffer()
    }

    const files = {
      oldSurvey:     await toBuffer("oldSurvey"),
      newSurvey:     await toBuffer("newSurvey"),
      exitInterview: await toBuffer("exitInterview"),
      beyondBain:    await toBuffer("beyondBain"),
      deptHierarchy: await toBuffer("deptHierarchy"),
      geoHierarchy:  await toBuffer("geoHierarchy"),
      pdGrade:       await toBuffer("pdGrade"),
    }

    // ── Run pipeline ──────────────────────────────────────────────────────
    const result = await runPipeline(files)

    // ── Generate XLSX output ──────────────────────────────────────────────
    const xlsxBuffer = await rowsToXLSX(result.rows)
    // Ecode anonymization mapping (traceability): original Ecode ↔ 5-digit code.
    const mappingCsv = rowsToCSV(result.ecodeMap)
    // PII-redacted variants: remove 8 approved columns and retain 7 fixed-schema
    // headers with blank values. Re-run both hard reconciliations at the file-
    // production boundary so future route changes cannot emit invalid files.
    const redactedRows = stripPIIColumns(result.rows)
    validatePIIRedaction(result.rows, redactedRows)
    validateEcodeIntegrity(result.rows, redactedRows, result.ecodeMap)
    const xlsxRedactedBuffer = await rowsToXLSX(redactedRows)
    const csvRedactedString  = rowsToCSV(redactedRows)

    // Cache the generated files server-side and return only a small JSON
    // payload (stats + audit + a resultId). The previous version inlined all
    // four files as base64 directly in this response — with the wider
    // 187-column schema that combined payload grew large enough to be
    // rejected by an intermediary before reaching the browser (the request
    // itself completed here, in ~95s, with a 200 — the failure was in
    // transit). Files are now fetched on demand, one at a time, as raw
    // binary via /api/pipeline/download, which is both smaller (no base64
    // overhead) and never bundles more than one file per response.
    const resultId = crypto.randomUUID()
    cacheResult(resultId, {
      xlsx: Buffer.from(xlsxBuffer),
      ecodeMap: Buffer.from(mappingCsv),
      xlsxRedacted: Buffer.from(xlsxRedactedBuffer),
      csvRedacted: Buffer.from(csvRedactedString),
    })

    return NextResponse.json({
      success: true,
      resultId,
      stats: result.stats,
      warnings: result.warnings,
      audit: result.audit,
    })

  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err)
    console.error("[pipeline] Error:", message)
    return NextResponse.json({ error: message }, { status: 500 })
  }
}
