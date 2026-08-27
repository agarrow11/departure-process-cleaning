import { type NextRequest, NextResponse } from "next/server"
import { runPipeline, rowsToXLSX, rowsToCSV, stripPIIColumns, validateEcodeIntegrity, validatePIIRedaction } from "@/lib/pipeline"

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

    // Encode outputs as base64 to return in JSON alongside stats
    const xlsxBase64 = Buffer.from(xlsxBuffer).toString("base64")
    const mappingBase64 = Buffer.from(mappingCsv).toString("base64")
    const xlsxRedactedBase64 = Buffer.from(xlsxRedactedBuffer).toString("base64")
    const csvRedactedBase64  = Buffer.from(csvRedactedString).toString("base64")

    return NextResponse.json({
      success: true,
      stats: result.stats,
      warnings: result.warnings,
      audit: result.audit,
      xlsx: xlsxBase64,
      ecodeMap: mappingBase64,
      xlsxRedacted: xlsxRedactedBase64,
      csvRedacted: csvRedactedBase64,
    })

  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err)
    console.error("[pipeline] Error:", message)
    return NextResponse.json({ error: message }, { status: 500 })
  }
}
