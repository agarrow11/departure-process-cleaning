import { type NextRequest, NextResponse } from "next/server"
import { getCachedResult, type PipelineFiles } from "@/lib/result-cache"

export const runtime = "nodejs"

type FileType = keyof PipelineFiles

const CONTENT_TYPES: Record<FileType, string> = {
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  xlsxRedacted: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ecodeMap: "text/csv",
  csvRedacted: "text/csv",
}

export async function GET(request: NextRequest) {
  const { searchParams } = new URL(request.url)
  const id = searchParams.get("id")
  const type = searchParams.get("type") as FileType | null

  if (!id || !type || !(type in CONTENT_TYPES)) {
    return NextResponse.json({ error: "Missing or invalid id/type" }, { status: 400 })
  }

  const files = getCachedResult(id)
  if (!files) {
    return NextResponse.json(
      { error: "This result has expired or the server restarted. Please re-run the pipeline." },
      { status: 404 }
    )
  }

  const buffer = files[type]
  return new NextResponse(buffer, {
    status: 200,
    headers: {
      "Content-Type": CONTENT_TYPES[type],
      "Content-Length": String(buffer.length),
    },
  })
}
