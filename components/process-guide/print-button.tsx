"use client"

import { Printer } from "lucide-react"

export function PrintButton() {
  return (
    <button className="guide-print-button" type="button" onClick={() => window.print()}>
      <Printer aria-hidden="true" size={16} />
      Export / Save as PDF
    </button>
  )
}
